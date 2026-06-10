<#
.SYNOPSIS
    End-to-end sync of Microsoft 365 Message Center posts into a SharePoint Online list
    using Microsoft Graph. No Power Automate. Runs locally or as an Azure Automation runbook.

.DESCRIPTION
    1. Authenticates to Microsoft Graph (Managed Identity, certificate, or client secret).
    2. Reads Message Center posts from /admin/serviceAnnouncement/messages (with paging).
    3. Ensures the target SharePoint list exists (optionally creates it with the right columns).
    4. Upserts each post into the list keyed by MessageId (MC######):
         - new posts      -> created, ChangeStatus = 'New'
         - changed posts  -> patched, ChangeStatus = 'Updated'   (detected via LastModified)
         - unchanged      -> skipped
       Idempotent: safe to run on any schedule.

    Everything goes through Invoke-MgGraphRequest, so the ONLY module required is
    Microsoft.Graph.Authentication.

.REQUIRED GRAPH APPLICATION PERMISSIONS  (grant admin consent on the app registration)
    ServiceMessage.Read.All   - read Message Center
    Sites.ReadWrite.All       - write to the SharePoint list
                                (least privilege: use Sites.Selected and grant the app
                                 write on just this one site - see notes at bottom)

.PREREQUISITES
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    (In Azure Automation: add the Microsoft.Graph.Authentication module to the Automation account.)

.EXAMPLE  Local run with an app registration + client secret
    ./Sync-MessageCenterToSharePoint.ps1 `
        -TenantId      'contoso.onmicrosoft.com' `
        -ClientId      '00000000-0000-0000-0000-000000000000' `
        -ClientSecret  $env:GRAPH_SECRET `
        -SiteHostname  'contoso.sharepoint.com' `
        -SitePath      '/sites/PlatformArchitecture' `
        -ListName      'Message Center' `
        -CreateListIfMissing

.EXAMPLE  Azure Automation runbook with a (system-assigned) Managed Identity
    ./Sync-MessageCenterToSharePoint.ps1 `
        -UseManagedIdentity `
        -TenantId     'contoso.onmicrosoft.com' `
        -SiteHostname 'contoso.sharepoint.com' `
        -SitePath     '/sites/PlatformArchitecture' `
        -CreateListIfMissing `
        -SinceDays 30 `
        -ServicesFilter 'Microsoft Teams','SharePoint Online','Exchange Online','Microsoft Copilot'
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]   $TenantId,

    # --- auth (choose ONE path) ---
    [string]   $ClientId,
    [string]   $ClientSecret,
    [string]   $CertificateThumbprint,
    [switch]   $UseManagedIdentity,
    [string]   $ManagedIdentityClientId,      # set for a USER-assigned MI

    # --- target SharePoint site/list ---
    [Parameter(Mandatory)] [string]   $SiteHostname,   # e.g. contoso.sharepoint.com
    [Parameter(Mandatory)] [string]   $SitePath,       # e.g. /sites/PlatformArchitecture
    [string]   $ListName = 'Message Center',
    [switch]   $CreateListIfMissing,

    # --- optional filtering ---
    [string[]] $ServicesFilter,                # only keep posts touching these services
    [int]      $SinceDays                      # only posts modified in the last N days
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$GraphBase = 'https://graph.microsoft.com/v1.0'

# --------------------------------------------------------------------------------------
# Auth
# --------------------------------------------------------------------------------------
function Connect-Graph {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    if ($UseManagedIdentity) {
        Write-Output 'Connecting to Graph with Managed Identity...'
        if ($ManagedIdentityClientId) {
            Connect-MgGraph -Identity -ClientId $ManagedIdentityClientId -NoWelcome
        } else {
            Connect-MgGraph -Identity -NoWelcome
        }
    }
    elseif ($CertificateThumbprint) {
        Write-Output 'Connecting to Graph with certificate (app-only)...'
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId `
            -CertificateThumbprint $CertificateThumbprint -NoWelcome
    }
    elseif ($ClientSecret) {
        Write-Output 'Connecting to Graph with client secret (app-only)...'
        # The Graph SDK has no direct -ClientSecret switch, so fetch a token via OAuth2
        # client_credentials and hand it to Connect-MgGraph as a SecureString.
        $tokenResponse = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body @{
                client_id     = $ClientId
                scope         = 'https://graph.microsoft.com/.default'
                client_secret = $ClientSecret
                grant_type    = 'client_credentials'
            }
        $secureToken = ConvertTo-SecureString $tokenResponse.access_token -AsPlainText -Force
        Connect-MgGraph -AccessToken $secureToken -NoWelcome
    }
    else {
        throw 'Provide one auth method: -UseManagedIdentity, -CertificateThumbprint (with -ClientId), or -ClientSecret (with -ClientId).'
    }
}

# --------------------------------------------------------------------------------------
# Graph helpers
# --------------------------------------------------------------------------------------
# GET that follows @odata.nextLink and returns the flattened .value collection.
function Get-GraphAll {
    param([Parameter(Mandatory)][string]$Uri)
    $all = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        if ($resp.PSObject.Properties.Name -contains 'value') {
            foreach ($v in $resp.value) { $all.Add($v) }
        } elseif ($resp) {
            $all.Add($resp)
        }
        $next = if ($resp.PSObject.Properties.Name -contains '@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
    }
    return $all
}

# Convert an HTML body to a trimmed plain-text summary.
function ConvertTo-PlainText {
    param([string]$Html, [int]$MaxLength = 1500)
    if ([string]::IsNullOrWhiteSpace($Html)) { return '' }
    $t = $Html -replace '(?s)<br\s*/?>', "`n"
    $t = $t -replace '(?s)<[^>]+>', ' '
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    $t = ($t -replace '[ \t]+', ' ') -replace '(\r?\n\s*){2,}', "`n"
    $t = $t.Trim()
    if ($t.Length -gt $MaxLength) { $t = $t.Substring(0, $MaxLength - 3) + '...' }
    return $t
}

# Best-effort extraction of M365 Roadmap feature IDs referenced in a post body.
function Get-RoadmapIds {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $ids = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($m in [regex]::Matches($Text, '(?:Roadmap ID[:\s]*|featureid=|searchterms=)(\d{4,6})', 'IgnoreCase')) {
        [void]$ids.Add($m.Groups[1].Value)
    }
    return @($ids)
}

# --------------------------------------------------------------------------------------
# SharePoint list provisioning
# --------------------------------------------------------------------------------------
function Get-OrCreateList {
    param([Parameter(Mandatory)][string]$SiteId)

    $lists = Get-GraphAll "$GraphBase/sites/$SiteId/lists?`$select=id,displayName,name&`$top=200"
    $list  = $lists | Where-Object { $_.displayName -eq $ListName -or $_.name -eq $ListName } | Select-Object -First 1
    if ($list) {
        Write-Output "Found existing list '$ListName' (id: $($list.id))."
        return $list.id
    }
    if (-not $CreateListIfMissing) {
        throw "List '$ListName' not found on the site. Re-run with -CreateListIfMissing to create it."
    }

    Write-Output "Creating list '$ListName'..."
    $columns = @(
        @{ name = 'MessageId';        indexed = $true; text = @{} },
        @{ name = 'Category';         choice  = @{ choices = @('planForChange','preventOrFixIssue','stayInformed'); displayAs = 'dropDownMenu' } },
        @{ name = 'Severity';         text = @{} },
        @{ name = 'Services';         text = @{ allowMultipleLines = $true } },
        @{ name = 'IsMajorChange';    boolean = @{} },
        @{ name = 'ActionRequiredBy'; dateTime = @{} },
        @{ name = 'MessageStart';     dateTime = @{} },
        @{ name = 'LastModified';     dateTime = @{} },
        @{ name = 'RoadmapIds';       text = @{} },
        @{ name = 'Summary';          text = @{ allowMultipleLines = $true } },
        @{ name = 'WebLink';          text = @{} },
        @{ name = 'ChangeStatus';     text = @{} },
        @{ name = 'SyncedAt';         dateTime = @{} }
    )
    $body = @{
        displayName = $ListName
        list        = @{ template = 'genericList' }
        columns     = $columns
    }
    $created = Invoke-MgGraphRequest -Method POST -Uri "$GraphBase/sites/$SiteId/lists" `
        -Body ($body | ConvertTo-Json -Depth 10) -ContentType 'application/json' -OutputType PSObject
    Write-Output "Created list (id: $($created.id))."
    return $created.id
}

# --------------------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------------------
Connect-Graph

# Resolve the site id. Note ${SiteHostname} braces so the ':' is not parsed as a scope.
Write-Output "Resolving site ${SiteHostname}${SitePath} ..."
$site   = Invoke-MgGraphRequest -Method GET -Uri "$GraphBase/sites/${SiteHostname}:${SitePath}" -OutputType PSObject
$siteId = $site.id
Write-Output "Site id: $siteId"

$listId = Get-OrCreateList -SiteId $siteId

# --- Pull Message Center posts -----------------------------------------------------------
$msgUri = "$GraphBase/admin/serviceAnnouncement/messages"
if ($SinceDays) {
    $since = (Get-Date).AddDays(-$SinceDays).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $msgUri += "?`$filter=lastModifiedDateTime ge $since"
}
Write-Output "Reading Message Center posts..."
$messages = Get-GraphAll $msgUri
Write-Output "Retrieved $($messages.Count) post(s) from Message Center."

if ($ServicesFilter) {
    $messages = $messages | Where-Object {
        $svc = $_.services
        $svc -and ($svc | Where-Object { $ServicesFilter -contains $_ })
    }
    Write-Output "After service filter: $($messages.Count) post(s)."
}

# --- Index existing list items by MessageId ---------------------------------------------
Write-Output "Loading existing list items..."
$existing = @{}
$itemsUri = "$GraphBase/sites/$siteId/lists/$listId/items?expand=fields&`$top=200"
foreach ($item in Get-GraphAll $itemsUri) {
    $mid = $null
    if ($item.fields -and ($item.fields.PSObject.Properties.Name -contains 'MessageId')) {
        $mid = $item.fields.MessageId
    }
    if ($mid) { $existing[$mid] = $item }
}
Write-Output "Existing items in list: $($existing.Count)."

# --- Upsert ------------------------------------------------------------------------------
$new = 0; $updated = 0; $unchanged = 0; $failed = 0
$syncedAt = (Get-Date).ToUniversalTime().ToString('o')

foreach ($m in $messages) {
    try {
        $bodyContent = if ($m.body) { $m.body.content } else { '' }
        $roadmapIds  = Get-RoadmapIds (("$($m.title) $bodyContent"))

        $fields = @{
            Title         = if ($m.title) { $m.title.Substring(0, [Math]::Min(255, $m.title.Length)) } else { $m.id }
            MessageId     = $m.id
            Category      = "$($m.category)"
            Severity      = "$($m.severity)"
            Services      = (@($m.services) -join '; ')
            IsMajorChange = [bool]$m.isMajorChange
            LastModified  = $m.lastModifiedDateTime
            Summary       = (ConvertTo-PlainText $bodyContent)
            WebLink       = "https://admin.microsoft.com/Adminportal/Home#/MessageCenter/:/messages/$($m.id)"
            SyncedAt      = $syncedAt
        }
        if ($m.actionRequiredByDateTime) { $fields.ActionRequiredBy = $m.actionRequiredByDateTime }
        if ($m.startDateTime)            { $fields.MessageStart     = $m.startDateTime }
        if ($roadmapIds.Count -gt 0)     { $fields.RoadmapIds       = ($roadmapIds -join ', ') }

        if ($existing.ContainsKey($m.id)) {
            $cur = $existing[$m.id]
            $changed = $true
            try {
                $changed = ([datetime]$cur.fields.LastModified).ToUniversalTime() -ne `
                           ([datetime]$m.lastModifiedDateTime).ToUniversalTime()
            } catch { $changed = $true }

            if ($changed) {
                $fields.ChangeStatus = 'Updated'
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "$GraphBase/sites/$siteId/lists/$listId/items/$($cur.id)/fields" `
                    -Body ($fields | ConvertTo-Json -Depth 5) -ContentType 'application/json' | Out-Null
                $updated++
                Write-Output "  [UPDATED] $($m.id)  $($fields.Title)"
            } else {
                $unchanged++
            }
        } else {
            $fields.ChangeStatus = 'New'
            Invoke-MgGraphRequest -Method POST `
                -Uri "$GraphBase/sites/$siteId/lists/$listId/items" `
                -Body (@{ fields = $fields } | ConvertTo-Json -Depth 6) -ContentType 'application/json' | Out-Null
            $new++
            Write-Output "  [NEW]     $($m.id)  $($fields.Title)"
        }
    }
    catch {
        $failed++
        Write-Warning "Failed on $($m.id): $($_.Exception.Message)"
    }
}

Write-Output ''
Write-Output '================ SYNC SUMMARY ================'
Write-Output ("  New:       {0}" -f $new)
Write-Output ("  Updated:   {0}" -f $updated)
Write-Output ("  Unchanged: {0}" -f $unchanged)
Write-Output ("  Failed:    {0}" -f $failed)
Write-Output ("  List:      {0}  ({1})" -f $ListName, $listId)
Write-Output '============================================='

Disconnect-MgGraph | Out-Null

<#
-----------------------------------------------------------------------------------------
LEAST-PRIVILEGE WRITE ACCESS (recommended): Sites.Selected
-----------------------------------------------------------------------------------------
Instead of Sites.ReadWrite.All, grant the app write to ONLY this one site:

  1. On the app registration, add application permission: Sites.Selected (admin consent).
  2. As a SharePoint/Global admin, grant the app write on the target site (run once,
     e.g. from a delegated Graph session or PnP):

     POST https://graph.microsoft.com/v1.0/sites/{siteId}/permissions
     {
       "roles": ["write"],
       "grantedToIdentities": [
         { "application": { "id": "<app-client-id>", "displayName": "MC-Sync" } }
       ]
     }

  Keep ServiceMessage.Read.All as a tenant-wide application permission (MC is tenant scope).

-----------------------------------------------------------------------------------------
SCHEDULING
-----------------------------------------------------------------------------------------
  Azure Automation:  import Microsoft.Graph.Authentication into the Automation account,
                     enable a Managed Identity, grant it the two permissions above, then
                     attach a daily schedule to this runbook (-UseManagedIdentity).
  Windows/Task Sched / cron+pwsh:  use -CertificateThumbprint or -ClientSecret.
#>
