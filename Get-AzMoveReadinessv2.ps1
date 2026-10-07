#Requires -Version 5.1
<#
.SYNOPSIS
    Read-only move-readiness assessment for consolidating Azure resources from several
    subscriptions into one target subscription (e.g. a Crayon CSP subscription).

.DESCRIPTION
    For every resource in the source subscriptions, the script reports:
      - Move readiness (Easy / Review / Blocked / Skip-Recreate) with the reasons
      - Move group: resources hard-linked to each other that Azure requires to move together
      - Dependencies (hard links), soft references, and what references the resource
      - Last control-plane change (Activity Log, max 90 days) and who made it
      - Optional: last data-plane activity from a usage metric (the closest thing Azure
        has to "last accessed"), 30-day cost, and Microsoft's own move validation

    NOTHING IS MOVED OR CHANGED. -ValidateMove calls Azure's validateMoveResources API,
    which is a dry run.

    Output: CSV files (plus an .xlsx if the ImportExcel module is installed).

.PARAMETER TargetSubscriptionId
    The subscription everything should end up in.

.PARAMETER SourceSubscriptionIds
    Subscriptions to assess. Default: every enabled subscription you can see except the target.

.PARAMETER ActivityDays
    Days of Activity Log to scan for last change (1-90). Default 90.

.PARAMETER IncludeUsageMetrics
    Pull one usage metric per supported resource type to find last data-plane activity.
    Adds one API call per supported resource, so it is slower.

.PARAMETER UsageDays
    Days of metrics to look at (1-93). Default 30.

.PARAMETER IncludeCost
    Pull actual cost per resource for the last 30 days from Cost Management.
    Some credit/sponsorship offers don't support the Cost Management API; those show "Unavailable".

.PARAMETER ValidateMove
    Ask Azure to validate each move group against -TargetResourceGroupName (dry run).
    Requires Contributor (moveResources/action) on source RGs and write on the target RG.
    Each validation takes ~10-60 seconds.

.PARAMETER TargetResourceGroupName
    Existing resource group in the target subscription used for validation.

.EXAMPLE
    .\Get-AzMoveReadiness.ps1 -TargetSubscriptionId 00000000-0000-0000-0000-000000000000

.EXAMPLE
    .\Get-AzMoveReadiness.ps1 -TargetSubscriptionId <target> -SourceSubscriptionIds <sub1>,<sub2> `
        -IncludeUsageMetrics -IncludeCost -ValidateMove -TargetResourceGroupName rg-move-validation
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TargetSubscriptionId,
    [string[]]$SourceSubscriptionIds,
    [ValidateRange(1, 90)][int]$ActivityDays = 90,
    [switch]$IncludeUsageMetrics,
    [ValidateRange(1, 93)][int]$UsageDays = 30,
    [switch]$IncludeCost,
    [switch]$ValidateMove,
    [string]$TargetResourceGroupName,
    [string]$OutputFolder = (Join-Path (Get-Location) ("AzMoveAssessment_{0:yyyyMMdd_HHmm}" -f (Get-Date)))
)

$ErrorActionPreference = 'Stop'
$WarningPreference     = 'SilentlyContinue'   # suppress Az breaking-change banners

#region ---------- Reference data (edit these to tune the assessment) ----------

# Move-support heuristics for cross-subscription moves. Microsoft's validateMoveResources
# (-ValidateMove) is the authority; this table just gives an early read without it.
# Source: learn.microsoft.com/azure/azure-resource-manager/management/move-support-resources
$MoveSupport = @{
    'microsoft.aad/domainservices'                 = @{ Status = 'NotSupported'; Note = 'Entra Domain Services cannot be moved.' }
    'microsoft.containerservice/managedclusters'   = @{ Status = 'NotSupported'; Note = 'AKS cannot be moved; redeploy in target.' }
    'microsoft.databricks/workspaces'              = @{ Status = 'NotSupported'; Note = 'Databricks workspaces cannot be moved; redeploy.' }
    'microsoft.network/applicationgateways'        = @{ Status = 'NotSupported'; Note = 'Application Gateway cannot be moved; redeploy.' }
    'microsoft.network/azurefirewalls'             = @{ Status = 'NotSupported'; Note = 'Azure Firewall cannot be moved; redeploy.' }
    'microsoft.network/bastionhosts'               = @{ Status = 'NotSupported'; Note = 'Bastion cannot be moved; delete and recreate.' }
    'microsoft.network/frontdoors'                 = @{ Status = 'NotSupported'; Note = 'Front Door (classic) cannot be moved.' }
    'microsoft.web/hostingenvironments'            = @{ Status = 'NotSupported'; Note = 'App Service Environment cannot be moved.' }
    'microsoft.classiccompute/*'                   = @{ Status = 'NotSupported'; Note = 'Classic (ASM) resource; migrate to ARM instead.' }
    'microsoft.classicnetwork/*'                   = @{ Status = 'NotSupported'; Note = 'Classic (ASM) resource; migrate to ARM instead.' }
    'microsoft.classicstorage/*'                   = @{ Status = 'NotSupported'; Note = 'Classic (ASM) resource; migrate to ARM instead.' }

    'microsoft.compute/virtualmachines'            = @{ Status = 'Restricted'; Note = 'Move with NICs/disks/VNet. Marketplace-plan images need terms accepted in target. Azure Backup restore points must be removed. ADE needs its Key Vault moved too.' }
    'microsoft.compute/disks'                      = @{ Status = 'Restricted'; Note = 'Move with its VM. Disks using a Disk Encryption Set need the DES moved too.' }
    'microsoft.compute/virtualmachinescalesets'    = @{ Status = 'Restricted'; Note = 'Check load balancer / public IP SKU constraints.' }
    'microsoft.compute/snapshots'                  = @{ Status = 'Restricted'; Note = 'Check for incremental snapshots / encryption constraints.' }
    'microsoft.network/virtualnetworks'            = @{ Status = 'Restricted'; Note = 'All dependent resources (NICs, VMs, gateways, PEs) must move together. Remove VNet peerings first.' }
    'microsoft.network/networkinterfaces'          = @{ Status = 'Restricted'; Note = 'Moves with its VM and VNet.' }
    'microsoft.network/publicipaddresses'          = @{ Status = 'Restricted'; Note = 'Move with the attached resource; SKU constraints apply. IP address may not be retained.' }
    'microsoft.network/loadbalancers'              = @{ Status = 'Restricted'; Note = 'SKU constraints apply; move with backend resources.' }
    'microsoft.network/virtualnetworkgateways'     = @{ Status = 'Restricted'; Note = 'Gateway move has SKU/connection constraints; often easier to recreate.' }
    'microsoft.network/natgateways'                = @{ Status = 'Restricted'; Note = 'Check current move support; may need recreation.' }
    'microsoft.network/privateendpoints'           = @{ Status = 'Restricted'; Note = 'Move with its VNet; target resource connection must still resolve.' }
    'microsoft.web/sites'                          = @{ Status = 'Restricted'; Note = 'Move all App Service resources in the RG together (plan, apps, certs). Target RG must not already contain App Service resources.' }
    'microsoft.web/serverfarms'                    = @{ Status = 'Restricted'; Note = 'Move with all its apps and certificates.' }
    'microsoft.web/certificates'                   = @{ Status = 'Restricted'; Note = 'Moves with the App Service apps that use it.' }
    'microsoft.recoveryservices/vaults'            = @{ Status = 'Restricted'; Note = 'Backup constraints; often simpler to stop protection and protect from a new vault in target.' }
    'microsoft.keyvault/vaults'                    = @{ Status = 'Restricted'; Note = 'Same tenant only. Check ADE/CMK consumers and anything referencing the vault by resource ID.' }
    'microsoft.sql/servers'                        = @{ Status = 'Restricted'; Note = 'Databases and elastic pools move with the server.' }
    'microsoft.sql/managedinstances'               = @{ Status = 'Restricted'; Note = 'Managed Instance has move constraints; check docs.' }
    'microsoft.operationalinsights/workspaces'     = @{ Status = 'Restricted'; Note = 'Unlink Automation account / remove linked solutions first.' }
    'microsoft.automation/automationaccounts'      = @{ Status = 'Restricted'; Note = 'Unlink from Log Analytics first; re-check identities.' }
    'microsoft.logic/workflows'                    = @{ Status = 'Restricted'; Note = 'API connections are separate resources; move them too.' }
    'microsoft.machinelearningservices/workspaces' = @{ Status = 'Restricted'; Note = 'Check current move support; dependent resources must move together.' }
    'microsoft.synapse/workspaces'                 = @{ Status = 'Restricted'; Note = 'Check current move support; may need redeploy.' }
    'microsoft.desktopvirtualization/hostpools'    = @{ Status = 'Restricted'; Note = 'AVD objects have move limits; check docs.' }
}

# A reference TO one of these types is "soft": the referencing resource can move without it.
# Soft references don't merge move groups. Edit to taste.
$SoftTargetTypes = @(
    'microsoft.operationalinsights/workspaces', 'microsoft.insights/components',
    'microsoft.insights/actiongroups', 'microsoft.keyvault/vaults',
    'microsoft.storage/storageaccounts', 'microsoft.managedidentity/userassignedidentities',
    'microsoft.network/networkwatchers', 'microsoft.network/privatednszones',
    'microsoft.network/dnszones', 'microsoft.compute/galleries', 'microsoft.compute/images',
    'microsoft.recoveryservices/vaults'
)
# References FROM these types are soft (alerts etc. point at things but don't bind them).
$SoftSourceTypes = @(
    'microsoft.insights/metricalerts', 'microsoft.insights/scheduledqueryrules',
    'microsoft.insights/activitylogalerts', 'microsoft.insights/diagnosticsettings',
    'microsoft.alertsmanagement/smartdetectoralertrules', 'microsoft.insights/workbooks',
    'microsoft.portal/dashboards', 'microsoft.security/automations'
)
# JSON property names whose resource IDs are provenance, not dependencies.
$IgnoreRefKeys = @('imageReference', 'creationData', 'galleryImageReference', 'sourceResourceId', 'sourceUri', 'hiddenLink')

# Resource groups / types Azure creates automatically. Usually skipped and recreated in target.
$AutoCreatedRgPatterns = @('^NetworkWatcherRG$', '^DefaultResourceGroup-', '^cloud-shell-storage-', '^MC_', '^databricks-rg-', '^AzureBackupRG_', '^LogAnalyticsDefaultResources$')
$AutoCreatedTypes      = @('microsoft.network/networkwatchers', 'microsoft.network/networkwatchers/flowlogs')

# One usage metric per type, used to estimate last data-plane activity ("last accessed").
$UsageMetricMap = @{
    'microsoft.compute/virtualmachines'         = @{ Metric = 'Network In Total';      Agg = 'Total' }
    'microsoft.storage/storageaccounts'         = @{ Metric = 'Transactions';          Agg = 'Total' }
    'microsoft.web/sites'                       = @{ Metric = 'Requests';              Agg = 'Total' }
    'microsoft.sql/servers/databases'           = @{ Metric = 'connection_successful'; Agg = 'Total' }
    'microsoft.keyvault/vaults'                 = @{ Metric = 'ServiceApiHit';         Agg = 'Total' }
    'microsoft.documentdb/databaseaccounts'     = @{ Metric = 'TotalRequests';         Agg = 'Count' }
    'microsoft.cache/redis'                     = @{ Metric = 'connectedclients';      Agg = 'Maximum' }
    'microsoft.logic/workflows'                 = @{ Metric = 'RunsStarted';           Agg = 'Total' }
    'microsoft.servicebus/namespaces'           = @{ Metric = 'IncomingMessages';      Agg = 'Total' }
    'microsoft.eventhub/namespaces'             = @{ Metric = 'IncomingMessages';      Agg = 'Total' }
    'microsoft.cognitiveservices/accounts'      = @{ Metric = 'TotalCalls';            Agg = 'Total' }
    'microsoft.apimanagement/service'           = @{ Metric = 'Requests';              Agg = 'Total' }
    'microsoft.dbforpostgresql/flexibleservers' = @{ Metric = 'active_connections';    Agg = 'Maximum' }
    'microsoft.dbformysql/flexibleservers'      = @{ Metric = 'active_connections';    Agg = 'Maximum' }
    'microsoft.containerregistry/registries'    = @{ Metric = 'TotalPullCount';        Agg = 'Total' }
}
#endregion

#region ---------- Helpers ----------
function Write-Step([string]$Message) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message) -ForegroundColor Cyan }
function Write-Note([string]$Message) { Write-Host ("           {0}" -f $Message) -ForegroundColor DarkYellow }

$ArmIdRegex = [regex]'(?i)/subscriptions/[0-9a-f\-]{36}/resourcegroups/[^/"''\s]+/providers/[^/"''\s]+/[^/"''\s]+/[^/"''\s?#]+'

function Get-TopLevelId([string]$Id) {
    # Collapse any ARM ID (including child resources like .../subnets/x) to its top-level resource.
    if (-not $Id) { return $null }
    $m = $ArmIdRegex.Match($Id)
    if ($m.Success) { return $m.Value.ToLowerInvariant() }
    return $null
}

function Get-IdParts([string]$Id) {
    $m = [regex]::Match($Id, '(?i)/subscriptions/([^/]+)/resourcegroups/([^/]+)/providers/([^/]+/[^/]+)/([^/]+)')
    if ($m.Success) {
        return [pscustomobject]@{ Sub = $m.Groups[1].Value.ToLowerInvariant(); Rg = $m.Groups[2].Value; Type = $m.Groups[3].Value.ToLowerInvariant(); Name = $m.Groups[4].Value }
    }
    return $null
}

function Find-ArmReferences {
    # Walk a resource's properties and collect every ARM resource ID it mentions.
    param($Node, [System.Collections.Generic.HashSet[string]]$Found, [int]$Depth = 0)
    if ($null -eq $Node -or $Depth -gt 40) { return }
    if ($Node -is [string]) {
        foreach ($m in $ArmIdRegex.Matches($Node)) { [void]$Found.Add($m.Value.ToLowerInvariant()) }
        return
    }
    if ($Node -is [System.Collections.IDictionary]) {
        foreach ($k in $Node.Keys) { if ($IgnoreRefKeys -notcontains $k) { Find-ArmReferences -Node $Node[$k] -Found $Found -Depth ($Depth + 1) } }
        return
    }
    if ($Node -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in $Node.PSObject.Properties) { if ($IgnoreRefKeys -notcontains $p.Name) { Find-ArmReferences -Node $p.Value -Found $Found -Depth ($Depth + 1) } }
        return
    }
    if ($Node -is [System.Collections.IEnumerable]) {
        foreach ($i in $Node) { Find-ArmReferences -Node $i -Found $Found -Depth ($Depth + 1) }
    }
}

function Invoke-ArgQuery([string]$Query, [string[]]$Subscriptions) {
    $results = New-Object System.Collections.Generic.List[object]
    $skip = $null
    do {
        $p = @{ Query = $Query; First = 1000; Subscription = $Subscriptions }
        if ($skip) { $p.SkipToken = $skip }
        $r = Search-AzGraph @p
        if ($r.PSObject.Properties['Data']) {
            foreach ($row in $r.Data) { $results.Add($row) }
            $skip = $r.SkipToken
        } else {
            foreach ($row in $r) { $results.Add($row) }   # older Az.ResourceGraph versions
            $skip = $null
        }
    } while ($skip)
    return $results
}

function Get-ErrorText([string]$Content) {
    try {
        $j = $Content | ConvertFrom-Json
        $msgs = @($j.error.message)
        if ($j.error.details) { $msgs += @($j.error.details | ForEach-Object { $_.message }) }
        $text = (@($msgs | Where-Object { $_ }) -join ' | ')
    } catch { $text = $Content }
    if ($text.Length -gt 2000) { $text = $text.Substring(0, 2000) + '...' }
    return $text
}

function Get-ActivityValue($Obj, [string]$Name) {
    # Handles both old (Category.Value) and new (CategoryValue) Az.Monitor output shapes.
    $flat = $Obj.PSObject.Properties["${Name}Value"]
    if ($flat -and $flat.Value) { return [string]$flat.Value }
    $nested = $Obj.$Name
    if ($nested -and $nested.PSObject.Properties['Value']) { return [string]$nested.Value }
    return [string]$nested
}

$script:UfParent = @{}
function Find-Root([string]$X) {
    while ($script:UfParent[$X] -ne $X) { $script:UfParent[$X] = $script:UfParent[$script:UfParent[$X]]; $X = $script:UfParent[$X] }
    return $X
}
function Join-Set([string]$A, [string]$B) {
    $ra = Find-Root $A; $rb = Find-Root $B
    if ($ra -ne $rb) { $script:UfParent[$ra] = $rb }
}

function Test-MoveGroup([string]$SubId, [string]$Rg, [string[]]$Ids, [string]$TargetRgId) {
    $payload = @{ resources = @($Ids); targetResourceGroup = $TargetRgId } | ConvertTo-Json -Depth 5
    $resp = Invoke-AzRestMethod -Method POST -Path "/subscriptions/$SubId/resourceGroups/$Rg/validateMoveResources?api-version=2021-04-01" -Payload $payload
    if ($resp.StatusCode -eq 204) { return @{ Status = 'Passed'; Message = '' } }
    if ($resp.StatusCode -ne 202) { return @{ Status = 'Failed'; Message = (Get-ErrorText $resp.Content) } }
    $loc = $resp.Headers.Location.AbsoluteUri
    for ($i = 0; $i -lt 90; $i++) {
        $wait = 10
        if ($resp.Headers.RetryAfter -and $resp.Headers.RetryAfter.Delta) { $wait = [int]$resp.Headers.RetryAfter.Delta.TotalSeconds }
        Start-Sleep -Seconds ([Math]::Max($wait, 5))
        $resp = Invoke-AzRestMethod -Method GET -Uri $loc
        if ($resp.StatusCode -eq 202) { continue }
        if ($resp.StatusCode -in 200, 204) { return @{ Status = 'Passed'; Message = '' } }
        return @{ Status = 'Failed'; Message = (Get-ErrorText $resp.Content) }
    }
    return @{ Status = 'Timeout'; Message = 'Validation still running after ~15 minutes.' }
}

function Get-SubscriptionCost([string]$SubId, [int]$Days) {
    $now  = [datetime]::UtcNow
    $body = @{
        type       = 'ActualCost'
        timeframe  = 'Custom'
        timePeriod = @{ from = $now.Date.AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:ssZ'); to = $now.ToString('yyyy-MM-ddTHH:mm:ssZ') }
        dataset    = @{
            granularity = 'None'
            aggregation = @{ totalCost = @{ name = 'Cost'; function = 'Sum' } }
            grouping    = @(@{ type = 'Dimension'; name = 'ResourceId' })
        }
    } | ConvertTo-Json -Depth 10
    $result = @{}; $currency = $null
    $uri = $null
    do {
        $attempt = 0
        do {
            $attempt++
            if ($uri) { $resp = Invoke-AzRestMethod -Method POST -Uri $uri -Payload $body }
            else      { $resp = Invoke-AzRestMethod -Method POST -Path "/subscriptions/$SubId/providers/Microsoft.CostManagement/query?api-version=2023-03-01" -Payload $body }
            if ($resp.StatusCode -eq 429) { Start-Sleep -Seconds 30 }
        } while ($resp.StatusCode -eq 429 -and $attempt -lt 4)
        if ($resp.StatusCode -ne 200) { throw (Get-ErrorText $resp.Content) }
        $j    = $resp.Content | ConvertFrom-Json
        $cols = @($j.properties.columns)
        $costIdx = $null; $ridIdx = $null; $curIdx = $null
        for ($c = 0; $c -lt $cols.Count; $c++) {
            if ($null -eq $costIdx -and $cols[$c].type -eq 'Number') { $costIdx = $c }
            if ($cols[$c].name -eq 'ResourceId') { $ridIdx = $c }
            if ($cols[$c].name -eq 'Currency')   { $curIdx = $c }
        }
        foreach ($row in $j.properties.rows) {
            $rid = ([string]$row[$ridIdx]).ToLowerInvariant()
            if (-not $result.ContainsKey($rid)) { $result[$rid] = 0.0 }
            $result[$rid] += [double]$row[$costIdx]
            if ($null -ne $curIdx) { $currency = $row[$curIdx] }
        }
        $uri = $j.properties.nextLink
    } while ($uri)
    return @{ Costs = $result; Currency = $currency }
}

function Format-RefList($Ids, [int]$Max = 15) {
    $list = @($Ids)
    if ($list.Count -eq 0) { return '' }
    $out = foreach ($d in ($list | Select-Object -First $Max)) {
        if ($byId.ContainsKey($d)) {
            $x = $byId[$d]; '{0} [{1}] ({2})' -f $x.name, ($x.type -split '/')[-1], $x.resourceGroup
        } else {
            $p = Get-IdParts $d
            if ($p) {
                $where = if ($p.Sub -eq $TargetSubscriptionId.ToLowerInvariant()) { 'TARGET sub' } else { "sub $($p.Sub.Substring(0,8))..." }
                '{0} [{1}] ({2}, {3})' -f $p.Name, ($p.Type -split '/')[-1], $p.Rg, $where
            } else { $d }
        }
    }
    $s = $out -join '; '
    if ($list.Count -gt $Max) { $s += "; ...(+$($list.Count - $Max) more)" }
    return $s
}
#endregion

#region ---------- Setup ----------
foreach ($m in 'Az.Accounts', 'Az.Resources', 'Az.ResourceGraph', 'Az.Monitor') {
    if (-not (Get-Module -ListAvailable -Name $m)) { throw "Module $m is missing. Install it with: Install-Module $m -Scope CurrentUser" }
    Import-Module $m -ErrorAction Stop
}
if ($ValidateMove -and -not $TargetResourceGroupName) { throw '-ValidateMove needs -TargetResourceGroupName (an existing RG in the target subscription).' }

if (-not (Get-AzContext)) { Connect-AzAccount | Out-Null }

Write-Step 'Locating target subscription'
# Find the target's tenant WITHOUT walking every tenant this account can reach
# (a CSP partner account can see many customer tenants).
$targetSub = $null
$tenantHint = if ($TenantId) { $TenantId } else { (Get-AzContext).Tenant.Id }
try { $targetSub = Get-AzSubscription -SubscriptionId $TargetSubscriptionId -TenantId $tenantHint -ErrorAction Stop | Select-Object -First 1 } catch { }
if (-not $targetSub) {
    throw "Target subscription $TargetSubscriptionId not found in tenant $tenantHint. Re-run with -TenantId <the target's tenant ID>, or Connect-AzAccount -TenantId <tenant> first."
}
$scanTenant = $targetSub.TenantId
Write-Note "Target: $($targetSub.Name) | Tenant: $scanTenant"

Write-Step "Enumerating subscriptions in tenant $scanTenant only"
$tenantSubs = @(Get-AzSubscription -TenantId $scanTenant | Sort-Object Id -Unique)

if ($SourceSubscriptionIds) {
    $sourceSubs = @($tenantSubs | Where-Object { $SourceSubscriptionIds -contains $_.Id })
    $missing = @($SourceSubscriptionIds | Where-Object { $tenantSubs.Id -notcontains $_ })
    if ($missing.Count) { Write-Note "Not found in the target's tenant, skipped: $($missing -join ', ')" }
} else {
    $sourceSubs = @($tenantSubs | Where-Object { $_.Id -ne $TargetSubscriptionId -and $_.State -eq 'Enabled' })
}
$sourceSubs = @($sourceSubs | Where-Object { $_.Id -ne $TargetSubscriptionId })
if ($sourceSubs.Count -eq 0) { throw 'No source subscriptions to assess.' }

if ($Select) {
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        $sourceSubs = @($sourceSubs | Select-Object Name, Id, State | Out-GridView -Title 'Select source subscriptions (Ctrl+click for several), then OK' -PassThru |
                        ForEach-Object { $id = $_.Id; $sourceSubs | Where-Object { $_.Id -eq $id } })
    } else {
        for ($i = 0; $i -lt $sourceSubs.Count; $i++) { Write-Host ("  [{0}] {1}  ({2})" -f ($i + 1), $sourceSubs[$i].Name, $sourceSubs[$i].Id) }
        $pick = Read-Host 'Enter numbers to scan (e.g. 1,3,5) or A for all'
        if ($pick -notmatch '^\s*[Aa]\s*$') {
            $idx = @($pick -split '[,\s]+' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ - 1 } | Where-Object { $_ -ge 0 -and $_ -lt $sourceSubs.Count })
            $sourceSubs = @($idx | Select-Object -Unique | ForEach-Object { $sourceSubs[$_] })
        }
    }
    if ($sourceSubs.Count -eq 0) { throw 'No subscriptions selected.' }
}

Write-Host ''
Write-Host "Source subscriptions to scan ($($sourceSubs.Count)):" -ForegroundColor Green
$sourceSubs | ForEach-Object { Write-Host ("  - {0}  ({1})" -f $_.Name, $_.Id) }
if (-not $Force) {
    $answer = Read-Host 'Proceed with this list? (Y/N)'
    if ($answer -notmatch '^\s*[Yy]') { Write-Host 'Cancelled. Use -SourceSubscriptionIds or -Select to narrow the list.'; return }
}

$subInfo = @{}
foreach ($s in $sourceSubs) {
    $subInfo[$s.Id.ToLowerInvariant()] = [pscustomobject]@{ Name = $s.Name; TenantId = $s.TenantId; SameTenant = ($s.TenantId -eq $targetSub.TenantId) }
    if ($s.TenantId -ne $targetSub.TenantId) { Write-Note "$($s.Name) is in a DIFFERENT tenant than the target - its resources cannot be moved directly." }
}
Write-Note ("Target: {0} | Sources: {1}" -f $targetSub.Name, (($sourceSubs | ForEach-Object Name) -join ', '))

$targetRgId = $null
if ($ValidateMove) {
    Set-AzContext -SubscriptionId $targetSub.Id -TenantId $targetSub.TenantId | Out-Null
    $trg = Get-AzResourceGroup -Name $TargetResourceGroupName -ErrorAction SilentlyContinue
    if (-not $trg) { throw "Target RG '$TargetResourceGroupName' not found in $($targetSub.Name). Create it first (this script never creates anything)." }
    $targetRgId = $trg.ResourceId
}
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
#endregion

#region ---------- Collect ----------
$resQuery = @'
resources
| project id, name, type, location, resourceGroup, subscriptionId, kind, sku, tags, identity, plan, managedBy, properties
'@
$rbacQuery = @'
authorizationresources
| where type =~ 'microsoft.authorization/roleassignments'
| extend scope = tolower(tostring(properties.scope))
| where scope contains '/providers/'
| project scope
'@

$inventory    = New-Object System.Collections.Generic.List[object]
$times        = @{}; $activity = @{}; $locksByScope = @{}; $rbacCount = @{}
$costs        = @{}; $costCurrency = $null; $costStatus = @{}
$usage        = @{}
$now          = Get-Date

foreach ($sub in $sourceSubs) {
    $subKey = $sub.Id.ToLowerInvariant()
    Write-Step "Scanning $($sub.Name)"
    Set-AzContext -SubscriptionId $sub.Id -TenantId $sub.TenantId | Out-Null

    $rows = Invoke-ArgQuery -Query $resQuery -Subscriptions @($sub.Id)
    foreach ($r in $rows) { $inventory.Add($r) }
    Write-Note "$($rows.Count) resources"

    try {
        foreach ($l in @(Get-AzResourceLock)) {
            $lid   = if ($l.LockId) { $l.LockId } else { $l.ResourceId }
            $scope = ($lid -replace '(?i)/providers/Microsoft\.Authorization/locks/[^/]+$', '').ToLowerInvariant()
            if (-not $locksByScope.ContainsKey($scope)) { $locksByScope[$scope] = New-Object System.Collections.Generic.List[string] }
            $locksByScope[$scope].Add("$($l.Name) ($($l.Properties.level))")
        }
    } catch { Write-Note "Locks: $($_.Exception.Message)" }

    try {
        foreach ($ra in (Invoke-ArgQuery -Query $rbacQuery -Subscriptions @($sub.Id))) {
            $top = Get-TopLevelId $ra.scope
            if ($top) { if (-not $rbacCount.ContainsKey($top)) { $rbacCount[$top] = 0 }; $rbacCount[$top]++ }
        }
    } catch { Write-Note "Role assignments: $($_.Exception.Message)" }

    try {
        $uri = $null
        do {
            if ($uri) { $resp = Invoke-AzRestMethod -Method GET -Uri $uri }
            else      { $resp = Invoke-AzRestMethod -Method GET -Path "/subscriptions/$($sub.Id)/resources?`$expand=createdTime,changedTime&api-version=2021-04-01" }
            $j = $resp.Content | ConvertFrom-Json
            foreach ($x in $j.value) { $times[$x.id.ToLowerInvariant()] = $x }
            $uri = $j.nextLink
        } while ($uri)
    } catch { Write-Note "Created/changed times: $($_.Exception.Message)" }

    Write-Note "Reading $ActivityDays days of Activity Log (can take a few minutes)"
    try {
        $events = Get-AzActivityLog -StartTime $now.AddDays(-$ActivityDays) -EndTime $now -MaxRecord 100000
        foreach ($e in $events) {
            if (-not $e.ResourceId) { continue }
            $cat = Get-ActivityValue $e 'Category'
            if ($cat -and $cat -ne 'Administrative') { continue }
            $rid  = $e.ResourceId.ToLowerInvariant()
            $keys = @($rid, (Get-TopLevelId $rid)) | Where-Object { $_ } | Select-Object -Unique
            foreach ($k in $keys) {
                if (-not $activity.ContainsKey($k) -or $activity[$k].Time -lt $e.EventTimestamp) {
                    $activity[$k] = [pscustomobject]@{ Time = $e.EventTimestamp; Caller = $e.Caller; Operation = (Get-ActivityValue $e 'OperationName') }
                }
            }
        }
        if (@($events).Count -ge 100000) { Write-Note 'Activity Log hit the 100,000 record cap; oldest events may be missing.' }
    } catch { Write-Note "Activity Log: $($_.Exception.Message)" }

    if ($IncludeCost) {
        try {
            $c = Get-SubscriptionCost -SubId $sub.Id -Days 30
            foreach ($k in $c.Costs.Keys) { $costs[$k] = $c.Costs[$k] }
            if ($c.Currency) { $costCurrency = $c.Currency }
            $costStatus[$subKey] = 'OK'
        } catch { $costStatus[$subKey] = 'Unavailable'; Write-Note "Cost: $($_.Exception.Message)" }
    }

    if ($IncludeUsageMetrics) {
        $metricTargets = @($rows | Where-Object { $UsageMetricMap.ContainsKey($_.type.ToLowerInvariant()) })
        $i = 0
        foreach ($r in $metricTargets) {
            $i++
            Write-Progress -Activity "Usage metrics ($($sub.Name))" -Status $r.name -PercentComplete (100 * $i / [Math]::Max($metricTargets.Count, 1))
            $map = $UsageMetricMap[$r.type.ToLowerInvariant()]
            try {
                $md = Get-AzMetric -ResourceId $r.id -MetricName $map.Metric -AggregationType $map.Agg -TimeGrain ([TimeSpan]::FromDays(1)) `
                        -StartTime $now.AddDays(-$UsageDays) -EndTime $now -ErrorAction Stop
                $points = @($md.Data)
                $activePts = @($points | Where-Object { $_.($map.Agg) -gt 0 })
                $last = $null
                if ($activePts.Count) { $last = ($activePts | Sort-Object TimeStamp -Descending | Select-Object -First 1).TimeStamp }
                $val = if ($map.Agg -in 'Total', 'Count') { ($points | Measure-Object -Property $map.Agg -Sum).Sum } else { ($points | Measure-Object -Property $map.Agg -Maximum).Maximum }
                $usage[$r.id.ToLowerInvariant()] = [pscustomobject]@{ Last = $last; Label = ('{0} ({1}, {2}d) = {3:N0}' -f $map.Metric, $map.Agg, $UsageDays, [double]$val) }
            } catch {
                $usage[$r.id.ToLowerInvariant()] = [pscustomobject]@{ Last = $null; Label = "Metric unavailable: $($map.Metric)"; Failed = $true }
            }
        }
        Write-Progress -Activity 'Usage metrics' -Completed
    }
}
Write-Step "Collected $($inventory.Count) resources across $($sourceSubs.Count) subscription(s)"
#endregion

#region ---------- Dependency graph and move groups ----------
Write-Step 'Mapping dependencies'
$byId = @{}
foreach ($r in $inventory) { $k = $r.id.ToLowerInvariant(); $byId[$k] = $r; $script:UfParent[$k] = $k }

$hardDeps = @{}; $softRefs = @{}; $refBy = @{}
foreach ($r in $inventory) {
    $id    = $r.id.ToLowerInvariant()
    $type  = $r.type.ToLowerInvariant()
    $found = New-Object 'System.Collections.Generic.HashSet[string]'
    Find-ArmReferences -Node $r.properties -Found $found
    if ($r.managedBy) { Find-ArmReferences -Node ([string]$r.managedBy) -Found $found }
    if ($r.identity -and $r.identity.userAssignedIdentities) {
        foreach ($uami in $r.identity.userAssignedIdentities.PSObject.Properties.Name) { $t = Get-TopLevelId $uami; if ($t) { [void]$found.Add($t) } }
    }
    $selfTop = Get-TopLevelId $id
    $isChild = ($selfTop -and $selfTop -ne $id)
    if ($isChild) { [void]$found.Add($selfTop) }       # child resource -> parent is always hard
    [void]$found.Remove($id)

    $hard = New-Object System.Collections.Generic.List[string]
    $soft = New-Object System.Collections.Generic.List[string]
    foreach ($d in $found) {
        $dType = if ($byId.ContainsKey($d)) { $byId[$d].type.ToLowerInvariant() } else { (Get-IdParts $d).Type }
        $isSoft = ($SoftSourceTypes -contains $type) -or ($SoftTargetTypes -contains $dType)
        if ($isChild -and $d -eq $selfTop) { $isSoft = $false }
        if ($isSoft) { $soft.Add($d) } else { $hard.Add($d) }
        if ($byId.ContainsKey($d)) {
            if (-not $refBy.ContainsKey($d)) { $refBy[$d] = New-Object System.Collections.Generic.List[string] }
            $refBy[$d].Add($id)
            if (-not $isSoft) { Join-Set $id $d }
        }
    }
    $hardDeps[$id] = $hard; $softRefs[$id] = $soft
}

# Number the move groups, biggest first
$groups = @{}
foreach ($k in $byId.Keys) {
    $root = Find-Root $k
    if (-not $groups.ContainsKey($root)) { $groups[$root] = New-Object System.Collections.Generic.List[string] }
    $groups[$root].Add($k)
}
$groupName = @{}; $n = 0
foreach ($g in ($groups.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending)) {
    $n++; $groupName[$g.Key] = 'MG-{0:D4}' -f $n
}
#endregion

#region ---------- Assess each resource ----------
Write-Step 'Assessing move readiness'
$targetSubKey = $TargetSubscriptionId.ToLowerInvariant()
$results = New-Object System.Collections.Generic.List[object]
$rank = @{ 'Blocked' = 4; 'Review' = 3; 'Easy' = 2; 'Skip-Recreate' = 1 }

foreach ($r in $inventory) {
    $id     = $r.id.ToLowerInvariant()
    $type   = $r.type.ToLowerInvariant()
    $subKey = $r.subscriptionId.ToLowerInvariant()
    $root   = Find-Root $id
    $members = $groups[$root]
    $issues = New-Object System.Collections.Generic.List[string]

    # Move support
    $ms = $MoveSupport[$type]
    if (-not $ms) { foreach ($key in $MoveSupport.Keys) { if ($key.EndsWith('/*') -and $type.StartsWith($key.TrimEnd('*'))) { $ms = $MoveSupport[$key]; break } } }
    if (-not $ms -and ($type.Split('/').Count -gt 2)) { $ms = @{ Status = 'Child'; Note = 'Child resource; moves with its parent.' } }
    if (-not $ms) { $ms = @{ Status = 'Supported (assumed)'; Note = '' } }

    $autoCreated = ($AutoCreatedTypes -contains $type) -or (@($AutoCreatedRgPatterns | Where-Object { $r.resourceGroup -match $_ }).Count -gt 0)

    # Dependencies
    $hard = $hardDeps[$id]
    $memberRgs  = @($members | ForEach-Object { $byId[$_].resourceGroup.ToLowerInvariant() } | Select-Object -Unique)
    $memberSubs = @($members | ForEach-Object { $byId[$_].subscriptionId.ToLowerInvariant() } | Select-Object -Unique)
    $crossRg    = @($hard | Where-Object { $byId.ContainsKey($_) -and $byId[$_].resourceGroup -ne $r.resourceGroup })
    $outside    = @($hard | Where-Object { -not $byId.ContainsKey($_) })
    $outsideTarget = @($outside | Where-Object { (Get-IdParts $_).Sub -eq $targetSubKey })
    $outsideOther  = @($outside | Where-Object { (Get-IdParts $_).Sub -ne $targetSubKey })

    # Locks (resource, RG or subscription level)
    $rgScope  = "/subscriptions/$subKey/resourcegroups/$($r.resourceGroup.ToLowerInvariant())"
    $lockList = @()
    foreach ($scope in @($id, (Get-TopLevelId $id), $rgScope, "/subscriptions/$subKey") | Select-Object -Unique) {
        if ($scope -and $locksByScope.ContainsKey($scope)) { $lockList += $locksByScope[$scope] }
    }

    $rbac = 0; if ($rbacCount.ContainsKey($id)) { $rbac = $rbacCount[$id] }
    $idType = if ($r.identity) { [string]$r.identity.type } else { '' }
    $plan   = if ($r.plan -and $r.plan.name) { '{0}/{1}/{2}' -f $r.plan.publisher, $r.plan.product, $r.plan.name } else { '' }

    if (-not $subInfo[$subKey].SameTenant) { $issues.Add('Different Entra tenant than target - direct move not possible') }
    if ($ms.Status -eq 'NotSupported')     { $issues.Add('Resource type not movable') }
    if ($ms.Status -eq 'Restricted')       { $issues.Add('Type has move restrictions (see note)') }
    if ($members.Count -gt 1)              { $issues.Add("Must move with $($members.Count - 1) other resource(s) in $($groupName[$root])") }
    if ($memberRgs.Count -gt 1)            { $issues.Add("Move group spans $($memberRgs.Count) RGs - consolidate into one RG first") }
    if ($memberSubs.Count -gt 1)           { $issues.Add('Move group spans subscriptions - needs planning') }
    if ($outsideOther.Count)               { $issues.Add('Hard dependency outside scanned scope (other sub or deleted)') }
    if ($outsideTarget.Count)              { $issues.Add('Depends on something already in target sub (OK, but verify)') }
    if ($lockList.Count)                   { $issues.Add('Resource lock - remove before move') }
    if ($rbac)                             { $issues.Add("$rbac resource-scoped role assignment(s) - do not move, re-create after") }
    if ($idType -match 'UserAssigned')     { $issues.Add('User-assigned identity - re-check assignment after move') }
    if ($idType -match 'SystemAssigned')   { $issues.Add('System-assigned identity - verify its role assignments after move') }
    if ($plan)                             { $issues.Add('Marketplace plan - accept terms / check offer in target') }

    $readiness =
        if ($autoCreated) { 'Skip-Recreate' }
        elseif (-not $subInfo[$subKey].SameTenant -or $ms.Status -eq 'NotSupported') { 'Blocked' }
        elseif ($issues.Count -eq 0) { 'Easy' }
        else { 'Review' }
    if ($autoCreated) { $issues.Insert(0, 'Auto-created/managed by Azure - usually skip and let Azure recreate in target') }

    # Activity / usage
    $t   = $times[$id]
    $act = $activity[$id]
    $use = $usage[$id]
    $state = ''
    if ($type -eq 'microsoft.compute/virtualmachines') { $state = [string]$r.properties.extended.instanceView.powerState.displayStatus }
    elseif ($r.properties -and $r.properties.PSObject.Properties['state']) { $state = [string]$r.properties.state }

    $signal = 'Unknown (no usage metric for type)'
    if ($state -eq 'VM deallocated') { $signal = 'Idle (VM deallocated)' }
    elseif ($use -and $use.PSObject.Properties['Failed']) { $signal = 'Unknown (metric unavailable)' }
    elseif ($use) { $signal = if ($use.Last) { 'Active' } else { "Idle (no activity in $UsageDays days)" } }
    elseif (-not $IncludeUsageMetrics) { $signal = 'Not checked (-IncludeUsageMetrics)' }

    $costVal = ''
    if ($IncludeCost) {
        if ($costStatus[$subKey] -eq 'OK') { $costVal = if ($costs.ContainsKey($id)) { [Math]::Round($costs[$id], 2) } else { 0 } }
        else { $costVal = 'Unavailable' }
    }

    $tagText = ''
    if ($r.tags) { $tagText = (@($r.tags.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; ') }

    $results.Add([pscustomobject][ordered]@{
        Readiness                 = $readiness
        Issues                    = ($issues -join ' | ')
        Subscription              = $subInfo[$subKey].Name
        ResourceGroup             = $r.resourceGroup
        Name                      = $r.name
        Type                      = $r.type
        Location                  = $r.location
        MoveGroup                 = $groupName[$root]
        MoveGroupSize             = $members.Count
        MoveSupport               = $ms.Status
        MoveSupportNote           = $ms.Note
        ValidationStatus          = if ($ValidateMove) { 'Not validated' } else { 'Not run (-ValidateMove)' }
        ValidationMessage         = ''
        State                     = $state
        Created                   = if ($t) { $t.createdTime } else { '' }
        LastModified              = if ($t) { $t.changedTime } else { '' }
        LastControlPlaneChange    = if ($act) { $act.Time } else { "None in $ActivityDays days" }
        LastChangeBy              = if ($act) { $act.Caller } else { '' }
        LastChangeOperation       = if ($act) { $act.Operation } else { '' }
        LastDataPlaneActivity     = if ($use -and $use.Last) { $use.Last } else { '' }
        UsageMetric               = if ($use) { $use.Label } else { '' }
        UsageSignal               = $signal
        DependsOn                 = Format-RefList $hard
        SoftReferences            = Format-RefList $softRefs[$id]
        ReferencedBy              = Format-RefList $refBy[$id]
        Locks                     = ($lockList -join '; ')
        ResourceRoleAssignments   = $rbac
        ManagedIdentity           = $idType
        MarketplacePlan           = $plan
        Cost30d                   = $costVal
        Currency                  = if ($IncludeCost) { $costCurrency } else { '' }
        Tags                      = $tagText
        SubscriptionId            = $r.subscriptionId
        ResourceId                = $r.id
        _Key                      = $id
        _Root                     = $root
    })
}

# A group is only as ready as its worst member
foreach ($g in ($results | Group-Object _Root)) {
    $worst = ($g.Group | Sort-Object { $rank[$_.Readiness] } -Descending | Select-Object -First 1).Readiness
    if ($worst -eq 'Blocked') {
        foreach ($row in $g.Group) { if ($row.Readiness -ne 'Blocked' -and $row.Readiness -ne 'Skip-Recreate') { $row.Readiness = 'Review'; $row.Issues = ($row.Issues, 'Move group contains a blocked resource' | Where-Object { $_ }) -join ' | ' } }
    }
}
#endregion

#region ---------- Optional: Microsoft move validation (dry run) ----------
$validationLog = New-Object System.Collections.Generic.List[object]
if ($ValidateMove) {
    # Validate each move group's portion per source RG; top-level resources only.
    $batches = $results | Where-Object {
        $_.Readiness -ne 'Skip-Recreate' -and $_.MoveSupport -ne 'NotSupported' -and $_.MoveSupport -ne 'Child' -and
        $subInfo[$_.SubscriptionId.ToLowerInvariant()].SameTenant
    } | Group-Object SubscriptionId, ResourceGroup, MoveGroup

    $i = 0
    foreach ($b in $batches) {
        $i++
        $first = $b.Group[0]
        Write-Step ("Validating {0}/{1}: {2} / {3} / {4} ({5} resources)" -f $i, @($batches).Count, $first.Subscription, $first.ResourceGroup, $first.MoveGroup, $b.Count)
        Set-AzContext -SubscriptionId $first.SubscriptionId | Out-Null
        try   { $v = Test-MoveGroup -SubId $first.SubscriptionId -Rg $first.ResourceGroup -Ids @($b.Group.ResourceId) -TargetRgId $targetRgId }
        catch { $v = @{ Status = 'Error'; Message = $_.Exception.Message } }
        foreach ($row in $b.Group) {
            $row.ValidationStatus  = $v.Status
            $row.ValidationMessage = $v.Message
            if ($v.Status -eq 'Failed' -and $row.Readiness -eq 'Easy') { $row.Readiness = 'Review'; $row.Issues = 'Microsoft validation failed (see message)' }
        }
        $validationLog.Add([pscustomobject]@{ Subscription = $first.Subscription; ResourceGroup = $first.ResourceGroup; MoveGroup = $first.MoveGroup; Resources = $b.Count; Status = $v.Status; Message = $v.Message })
    }
}
#endregion

#region ---------- Output ----------
Write-Step 'Writing output'
$resourceRows = $results | Sort-Object { $rank[$_.Readiness] }, MoveGroup, Subscription, ResourceGroup, Name |
                Select-Object -Property * -ExcludeProperty _Key, _Root

$groupRows = foreach ($g in ($results | Group-Object MoveGroup)) {
    $worst = ($g.Group | Sort-Object { $rank[$_.Readiness] } -Descending | Select-Object -First 1).Readiness
    $costSum = ''
    if ($IncludeCost) { $nums = @($g.Group.Cost30d | Where-Object { $_ -is [double] -or $_ -is [int] }); if ($nums.Count) { $costSum = [Math]::Round(($nums | Measure-Object -Sum).Sum, 2) } }
    $lastAct = @($g.Group | Where-Object { $_.LastControlPlaneChange -is [datetime] } | ForEach-Object LastControlPlaneChange | Sort-Object -Descending | Select-Object -First 1)
    [pscustomobject][ordered]@{
        MoveGroup          = $g.Name
        Readiness          = $worst
        Resources          = $g.Count
        Subscriptions      = (@($g.Group.Subscription | Select-Object -Unique) -join '; ')
        ResourceGroups     = (@($g.Group.ResourceGroup | Select-Object -Unique) -join '; ')
        Types              = (@($g.Group | ForEach-Object { ($_.Type -split '/')[-1] } | Group-Object | ForEach-Object { "$($_.Name) x$($_.Count)" }) -join '; ')
        LastChange         = if ($lastAct.Count) { $lastAct[0] } else { "None in $ActivityDays days" }
        Cost30d            = $costSum
        Validation         = (@($g.Group.ValidationStatus | Select-Object -Unique) -join '; ')
        Members            = (@($g.Group.Name) -join '; ')
    }
}
$groupRows = $groupRows | Sort-Object { $rank[$_.Readiness] }, Resources -Descending

$resourceRows | Export-Csv (Join-Path $OutputFolder 'Resources.csv')  -NoTypeInformation -Encoding UTF8
$groupRows    | Export-Csv (Join-Path $OutputFolder 'MoveGroups.csv') -NoTypeInformation -Encoding UTF8
if ($validationLog.Count) { $validationLog | Export-Csv (Join-Path $OutputFolder 'Validation.csv') -NoTypeInformation -Encoding UTF8 }

if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module ImportExcel
    $xl = Join-Path $OutputFolder 'AzMoveAssessment.xlsx'
    $common = @{ Path = $xl; AutoSize = $true; FreezeTopRow = $true; AutoFilter = $true; BoldTopRow = $true }
    $resourceRows | Export-Excel @common -WorksheetName 'Resources'
    $groupRows    | Export-Excel @common -WorksheetName 'MoveGroups'
    if ($validationLog.Count) { $validationLog | Export-Excel @common -WorksheetName 'Validation' }
    Write-Note "Excel workbook: $xl"
} else {
    Write-Note 'Tip: Install-Module ImportExcel -Scope CurrentUser to also get a formatted .xlsx'
}

Write-Host ''
Write-Host 'Summary by readiness' -ForegroundColor Green
$results | Group-Object Readiness | Sort-Object { $rank[$_.Name] } |
    Select-Object @{ n = 'Readiness'; e = { $_.Name } }, @{ n = 'Resources'; e = { $_.Count } },
                  @{ n = 'MoveGroups'; e = { @($_.Group.MoveGroup | Select-Object -Unique).Count } } | Format-Table -AutoSize
Write-Host "Output folder: $OutputFolder" -ForegroundColor Green
#endregion
