<#
.SYNOPSIS
    N2W Azure Backup Estate Audit, script edition. Read-only.

.DESCRIPTION
    Counts your Azure VMs, what Azure Backup protects, your vaults, policies and reporting plumbing,
    then prices the estate at Azure list prices and compares it with the same estate on N2W.

    What it does:
      - Runs Azure Resource Graph queries (the same ones as the N2W workbook) as you.
      - Reads each Recovery Services vault's storage usage and diagnostic settings (GET only).
      - Fetches public list prices from Microsoft's Retail Prices API (prices.azure.com, no auth).
      - Writes three files to the current folder: a VM CSV, a vault CSV and a summary text file.

    What it does not do:
      - Change anything in Azure. Every Azure call is a GET or a Resource Graph read. Reader is enough.
      - Send anything anywhere except management.azure.com and prices.azure.com.
      - Collect names. The summary line at the end has counts and prices only.

    Runs in Azure Cloud Shell (PowerShell) with nothing to install, or locally with the Az module.

.PARAMETER SubscriptionId
    Limit the scan to these subscriptions. Default: every enabled subscription you can see in the current tenant.

.PARAMETER Region
    Region to price in (Azure ARM name, e.g. eastus, westeurope). Default: the region holding most of your protected VMs.

.PARAMETER UsedRatio
    Used data as a percentage of provisioned disk size. Azure bills the instance fee on used data; Resource Graph
    only sees provisioned size. Default 100, which is the upper bound for the fee.

.PARAMETER N2WTier
    Blob tier N2W would store backups in: Cool (default) or Cold.

.PARAMETER N2WLicense
    N2W price per VM per month. Default 5.

.PARAMETER N2WServerMonthly
    Monthly cost of the VM N2W runs on. Default 0 (left out of the comparison).

.PARAMETER ProdTagKey
.PARAMETER ProdTagValue
    Tag that marks production VMs. Default env=prod (matches prod, production; case-insensitive).

.PARAMETER Offline
    Don't call the Retail Prices API; use the East US list prices built into this script.

.PARAMETER OutputFolder
    Where to write the three files. Default: current folder.

.EXAMPLE
    ./Invoke-N2WBackupAudit.ps1
.EXAMPLE
    ./Invoke-N2WBackupAudit.ps1 -Region westeurope -N2WTier Cold -UsedRatio 70

.LINK
    https://github.com/N2W-Software/azure-backup-audit
#>
[CmdletBinding()]
param(
    [string[]] $SubscriptionId,
    [string]   $Region,
    [ValidateRange(1, 100)] [int] $UsedRatio = 100,
    [ValidateSet('Cool', 'Cold')] [string] $N2WTier = 'Cool',
    [double]   $N2WLicense = 5,
    [double]   $N2WServerMonthly = 0,
    [string]   $ProdTagKey = 'env',
    [string]   $ProdTagValue = 'prod',
    [switch]   $Offline,
    [string]   $OutputFolder = '.'
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.9'

# East US list prices, Azure Retail Prices API, retrieved 2026-11-02. Used with -Offline or if the API is unreachable.
$FallbackPrices = [ordered]@{
    Region = 'eastus'; Source = 'built-in, East US list, 2026-11-02'
    PiUnit = 10.0; SqlFee = 25.0
    Vault  = @{ LRS = 0.0224; ZRS = 0.028; GRS = 0.0448; RAGRS = 0.0569 }
    Cool   = @{ LRS = 0.0152; ZRS = 0.019; GRS = 0.0334 }
    Cold   = @{ LRS = 0.0036; ZRS = 0.004; GRS = 0.0081 }
}

# BEGIN GENERATED QUERIES (build/build_workbook.py writes this block; edit the queries there)
$Q = @{}
$Q['vmp'] = @'
Resources
| where type =~ 'microsoft.compute/virtualmachines'
| extend vmId = tolower(id)
| extend isProd = tolower(tostring(tags)) contains tolower(strcat('"', '{ProdTagKey}', '":"', '{ProdTagValue}'))
| extend os = tostring(properties.storageProfile.osDisk.osType), vmSize = tostring(properties.hardwareProfile.vmSize)
| project vmId, id, name, subscriptionId, resourceGroup, location, os, vmSize, isProd
| join kind=leftouter (
    Resources
    | where type =~ 'microsoft.compute/disks' and isnotempty(managedBy)
    | summarize provGB = sum(toint(properties.diskSizeGB)) by vmId = tolower(tostring(managedBy))
  ) on vmId
| join kind=leftouter (
    RecoveryServicesResources
    | where type =~ 'microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems'
    | where tostring(properties.backupManagementType) =~ 'AzureIaasVM'
    | where tostring(properties.currentProtectionState) in~ ('ProtectionConfigured','ConfiguringProtection','ProtectionError','ConfiguringProtectionFailed','IRPending')
    | extend vmId = tolower(tostring(properties.sourceResourceId))
    | extend vaultPart = tostring(split(tolower(id), '/providers/microsoft.recoveryservices/vaults/')[1])
    | summarize vaults = make_set(tostring(split(vaultPart, '/')[0])), policies = make_set(tostring(properties.policyName)) by vmId
  ) on vmId
| extend provGB = iff(isnull(provGB), 0, provGB)
| extend protected = coalesce(array_length(vaults), 0) > 0
| extend vault = tostring(vaults[0]), policy = tostring(policies[0])
| extend usedGB = todouble(provGB) * {UsedRatio} / 100.0
| extend piUnits = iff(usedGB <= 50, 0.5, ceiling(usedGB / 500.0))
| extend piFee = iff(usedGB <= 50, todouble({PiSmall}), {PiUnit} * ceiling(usedGB / 500.0))
| extend band = case(usedGB <= 50, '1 · up to 50 GB', usedGB <= 500, '2 · 50–500 GB', usedGB <= 1000, '3 · 0.5–1 TB', usedGB <= 2000, '4 · 1–2 TB', '5 · over 2 TB')
| project id, name, subscriptionId, resourceGroup, location, os, vmSize, isProd, provGB, protected, vault, policy
'@
$Q['sql'] = @'
RecoveryServicesResources
| where type =~ 'microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems'
| where tostring(properties.backupManagementType) =~ 'AzureWorkload' and tostring(properties.workloadType) =~ 'SQLDataBase'
| where tostring(properties.currentProtectionState) in~ ('ProtectionConfigured','ConfiguringProtection','ProtectionError','ConfiguringProtectionFailed','IRPending')
| extend inst = strcat(tostring(split(tolower(id), '/protecteditems/')[0]), '|', tolower(tostring(properties.parentName)))
| summarize sqlInstances = dcount(inst), sqlDbs = count()
'@
$Q['sqlEstate'] = @'
Resources
| where type in~ ('microsoft.sqlvirtualmachine/sqlvirtualmachines', 'microsoft.sql/servers/databases')
| where not(type =~ 'microsoft.sql/servers/databases' and name =~ 'master')
| summarize sqlVms = countif(type =~ 'microsoft.sqlvirtualmachine/sqlvirtualmachines'), sqlDbPaas = countif(type =~ 'microsoft.sql/servers/databases')
'@
$Q['disks'] = @'
Resources
| where type in~ ('microsoft.compute/disks', 'microsoft.compute/snapshots', 'microsoft.compute/restorepointcollections')
| extend isUn = type =~ 'microsoft.compute/disks' and tostring(properties.diskState) =~ 'Unattached'
| extend isSnap = type =~ 'microsoft.compute/snapshots'
| extend unGB = iff(isUn, toint(properties.diskSizeGB), 0)
| extend snapGB = iff(isSnap, toint(properties.diskSizeGB), 0)
| extend snapAt = iff(isSnap, todatetime(properties.timeCreated), datetime(null))
| summarize unattached = countif(isUn), unattachedGB = sum(unGB), snaps = countif(isSnap), snapGB = sum(snapGB), oldest = min(snapAt), rpc = countif(type =~ 'microsoft.compute/restorepointcollections')
| extend oldestDays = iff(isnull(oldest), 0, toint(datetime_diff('day', now(), oldest)))
'@
$Q['vaultsSummary'] = @'
Resources
| where type in~ ('microsoft.recoveryservices/vaults', 'microsoft.dataprotection/backupvaults')
    or (type =~ 'microsoft.logic/workflows' and tolower(tostring(tags)) contains '"usedbybackupreports":"true"')
| extend isRs = type =~ 'microsoft.recoveryservices/vaults', isBv = type =~ 'microsoft.dataprotection/backupvaults'
| extend isRpt = type =~ 'microsoft.logic/workflows'
| extend vloc = iff(isRpt, '', location)
| summarize rs = countif(isRs), bv = countif(isBv), rpt = countif(isRpt), locs = dcount(vloc)
| extend vaultRegions = locs - iff(rpt > 0, 1, 0)
'@
$Q['rsPlumb'] = @'
RecoveryServicesResources
| where type in~ ('microsoft.recoveryservices/vaults/backuppolicies', 'microsoft.recoveryservices/vaults/backupjobs', 'microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems')
| extend isPol = type =~ 'microsoft.recoveryservices/vaults/backuppolicies'
| extend polId = iff(isPol, tolower(id), '')
| join kind=leftouter (
    RecoveryServicesResources
    | where type =~ 'microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems' and isnotempty(tostring(properties.policyId))
    | summarize usedBy = count() by polId = tolower(tostring(properties.policyId))
  ) on polId
| extend isDefault = isPol and name in~ ('DefaultPolicy', 'HourlyLogBackup', 'EnhancedPolicy')
| extend isUnused = isPol and isnull(usedBy) and not(isDefault)
| extend isUnusedDefault = isPol and isnull(usedBy) and isDefault
| extend isFailed = type =~ 'microsoft.recoveryservices/vaults/backupjobs' and tostring(properties.status) =~ 'Failed' and todatetime(properties.startTime) > ago(14d)
| extend isStopped = type =~ 'microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems' and tostring(properties.currentProtectionState) =~ 'ProtectionStopped'
| summarize policies = countif(isPol), unusedPolicies = countif(isUnused), unusedDefaults = countif(isUnusedDefault), failedJobs14d = countif(isFailed), stoppedRetained = countif(isStopped)
'@
$Q['dupGroups'] = @'
RecoveryServicesResources
| where type =~ 'microsoft.recoveryservices/vaults/backuppolicies'
| where tostring(properties.backupManagementType) =~ 'AzureIaasVM'
| extend p = properties
| extend vault = tostring(split(tostring(split(tolower(id), '/providers/microsoft.recoveryservices/vaults/')[1]), '/')[0])
| extend sig = strcat(tostring(p.schedulePolicy.scheduleRunFrequency), ' · every ', tostring(p.schedulePolicy.hourlySchedule['interval']), 'h · keep ',
    tostring(p.retentionPolicy.dailySchedule.retentionDuration['count']), 'd / ', tostring(p.retentionPolicy.weeklySchedule.retentionDuration['count']), 'w / ',
    tostring(p.retentionPolicy.monthlySchedule.retentionDuration['count']), 'm / ', tostring(p.retentionPolicy.yearlySchedule.retentionDuration['count']), 'y · instant restore ',
    tostring(p.instantRpRetentionRangeInDays), 'd')
| summarize policyCount = count(), vaultCount = dcount(vault), names = make_set(name) by sig
| where policyCount > 1
| project policyCount, vaultCount
'@
$Q['vaults'] = @'
Resources
| where type in~ ('microsoft.recoveryservices/vaults', 'microsoft.dataprotection/backupvaults')
| extend vaultId = tolower(id)
| extend Kind = iff(type =~ 'microsoft.recoveryservices/vaults', 'Recovery Services', 'Backup vault')
| extend Redundancy = iff(type =~ 'microsoft.recoveryservices/vaults', tostring(properties.redundancySettings.standardTierStorageRedundancy), tostring(properties.storageSettings[0]['type']))
| join kind=leftouter (
    RecoveryServicesResources
    | where type in~ ('microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems', 'microsoft.dataprotection/backupvaults/backupinstances')
    | extend vaultId = iff(type =~ 'microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems', tostring(split(tolower(id), '/backupfabrics/')[0]), tostring(split(tolower(id), '/backupinstances/')[0]))
    | summarize Items = count() by vaultId
  ) on vaultId
| project Vault = id, Kind, Location = location, Redundancy, Items = iff(isnull(Items), 0, Items), Subscription = strcat('/subscriptions/', subscriptionId)
| order by Location asc, Items desc
'@
# END GENERATED QUERIES

# ------------------------------------------------------------------ helpers
function Write-Line([string]$Text = '', [string]$Color) {
    if ($Color) { Write-Host $Text -ForegroundColor $Color } else { Write-Host $Text }
}

function Format-Num($n, [int]$Decimals = 0) {
    if ($null -eq $n) { return '0' }
    return ([double]$n).ToString("N$Decimals", [Globalization.CultureInfo]::InvariantCulture)
}

function Get-InstanceFee([double]$UsedGB, [double]$Unit) {
    # Pricing page: up to 50 GB = half a unit; above that, one unit per started 500 GB.
    if ($UsedGB -le 50) { return [pscustomobject]@{ Units = 0.5; Fee = $Unit / 2 } }
    $u = [math]::Ceiling($UsedGB / 500)
    return [pscustomobject]@{ Units = $u; Fee = $Unit * $u }
}

function Get-Band([double]$UsedGB) {
    if ($UsedGB -le 50) { return 'up to 50 GB' }
    if ($UsedGB -le 500) { return '50-500 GB' }
    if ($UsedGB -le 1000) { return '0.5-1 TB' }
    if ($UsedGB -le 2000) { return '1-2 TB' }
    return 'over 2 TB'
}

function Format-Size([double]$GB) {
    if ($GB -lt 1024) { return '{0} GB' -f (Format-Num $GB) }
    return '{0} TB' -f (Format-Num ($GB / 1024) 1)
}

function Get-Sum($Items, [string]$Property) {
    # StrictMode-safe sum: empty input gives 0, null values are skipped.
    $t = 0.0
    foreach ($i in @($Items)) { if ($null -ne $i -and $null -ne $i.$Property) { $t += [double]$i.$Property } }
    return $t
}

function Invoke-Arg([string]$Name, [hashtable]$Tokens = @{}) {
    $query = $Q[$Name]
    foreach ($k in $Tokens.Keys) { $query = $query.Replace("{$k}", [string]$Tokens[$k]) }
    $rows = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $script:Subs.Count; $i += 1000) {
        $batch = @($script:Subs[$i..([math]::Min($i + 999, $script:Subs.Count - 1))])
        $skip = $null
        do {
            $p = @{ Query = $query; Subscription = $batch; First = 1000 }
            if ($skip) { $p.SkipToken = $skip }
            $r = Search-AzGraph @p
            if ($null -ne $r -and $r.PSObject.Properties['Data']) { $data = $r.Data; $skip = $r.SkipToken }
            else { $data = $r; $skip = $null }
            foreach ($d in @($data)) { if ($null -ne $d) { $rows.Add($d) } }
        } while ($skip)
    }
    return , $rows
}

function Invoke-ArmGet([string]$Path) {
    $r = Invoke-AzRestMethod -Method GET -Path $Path
    if ($r.StatusCode -ge 300) { throw "GET $Path returned $($r.StatusCode)" }
    return ($r.Content | ConvertFrom-Json)
}

function Get-RetailPrices([string]$Region) {
    function Get-Items([string]$Filter) {
        $uri = 'https://prices.azure.com/api/retail/prices?$filter=' + [uri]::EscapeDataString($Filter)
        $items = @()
        while ($uri) {
            $r = Invoke-RestMethod -Uri $uri -Method GET -TimeoutSec 30
            $items += @($r.Items); $uri = $r.NextPageLink
        }
        return $items
    }
    function Pick($Items, [string]$Meter) {
        $hit = @($Items | Where-Object { $_.meterName -eq $Meter -and $_.tierMinimumUnits -eq 0 -and $_.type -eq 'Consumption' }) | Select-Object -First 1
        if (-not $hit) { throw "No retail price for '$Meter' in $Region" }
        return [double]$hit.retailPrice
    }
    $b = Get-Items "serviceName eq 'Backup' and armRegionName eq '$Region' and priceType eq 'Consumption'"
    $s = Get-Items "serviceName eq 'Storage' and armRegionName eq '$Region' and productName eq 'General Block Blob v2' and priceType eq 'Consumption'"
    return [ordered]@{
        Region = $Region; Source = "Azure Retail Prices API, $Region list, $(Get-Date -Format 'yyyy-MM-dd')"
        PiUnit = Pick $b 'Azure VM Protected Instance'
        SqlFee = Pick $b 'SQL Server in Azure VM Protected Instance'
        Vault  = @{ LRS = Pick $b 'Standard LRS Data Stored'; ZRS = Pick $b 'Standard ZRS Data Stored'; GRS = Pick $b 'Standard GRS Data Stored'; RAGRS = Pick $b 'Standard RA-GRS Data Stored' }
        Cool   = @{ LRS = Pick $s 'Cool LRS Data Stored'; ZRS = Pick $s 'Cool ZRS Data Stored'; GRS = Pick $s 'Cool GRS Data Stored' }
        Cold   = @{ LRS = Pick $s 'Cold LRS Data Stored'; ZRS = Pick $s 'Cold ZRS Data Stored'; GRS = Pick $s 'Cold GRS Data Stored' }
    }
}

# ------------------------------------------------------------------ connect
foreach ($m in 'Az.Accounts', 'Az.ResourceGraph') {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        throw "Module $m is missing. In Cloud Shell it is preinstalled; locally run: Install-Module $m -Scope CurrentUser"
    }
    Import-Module $m -ErrorAction Stop | Out-Null
}
$ctx = Get-AzContext
if (-not $ctx) { throw 'Not signed in. Run Connect-AzAccount first (Cloud Shell signs you in automatically).' }

if ($SubscriptionId) { $script:Subs = @($SubscriptionId) }
else { $script:Subs = @(Get-AzSubscription -TenantId $ctx.Tenant.Id -WarningAction SilentlyContinue | Where-Object State -eq 'Enabled' | ForEach-Object Id) }
if ($script:Subs.Count -eq 0) { throw 'No enabled subscriptions visible to this account.' }

$runAt = Get-Date
Write-Line ("N2W Azure Backup Estate Audit  v{0}          read-only · Reader is enough · nothing leaves your tenant" -f $ScriptVersion) Cyan
Write-Line ("Tenant: {0}   Subscriptions: {1}   Run: {2:yyyy-MM-dd HH:mm}" -f $ctx.Tenant.Id, $script:Subs.Count, $runAt)
Write-Line 'Reading Resource Graph...' DarkGray

# ------------------------------------------------------------------ inventory
$safeKey = ($ProdTagKey -replace '["\\'']', '')
$safeVal = ($ProdTagValue -replace '["\\'']', '')
$tok = @{ ProdTagKey = $safeKey; ProdTagValue = $safeVal; UsedRatio = $UsedRatio; PiSmall = 5; PiUnit = 10 }

$vms = Invoke-Arg 'vmp' $tok
function Invoke-ArgSoft([string]$Name, [string[]]$Fields) {
    # Secondary queries: if one fails on an unusual tenant, warn and leave its fields blank instead of stopping.
    try {
        $r = Invoke-Arg $Name
        if ($r.Count -gt 0) { return $r }
        $o = [ordered]@{}; foreach ($f in $Fields) { $o[$f] = 0 }   # no rows: report zeros, not a crash
        return [pscustomobject]$o
    }
    catch {
        $msg = ($_.Exception.Message -split "`n")[0]
        Write-Line "  Query '$Name' failed and was skipped; its fields stay blank. ($msg)" Yellow
        $script:SkippedQueries += $Name
        $o = [ordered]@{}; foreach ($f in $Fields) { $o[$f] = $null }
        return , @([pscustomobject]$o)
    }
}
$script:SkippedQueries = @()
$sql = (Invoke-ArgSoft 'sql' 'sqlInstances', 'sqlDbs')[0]
$sqlE = (Invoke-ArgSoft 'sqlEstate' 'sqlVms', 'sqlDbPaas')[0]
$disk = (Invoke-ArgSoft 'disks' 'unattached', 'unattachedGB', 'snaps', 'snapGB', 'oldestDays')[0]
$vsum = (Invoke-ArgSoft 'vaultsSummary' 'rs', 'bv', 'rpt', 'vaultRegions')[0]
$rsp = (Invoke-ArgSoft 'rsPlumb' 'policies', 'unusedPolicies', 'unusedDefaults', 'failedJobs14d', 'stoppedRetained')[0]
$dups = @(Invoke-ArgSoft 'dupGroups' 'policyCount', 'vaultCount' | Where-Object { $null -ne $_.policyCount })
$vaults = Invoke-Arg 'vaults'

$prot = @($vms | Where-Object { $_.protected })
$unprot = @($vms | Where-Object { -not $_.protected })

# ------------------------------------------------------------------ prices
if (-not $Region) {
    $topRegion = $prot | Group-Object location | Sort-Object Count -Descending | Select-Object -First 1
    if ($topRegion) { $Region = $topRegion.Name }
    if (-not $Region) { $Region = 'eastus' }
}
$prices = $null
if (-not $Offline) {
    try { $prices = Get-RetailPrices $Region }
    catch { Write-Line "Retail Prices API unavailable or incomplete for $Region ($($_.Exception.Message)); using built-in East US list prices." Yellow }
}
if (-not $prices) { $prices = $FallbackPrices }

# ------------------------------------------------------------------ vault storage + diagnostics (GET only)
Write-Line 'Reading vault storage and diagnostic settings...' DarkGray
$storeGB = @{ LRS = 0.0; ZRS = 0.0; GRS = 0.0; RAGRS = 0.0 }
$measured = $true
$usageReadFailed = 0
$vaultRows = foreach ($v in $vaults) {
    $id = [string]$v.Vault
    $diag = $null; $usedGB = $null
    try {
        $d = Invoke-ArmGet "$id/providers/Microsoft.Insights/diagnosticSettings?api-version=2021-05-01-preview"
        $diag = @($d.value | Where-Object { $_.properties.workspaceId }).Count
    } catch { $diag = $null }
    if ($v.Kind -eq 'Recovery Services') {
        try {
            $u = Invoke-ArmGet "$id/usages?api-version=2016-06-01"
            $usedGB = 0.0
            foreach ($x in @($u.value)) {
                $n = [string]$x.name.value
                if ($n -match '^(LRS|ZRS|GRS|RAGRS)StorageUsage$') {
                    $gb = [double]$x.currentValue / 1GB
                    $storeGB[$Matches[1]] += $gb; $usedGB += $gb
                }
            }
        } catch { $measured = $false; $usageReadFailed++; Write-Verbose "usages failed for ${id}: $($_.Exception.Message)" }
    }
    $parts = $id -split '/'
    [pscustomobject]@{
        Vault = $parts[-1]; Kind = $v.Kind; Location = $v.Location; Redundancy = $v.Redundancy; ProtectedItems = $v.Items
        StorageGB = if ($null -ne $usedGB) { [math]::Round($usedGB, 1) } else { '' }
        DiagnosticSettingsToLogAnalytics = if ($null -ne $diag) { $diag } else { 'unknown' }
        ResourceGroup = $parts[4]; SubscriptionId = $parts[2]
    }
}
$vaultRows = @($vaultRows)
$vaultsWithDiag = @($vaultRows | Where-Object { $_.DiagnosticSettingsToLogAnalytics -is [int] -and $_.DiagnosticSettingsToLogAnalytics -gt 0 }).Count
$rsCount = @($vaultRows | Where-Object Kind -eq 'Recovery Services').Count

# ------------------------------------------------------------------ cost
$vmRows = foreach ($v in $vms) {
    $used = [double]$v.provGB * $UsedRatio / 100
    $fee = Get-InstanceFee $used $prices.PiUnit
    [pscustomobject]@{
        VM = $v.name; Protected = [bool]$v.protected; Prod = [bool]$v.isProd; Vault = $v.vault; Policy = $v.policy
        ProvisionedGB = [int]$v.provGB; UsedGBAssumed = [math]::Round($used, 0); FeeBand = Get-Band $used
        AzureInstanceFee = if ($v.protected) { $fee.Fee } else { 0 }; FeeUnits = if ($v.protected) { $fee.Units } else { 0 }
        OS = $v.os; Size = $v.vmSize; Location = $v.location; ResourceGroup = $v.resourceGroup; SubscriptionId = $v.subscriptionId
    }
}
$vmRows = @($vmRows)
$protRows = @($vmRows | Where-Object Protected)
$piFee = Get-Sum $protRows 'AzureInstanceFee'
$piUnits = Get-Sum $protRows 'FeeUnits'
$protUsedGB = Get-Sum $protRows 'UsedGBAssumed'

$totalStoreGB = 0.0; foreach ($x in $storeGB.Values) { $totalStoreGB += [double]$x }
$storageRead = if ($usageReadFailed -gt 0) { 'failed' } elseif ($rsCount -gt 0 -and $totalStoreGB -eq 0) { 'zero' } else { 'ok' }
if (-not $measured -or ($rsCount -gt 0 -and $totalStoreGB -eq 0 -and $protRows.Count -gt 0)) {
    # Couldn't read usage: fall back to one full copy of protected data, at LRS. Conservative.
    $measured = $false
    $storeGB = @{ LRS = [double]$protUsedGB; ZRS = 0.0; GRS = 0.0; RAGRS = 0.0 }
    $totalStoreGB = [double]$protUsedGB
}
$azStore = 0.0; $n2Store = 0.0
$tier = $prices[$N2WTier]
foreach ($r in 'LRS', 'ZRS', 'GRS', 'RAGRS') {
    $azStore += $storeGB[$r] * $prices.Vault[$r]
    $n2r = if ($r -eq 'RAGRS') { 'GRS' } else { $r }
    $n2Store += $storeGB[$r] * $tier[$n2r]
}
$sqlInst = if ($null -ne $sql.sqlInstances) { [int]$sql.sqlInstances } else { 0 }
$sqlFee = $sqlInst * $prices.SqlFee
$azTotal = $piFee + $sqlFee + $azStore
$n2Lic = $protRows.Count * $N2WLicense
$n2Total = $n2Lic + $n2Store + $N2WServerMonthly
$pct = if ($azTotal -gt 0) { [math]::Round(100 * ($azTotal - $n2Total) / $azTotal) } else { 0 }
$otherTier = if ($N2WTier -eq 'Cool') { 'Cold' } else { 'Cool' }
$n2Other = $n2Lic + $N2WServerMonthly
foreach ($r in 'LRS', 'ZRS', 'GRS', 'RAGRS') { $n2r = if ($r -eq 'RAGRS') { 'GRS' } else { $r }; $n2Other += $storeGB[$r] * $prices[$otherTier][$n2r] }
$pctOther = if ($azTotal -gt 0) { [math]::Round(100 * ($azTotal - $n2Other) / $azTotal) } else { 0 }

# ------------------------------------------------------------------ report
$unprotProd = @($unprot | Where-Object { $_.isProd }).Count
$regions = @($vms | ForEach-Object location | Sort-Object -Unique).Count
$repeated = Get-Sum $dups 'policyCount'
if ($script:SkippedQueries -contains 'dupGroups') { $repeated = $null }
$vaultByRegion = ($vaultRows | Group-Object Location | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) $($_.Count)" }) -join ' · '
$bands = $protRows | Group-Object AzureInstanceFee | Sort-Object { [double]$_.Name } | ForEach-Object { "$($_.Count) @ `$$($_.Name)" }
$fmtPct = { param($a, $b) if ($b -gt 0) { '{0,3}%' -f [math]::Round(100 * $a / $b) } else { '  -' } }

function Row([string]$Label, $Value, [string]$Note = '', [string]$Color) {
    Write-Line ('  {0,-38}{1,10}   {2}' -f $Label, $Value, $Note) $Color
}
Write-Line ''
Write-Line 'ESTATE' Cyan
Row 'Azure VMs' (Format-Num $vms.Count) ("in {0} regions" -f $regions)
Row 'Managed disks, unattached' (Format-Num $disk.unattached) ("{0} GB" -f (Format-Num $disk.unattachedGB))
Row 'SQL Server VMs (registered)' (Format-Num $sqlE.sqlVms) 'SQL IaaS Agent extension; hand-installed SQL not counted'
Row 'Azure SQL databases (PaaS)' (Format-Num $sqlE.sqlDbPaas)
Row 'Standalone disk snapshots' (Format-Num $disk.snaps) ("oldest {0} days" -f $disk.oldestDays)
Write-Line ''
Write-Line 'PROTECTION' Cyan
Row 'VMs backed up by Azure Backup' (Format-Num $prot.Count) (& $fmtPct $prot.Count $vms.Count)
$unNote = (& $fmtPct $unprot.Count $vms.Count) + $(if ($unprotProd) { "   <- $unprotProd tagged $safeKey=$safeVal" } else { '' })
Row 'VMs with no backup' (Format-Num $unprot.Count) $unNote $(if ($unprotProd) { 'Yellow' })
Row 'SQL instances backed up in VMs' (Format-Num $sqlInst) ("{0} databases; instance fee applies to each" -f $sql.sqlDbs)
Row 'Failed backup jobs, last 14 days' (Format-Num $rsp.failedJobs14d)
Row 'Stopped, data still retained' (Format-Num $rsp.stoppedRetained)
Write-Line ''
Write-Line 'VAULTS AND POLICIES' Cyan
Row 'Recovery Services vaults' (Format-Num $vsum.rs) $vaultByRegion
Row 'Backup vaults' (Format-Num $vsum.bv)
Row 'Backup policies' (Format-Num $rsp.policies) ("{0} repeat another's schedule and retention; {1} of yours protect nothing (+{2} unused Azure defaults, not counted)" -f $repeated, $rsp.unusedPolicies, $rsp.unusedDefaults)
Write-Line ''
Write-Line 'REPORTING PLUMBING FOUND' Cyan
Row 'Vaults sending to Log Analytics' ("{0} of {1}" -f $vaultsWithDiag, $vaultRows.Count) $(if ($vaultRows.Count -eq 0) { 'no vaults in scope' } elseif ($vaultsWithDiag -lt $vaultRows.Count) { 'the rest are invisible to Backup Reports' } else { '' })
Row 'Logic Apps tagged UsedByBackupReports' (Format-Num $vsum.rpt)
Write-Line ''
if ($prot.Count -eq 0) { Write-Line '  No VMs are actively backed up by Azure Backup in scope. Any cost below is storage still held in your vaults. Use the calculator link with your VM count to price protecting them.' Yellow; Write-Line '' }
Write-Line ("ESTIMATED AZURE BACKUP COST, LIST PRICE, PER MONTH  ({0})" -f $prices.Source) Cyan
Row 'Protected-instance fees' ('$' + (Format-Num $piFee)) ("{0} VMs: {1}" -f $protRows.Count, ($bands -join ' · '))
Row 'SQL-in-VM instance fees' ('$' + (Format-Num $sqlFee)) ("{0} x `${1} (lowest tier)" -f $sqlInst, $prices.SqlFee)
$storeNote = if ($measured) { "{0} measured in your vaults" -f (Format-Size $totalStoreGB) } elseif ($storageRead -eq 'zero') { "{0} ESTIMATED (vaults report 0 GB: new vault or no vault-tier copy yet)" -f (Format-Size $totalStoreGB) } else { "{0} ESTIMATED (couldn't read usage on {1} vault(s); rerun with -Verbose)" -f (Format-Size $totalStoreGB), $usageReadFailed }
Row 'Vault storage' ('$' + (Format-Num $azStore)) $storeNote
Row 'Total' ('$' + (Format-Num $azTotal)) ("/ month   ~ `${0} / year" -f (Format-Num ($azTotal * 12))) White
Write-Line ''
Write-Line 'SAME ESTATE ON N2W' Cyan
Row 'License, flat' ('$' + (Format-Num $n2Lic)) ("{0} VMs x `${1}, SQL included" -f $protRows.Count, $N2WLicense)
Row "Storage on Blob $N2WTier" ('$' + (Format-Num $n2Store)) ("same {0}, same redundancy" -f (Format-Size $totalStoreGB))
if ($N2WServerMonthly -gt 0) { Row 'N2W server VM' ('$' + (Format-Num $N2WServerMonthly)) }
Row 'Total' ('$' + (Format-Num $n2Total)) $(if ($azTotal -le 0 -and $n2Total -le 0) { '/ month   nothing is backed up, so there is nothing to compare' } else { "/ month   {0}% lower   ({1}: {2}%)" -f $pct, $otherTier, $pctOther }) White
Write-Line ("  Excluded on both sides: instant-restore snapshots, Log Analytics ingestion, storage transactions. Used data assumed {0}% of provisioned." -f $UsedRatio) DarkGray
Write-Line '  Cold tier: 90-day minimum storage and per-GB retrieval fees apply; Cool has a 30-day minimum.' DarkGray
if ($N2WServerMonthly -eq 0) { Write-Line '  N2W server VM not included; pass -N2WServerMonthly to add it.' DarkGray }

# ------------------------------------------------------------------ summary + files
$b1 = @($protRows | Where-Object { $_.UsedGBAssumed -le 50 }).Count
$b2 = @($protRows | Where-Object { $_.UsedGBAssumed -gt 50 -and $_.UsedGBAssumed -le 500 }).Count
$b3 = @($protRows | Where-Object { $_.UsedGBAssumed -gt 500 }).Count
$inv = [Globalization.CultureInfo]::InvariantCulture
$pairs = [ordered]@{
    'N2W-AUDIT/1' = $null; source = 'script'; tool_version = $ScriptVersion
    subscriptions = @($vms | ForEach-Object subscriptionId | Sort-Object -Unique).Count; subscriptions_scanned = $script:Subs.Count; regions = $regions; vms = $vms.Count; vms_protected = $prot.Count; vms_unprotected = $unprot.Count
    vms_unprotected_prod = $unprotProd; vm_band_le50 = $b1; vm_band_50_500 = $b2; vm_band_gt500 = $b3; vm_fee_units = $piUnits
    sql_instances = $sqlInst; sql_databases = $sql.sqlDbs; sql_vms_registered = $sqlE.sqlVms; azure_sql_databases = $sqlE.sqlDbPaas
    rs_vaults = $vsum.rs; backup_vaults = $vsum.bv; vault_regions = $vsum.vaultRegions; report_logic_apps = $vsum.rpt
    policies = $rsp.policies; policies_unused = $rsp.unusedPolicies; policies_default_unused = $rsp.unusedDefaults; failed_jobs_14d = $rsp.failedJobs14d; stopped_retained = $rsp.stoppedRetained
    policies_repeated = $repeated; unattached_disks = $disk.unattached; snapshots = $disk.snaps; oldest_snapshot_days = $disk.oldestDays
    vaults_with_diagnostics = $vaultsWithDiag
    vault_storage_gb = [math]::Round($totalStoreGB); vault_storage_basis = $(if ($measured) { 'measured' } else { 'estimate' }); vault_storage_read = $storageRead
    used_ratio_pct = $UsedRatio; price_region = $prices.Region; price_source = $(if ($prices.Source -like 'built-in*') { 'builtin' } else { 'live' })
    price_pi_unit = $prices.PiUnit; price_sql = $prices.SqlFee; price_vault_gb = $(if ($totalStoreGB -gt 0) { [math]::Round($azStore / $totalStoreGB, 5) } else { $prices.Vault.LRS }); price_n2w_vm = $N2WLicense
    price_n2w_gb = $(if ($totalStoreGB -gt 0) { [math]::Round($n2Store / $totalStoreGB, 5) } else { $tier.LRS }); n2w_server = $N2WServerMonthly
    azure_monthly_usd = [math]::Round($azTotal); n2w_monthly_usd = [math]::Round($n2Total); date = $runAt.ToString('yyyy-MM-dd')
}
$summary = ($pairs.GetEnumerator() | ForEach-Object {
        if ($_.Key -eq 'N2W-AUDIT/1') { $_.Key } elseif ($null -eq $_.Value) { '{0}=' -f $_.Key } else { '{0}={1}' -f $_.Key, ([string]::Format($inv, '{0}', $_.Value)) }
    }) -join ';'

New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$stem = Join-Path $OutputFolder ('n2w-backup-audit-{0:yyyy-MM-dd-HHmm}' -f $runAt)
$vmRows | Sort-Object @{ e = 'Protected' }, @{ e = 'Prod'; Descending = $true }, VM | Export-Csv "$stem-vms.csv" -NoTypeInformation -Encoding utf8
$vaultRows | Export-Csv "$stem-vaults.csv" -NoTypeInformation -Encoding utf8
$calc = 'https://n2ws.com/azure-cost-savings-calculator?vms={0}&tb={1}&src=script' -f $(if ($prot.Count -gt 0) { $prot.Count } else { $vms.Count }), ([math]::Round($totalStoreGB / 1024, 2)).ToString($inv)
@(
    "N2W Azure Backup Estate Audit v$ScriptVersion, $($runAt.ToString('yyyy-MM-dd HH:mm'))"
    ''
    'Summary (counts and prices only, no names):'
    $summary
    ''
    "Calculator: $calc"
) | Set-Content "$stem-summary.txt" -Encoding utf8

Write-Line ''
Write-Line 'YOUR SUMMARY (counts and prices only; read it before you share it)' Cyan
Write-Line $summary
Write-Line ''
Write-Line "  Paste it at n2ws.com/azure-backup-audit for the itemised receipt or a 20-minute review with an N2W engineer."
Write-Line "  -> $calc"
Write-Line ("  Files: {0}-vms.csv · {0}-vaults.csv · {0}-summary.txt" -f (Split-Path $stem -Leaf)) DarkGray
