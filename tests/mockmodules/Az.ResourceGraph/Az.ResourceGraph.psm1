$script:rng = [Random]::new(7)
function New-Vms {
  $locs = 'eastus','westeurope','uksouth'
  $sizes = @(30,40,128,256,300,512,600,900,1100,2500)
  $out = foreach ($i in 1..312) {
    $prot = $i -le 241
    $prod = (-not $prot) -and ($i -le 260)
    [pscustomobject]@{ id = "/subscriptions/0000000$(1+$i%4)-aaaa/resourceGroups/rg$($i%9)/providers/Microsoft.Compute/virtualMachines/vm$i"; name = "vm$i"
      subscriptionId = "0000000$(1+$i%4)-aaaa-bbbb-cccc-dddddddddddd"; resourceGroup = "rg$($i%9)"; location = $locs[$i % 3 -eq 0 ? 1 : ($i % 7 -eq 0 ? 2 : 0)]
      os = ($i % 7 -eq 0 ? 'Linux' : 'Windows'); vmSize = 'Standard_D4s_v5'; isProd = $prod; provGB = $sizes[$i % 10]
      protected = ($prot ? $true : $null); vault = ($prot ? "vault-$(1+$i%7)" : ''); policy = ($prot ? 'DailyPolicy' : '') }
  }
  $out
}
$script:VMS = New-Vms
function Search-AzGraph { param($Query, $Subscription, $First, $SkipToken)
  if ($Query -match "virtualmachines'\s*\n\| extend vmId") {
    $start = if ($SkipToken) { [int]$SkipToken } else { 0 }
    $page = @($script:VMS[$start..([math]::Min($start + 99, $script:VMS.Count - 1))])   # page size 100 to exercise paging
    $next = if ($start + 100 -lt $script:VMS.Count) { [string]($start + 100) } else { $null }
    return [pscustomobject]@{ Data = $page; SkipToken = $next }
  }
  if ($Query -match 'sqlvirtualmachines') { return ,@([pscustomobject]@{ sqlVms = 3; sqlDbPaas = 7 }) }
  if ($Query -match 'sqlInstances') { return ,@([pscustomobject]@{ sqlInstances = 14; sqlDbs = 61 }) }
  if ($Query -match 'restorepointcollections') { return ,@([pscustomobject]@{ unattached = 41; unattachedGB = 5120; snaps = 88; snapGB = 9000; oldest = '2023-01-01'; rpc = 241; oldestDays = 1401 }) }
  if ($Query -match 'usedbybackupreports') { return ,@([pscustomobject]@{ rs = 7; bv = 3; rpt = 3; locs = 4; vaultRegions = 3 }) }
  if ($Query -match 'stoppedRetained') { return ,@([pscustomobject]@{ policies = 23; unusedPolicies = 4; unusedDefaults = 6; failedJobs14d = 9; stoppedRetained = 12 }) }
  if ($Query -match 'policyCount > 1') { return ,@([pscustomobject]@{ policyCount = 7; vaultCount = 7 }, [pscustomobject]@{ policyCount = 4; vaultCount = 3 }) }
  if ($Query -match 'Kind = iff') {
    $rows = foreach ($n in 1..10) { [pscustomobject]@{ Vault = "/subscriptions/00000001-aaaa/resourceGroups/rgb/providers/Microsoft.RecoveryServices/vaults/vault-$n"
        Kind = ($n -le 7 ? 'Recovery Services' : 'Backup vault'); Location = @('eastus','westeurope','uksouth')[$n % 3]; Redundancy = ($n -le 3 ? 'GeoRedundant' : 'LocallyRedundant'); Items = 30 + $n } }
    return [pscustomobject]@{ Data = @($rows); SkipToken = $null }
  }
  throw "Unmocked query: $($Query.Substring(0,80))"
}
