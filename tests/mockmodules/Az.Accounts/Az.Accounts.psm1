function Get-AzContext { [pscustomobject]@{ Tenant = [pscustomobject]@{ Id = 'contoso.onmicrosoft.com' } } }
function Get-AzSubscription { param($TenantId) 1..4 | % { [pscustomobject]@{ Id = "0000000$_-aaaa-bbbb-cccc-dddddddddddd"; State = 'Enabled' } } }
function Invoke-AzRestMethod { param($Method, $Path)
  if ($Method -ne 'GET') { throw "NON-GET CALL: $Method $Path" }
  $global:RestCalls += ,$Path
  if ($Path -match 'diagnosticSettings') {
    $n = if ($Path -match 'vault-(\d+)') { [int]$Matches[1] } else { 0 }
    $v = if ($n % 5 -ne 0) { @(@{ name = 'toLA'; properties = @{ workspaceId = '/ws/1' } }) } else { @() }
    return [pscustomobject]@{ StatusCode = 200; Content = (@{ value = $v } | ConvertTo-Json -Depth 5) }
  }
  if ($Path -match '/usages') {
    $n = [int]([regex]::Match($Path,'vault-(\d+)').Groups[1].Value)
    $val = @(@{ name=@{value='GRSStorageUsage'}; currentValue = [double]$(if ($n -le 3) { 30TB } else { 0 }); unit='Bytes' },
             @{ name=@{value='LRSStorageUsage'}; currentValue = [double]$(if ($n -gt 3) { 23TB } else { 0 }); unit='Bytes' },
             @{ name=@{value='LRSDedupStorageUsage'}; currentValue = 99TB; unit='Bytes' })
    return [pscustomobject]@{ StatusCode = 200; Content = (@{ value = $val } | ConvertTo-Json -Depth 5) }
  }
  [pscustomobject]@{ StatusCode = 404; Content = '{}' }
}
