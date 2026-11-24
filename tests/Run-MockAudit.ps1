param([hashtable]$ScriptArgs = @{})
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$env:PSModulePath = (Join-Path $here 'mockmodules') + [IO.Path]::PathSeparator + $env:PSModulePath
$global:RestCalls = @()
$out = Join-Path $here 'out'; Remove-Item $out -Recurse -ErrorAction SilentlyContinue
& (Join-Path $here '..' 'script' 'Invoke-N2WBackupAudit.ps1') -OutputFolder $out @ScriptArgs
"`n--- test checks ---"
"REST calls: $($global:RestCalls.Count) (all GET; any non-GET would have thrown)"
Get-ChildItem $out | % Name
