[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$BuildReportPath)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..')).TrimEnd('\')
$report=Get-Content -LiteralPath $BuildReportPath -Encoding UTF8 -Raw|ConvertFrom-Json
$workRoot=[IO.Path]::GetFullPath((Join-Path $root '.tools\package-work')).TrimEnd('\')
$staging=[IO.Path]::GetFullPath([string]$report.staging)
if(-not $staging.StartsWith($workRoot+'\',[StringComparison]::OrdinalIgnoreCase)){throw '建置 staging 不在專案預期位置'}
$bundle=Join-Path $staging 'bundle';$cfg=Join-Path $bundle 'Installer\Config'
$test=Join-Path $root ('.tools\package-tests\'+[Guid]::NewGuid().ToString('N'))
if(Test-Path -LiteralPath $test){throw '測試資料夾已存在'}
$fixture=Join-Path $test 'fixture';[IO.Directory]::CreateDirectory((Join-Path $fixture 'bin'))|Out-Null;[IO.Directory]::CreateDirectory((Join-Path $fixture 'Installer'))|Out-Null
Copy-Item -LiteralPath (Join-Path $bundle 'bin\net472') -Destination (Join-Path $fixture 'bin') -Recurse
foreach($f in @('InstallManifest.json','Verify-ExcelKit.ps1','ComTransaction.ps1')) {Copy-Item -LiteralPath (Join-Path $cfg $f) -Destination (Join-Path $fixture 'Installer')}
Copy-Item -LiteralPath (Join-Path $cfg 'LibreOffice-下載連結.txt') -Destination $fixture
$manifest=Join-Path $fixture 'Installer\InstallManifest.json';$selection=Join-Path $fixture 'Installer\install-selection.txt';$mhash=(Get-FileHash -LiteralPath $manifest).Hash
function Set-Selection([string]$Hash,[int]$Manual=0) {
  $tx=[IO.Path]::Combine($env:ProgramData,'ExcelKit','Transactions','nonexistent-test-'+[Guid]::NewGuid().ToString('N'))
  $lines=@(
    "installer_version=$($report.installer_version)"
    "manifest_sha256=$Hash"
    "transaction=$tx"
    'com32=1'
    'com64=0'
    "manual=$Manual"
    'examples=0'
    'net10=0'
  )
  $lines | Set-Content -LiteralPath $selection -Encoding UTF8
}
function Invoke-Check([string]$Name,[bool]$Functional=$false) {
  $args=@('-NoProfile','-NonInteractive','-File',(Join-Path $root 'Installer\Verify-ExcelKit.ps1'),'-InstallDir',$fixture)
  if($Functional){$args+='-Functional'}
  $output=& powershell.exe @args; $exit=$LASTEXITCODE
  $output | Set-Content -LiteralPath (Join-Path $test ($Name+'.json')) -Encoding UTF8
  return @{exit=$exit;report=(($output -join "`n")|ConvertFrom-Json)}
}
function Check([bool]$Condition,[string]$Name){if(-not $Condition){throw "FAILED: $Name"};Write-Host "PASS: $Name"}
Set-Selection $mhash
$r=Invoke-Check 'baseline'
Check (@($r.report.results|Where-Object {$_.name -like 'file:*' -and $_.status -eq 'pass'}).Count -ge 4 -and @($r.report.results|Where-Object {$_.name -like 'file:*' -and $_.status -ne 'pass'}).Count -eq 0) 'required file hashes'
Check ($r.report.overall -ne 'pass') 'unknown transaction or COM mismatch does not pass'
Copy-Item -LiteralPath (Join-Path $cfg 'LibreOffice-下載連結.txt') -Destination (Join-Path $fixture 'LibreOffice-下載連結.original.txt')
Add-Content -LiteralPath (Join-Path $fixture 'LibreOffice-下載連結.txt') -Value 'test-only modification'
$r=Invoke-Check 'tampered' $true
Check ($r.exit -eq 1 -and @($r.report.results|Where-Object {$_.name -like 'file:LibreOffice*' -and $_.status -eq 'fail'}).Count -eq 1) 'tampered file fails'
Check (@($r.report.results|Where-Object {$_.name -eq 'functional:x86' -and $_.status -eq 'not_run'}).Count -eq 1) 'functional execution blocked after failure'
Copy-Item -LiteralPath (Join-Path $fixture 'LibreOffice-下載連結.original.txt') -Destination (Join-Path $fixture 'LibreOffice-下載連結.txt') -Force
Set-Selection $mhash 1
$r=Invoke-Check 'missing-selected'
Check ($r.exit -eq 1 -and @($r.report.results|Where-Object {$_.name -like 'file:Manual*' -and $_.status -eq 'fail'}).Count -gt 0) 'missing selected content fails'
Move-Item -LiteralPath $selection -Destination (Join-Path $fixture 'Installer\install-selection.saved')
$r=Invoke-Check 'missing-selection'
Check ($r.exit -eq 2 -and $r.report.overall -eq 'unknown') 'missing selection is unknown'
$doc=Get-Content -LiteralPath $manifest -Encoding UTF8 -Raw|ConvertFrom-Json
$doc.files[0].path='..\outside-test.txt'
$doc|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $manifest -Encoding UTF8
Set-Selection (Get-FileHash -LiteralPath $manifest).Hash
$r=Invoke-Check 'traversal'
Check ($r.exit -eq 2 -and @($r.report.results|Where-Object {$_.name -eq 'verification' -and $_.detail -match 'Invalid relative path'}).Count -eq 1) 'manifest traversal rejected'
Write-Host "EVIDENCE: $test"