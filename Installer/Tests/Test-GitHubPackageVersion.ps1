$ErrorActionPreference='Stop'
$scriptPath=Join-Path $PSScriptRoot '..\PackageVersion.ps1'
$root=Join-Path ([IO.Path]::GetTempPath()) ('excelkit-version-test-'+[guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($root)|Out-Null
$global:testRefs=@('1111111111111111111111111111111111111111' + "`t" + 'refs/tags/excelkit-package/20261117/01')
$global:testRule=[pscustomobject]@{target='tag';enforcement='active';updated_at='2026-11-17T04:26:25.021Z';conditions=[pscustomobject]@{ref_name=[pscustomobject]@{include=@('refs/tags/excelkit-package/**/*');exclude=@()}};rules=@([pscustomobject]@{type='update'},[pscustomobject]@{type='deletion'})}
$global:remoteFailure=$false;$global:pushFailure=$false;$global:dirty=$false
function global:Invoke-RestMethod { if($global:remoteFailure){throw 'offline'}; return $global:testRule }
function global:git {
 $global:LASTEXITCODE=0;$a=@($args);$op=$a[2]
 if($op -eq 'remote'){return 'https://github.com/Ryan-tw-Hans/ExcelKit.git'}
 if($op -eq 'ls-remote'){
  if($global:remoteFailure){$global:LASTEXITCODE=1;return}
  $pattern=$a[$a.Count-1];if($pattern -like '*`**'){return $global:testRefs};return @($global:testRefs|Where-Object {$_ -like ("*`t"+$pattern)})
 }
 if($op -eq 'diff'){if($global:dirty){$global:LASTEXITCODE=1};return}
 if($op -eq 'ls-files'){return}
 if($op -eq 'rev-parse'){return '2222222222222222222222222222222222222222'}
 if($op -eq 'push'){
  if($global:pushFailure){$global:LASTEXITCODE=1;return}
  $pair=$a[$a.Count-1].Split(':');$line=$pair[0]+"`t"+$pair[1]
  if(@($global:testRefs|Where-Object {$_ -like ("*`t"+$pair[1])}).Count){$global:LASTEXITCODE=1;return}
  $global:testRefs+= $line;return
 }
 throw "unexpected git command: $op"
}
function Invoke-Version([string]$mode,[string]$date,[string]$version='') { & $scriptPath -Mode $mode -Root $root -FileVersion '1.2.4.3' -Date $date -InstallerVersion $version }
function Assert-Equal($a,$b,$name){if($a -cne $b){throw "$name expected $b actual $a"};Write-Host "PASS $name"}
function Assert-Fails([scriptblock]$block,$name){$failed=$false;try{& $block|Out-Null}catch{$failed=$true};if(-not $failed){throw "$name did not fail"};Write-Host "PASS $name"}
try {
 Assert-Equal (Invoke-Version Suggest '20261117') '1.2.4.3-20261117-02' 'historical migration'
 Assert-Equal (Invoke-Version Reserve '20261117' '1.2.4.3-20261117-02') '1.2.4.3-20261117-02' 'atomic reservation'
 Assert-Equal (Invoke-Version Suggest '20261117') '1.2.4.3-20261117-03' 'next sequence'
 Assert-Fails {Invoke-Version Reserve '20261117' '1.2.4.3-20261117-02'} 'no reuse'
 Assert-Equal (Invoke-Version Suggest '20261118') '1.2.4.3-20261118-01' 'next day'
 $global:dirty=$true;Assert-Fails {Invoke-Version Reserve '20261118' '1.2.4.3-20261118-01'} 'dirty source rejected';$global:dirty=$false
 $global:pushFailure=$true;Assert-Fails {Invoke-Version Reserve '20261118' '1.2.4.3-20261118-01'} 'unknown push rejected';$global:pushFailure=$false
 $global:testRule.enforcement='disabled';Assert-Fails {Invoke-Version Suggest '20261118'} 'disabled protection rejected';$global:testRule.enforcement='active'
 $global:testRule.updated_at='2026-11-17T04:27:25Z';Assert-Fails {Invoke-Version Suggest '20261118'} 'changed ruleset rejected';$global:testRule.updated_at='2026-11-17T04:26:25.021Z'
 $global:remoteFailure=$true;Assert-Fails {Invoke-Version Suggest '20261118'} 'offline rejected';$global:remoteFailure=$false
 Assert-Fails {Invoke-Version Suggest '20260230'} 'invalid date rejected'
} finally {Remove-Item -LiteralPath $root -Recurse -Force;Remove-Item Function:\git,Function:\Invoke-RestMethod -ErrorAction SilentlyContinue}
