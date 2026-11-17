$ErrorActionPreference='Stop'
$publisher=Join-Path $PSScriptRoot '..\Publish-GitHubRelease.ps1'
$root=Join-Path ([IO.Path]::GetTempPath()) ('excelkit-release-test-'+[Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($root)|Out-Null
$version='1.2.4.3-20990101-01';$base='ExcelKit-Setup-'+$version;$commit='2222222222222222222222222222222222222222'
$exe=Join-Path $root ($base+'.exe');$manifest=Join-Path $root ($base+'.manifest.json');$report=Join-Path $root ($base+'.build.json')
[IO.File]::WriteAllText($exe,'test installer');[IO.File]::WriteAllText($manifest,(@{installer_version=$version}|ConvertTo-Json))
$body=@{installer_version=$version;release_type='test';source_commit=$commit;installer_sha256=(Get-FileHash $exe).Hash;manifest_sha256=(Get-FileHash $manifest).Hash}
[IO.File]::WriteAllText($report,($body|ConvertTo-Json))
$global:mockState='missing';$global:mockFiles=@();$global:mockTamper=$false;$global:mockTagMismatch=$false;$global:mockUploadFail=$false
$oldToken=$env:GH_TOKEN;$env:GH_TOKEN='mock-token'
function global:git { $global:LASTEXITCODE=0; $id=if($global:mockTagMismatch){'1111111111111111111111111111111111111111'}else{'2222222222222222222222222222222222222222'};return ($id+"`t"+'refs/tags/excelkit-package/20990101/01') }
function global:Invoke-WebRequest {if($global:mockState -eq 'missing'){return [pscustomobject]@{StatusCode=200;Content='[]'}};return [pscustomobject]@{StatusCode=200;Content='[{"tag_name":"excelkit-package/20990101/01"}]'}}
function global:gh {
 $a=@($args);$global:LASTEXITCODE=0
 switch($a[1]) {
 'create' {if($a -notcontains '--draft' -or $a -notcontains '--prerelease' -or $a -notcontains '--verify-tag'){throw 'unsafe release creation'};$global:mockState='draft'}
 'upload' {if($global:mockUploadFail){$global:LASTEXITCODE=1;return};$global:mockFiles=@($a|Where-Object {$_ -match '\.(exe|json)$'})}
 'download' {$dir=$a[[array]::IndexOf($a,'--dir')+1];foreach($f in $global:mockFiles){Copy-Item -LiteralPath $f -Destination $dir};if($global:mockTamper){[IO.File]::AppendAllText((Join-Path $dir ([IO.Path]::GetFileName($global:mockFiles[0]))),'tamper')}}
 'edit' {if($global:mockState -ne 'draft'){throw 'invalid state'};$global:mockState='published'}
 default {throw 'unexpected gh invocation'}
 }
}
function Check-Fails([scriptblock]$f,[string]$name){$failed=$false;try{& $f|Out-Null}catch{$failed=$true};if(-not $failed){throw "$name unexpectedly succeeded"};Write-Host "PASS $name"}
try {
 & $publisher -Directory $root -Version $version | Out-Null;if($global:mockState -ne 'published'){throw 'not published'};Write-Host 'PASS draft upload download verification and prerelease'
 Check-Fails {& $publisher -Directory $root -Version $version} 'existing release not overwritten'
 $global:mockState='missing';$global:mockTagMismatch=$true;Check-Fails {& $publisher -Directory $root -Version $version} 'tag commit mismatch rejected';$global:mockTagMismatch=$false
 $global:mockUploadFail=$true;Check-Fails {& $publisher -Directory $root -Version $version} 'upload failure stopped';if($global:mockState -ne 'draft'){throw 'failed upload did not leave draft'};$global:mockUploadFail=$false
 $global:mockState='missing';$global:mockTamper=$true;Check-Fails {& $publisher -Directory $root -Version $version} 'tampered asset rejected';if($global:mockState -ne 'draft'){throw 'tampered asset published'};$global:mockTamper=$false
 $global:mockState='missing';[IO.File]::AppendAllText($exe,'altered');Check-Fails {& $publisher -Directory $root -Version $version} 'local hash mismatch rejected';if($global:mockState -ne 'missing'){throw 'created draft before validating hashes'}
} finally { $env:GH_TOKEN=$oldToken; Remove-Item Function:\git,Function:\Invoke-WebRequest,Function:\gh -ErrorAction SilentlyContinue;$temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\';$full=[IO.Path]::GetFullPath($root);if($full.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase)){Remove-Item -LiteralPath $full -Recurse -Force -ErrorAction SilentlyContinue} }
