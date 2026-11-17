[CmdletBinding()]
param([string]$CompilerPath='')
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..')).TrimEnd('\')
$fixture=Join-Path $root ('.tools\package-preflight-tests\'+[Guid]::NewGuid().ToString('N'))
$src=Join-Path $fixture 'src\Excelopen';$inst=Join-Path $fixture 'Installer';$ref=Join-Path $fixture 'Reference';$setup=Join-Path $fixture 'publish\Installer\Setup'
foreach($d in @($src,$inst,$ref,$setup,(Join-Path $fixture '.tools'))){[IO.Directory]::CreateDirectory($d)|Out-Null}
$sourceRoot=Join-Path $root 'src';$destRoot=Join-Path $fixture 'src'
foreach($f in Get-ChildItem -LiteralPath $sourceRoot -Recurse -File | Where-Object {$_.FullName -notmatch '[\\/](bin|obj)[\\/]'}){$rel=$f.FullName.Substring($sourceRoot.Length).TrimStart('\');$dest=Join-Path $destRoot $rel;[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($dest))|Out-Null;[IO.File]::Copy($f.FullName,$dest,$false)}
foreach($f in @('ExcelKit.iss','Startup.iss','ComTransaction.ps1','LibreOffice-下載連結.txt','Verify-ExcelKit.ps1','FormalBaseline.txt','Toolchain.json','PackageVersion.ps1','FileVersion.ps1','FileVersionLedger.json','Package-ExcelKit.ps1','Rebuild-Setup.ps1')){Copy-Item -LiteralPath (Join-Path $root ('Installer\'+$f)) -Destination $inst}
foreach($d in @('Manual','Media','VBS','PowerShell','LabVIEW')){Copy-Item -LiteralPath (Join-Path $root ('Reference\'+$d)) -Destination $ref -Recurse}
Copy-Item -LiteralPath (Join-Path $root 'Reference\Examples-README.md') -Destination $ref
$package=Join-Path $inst 'Package-ExcelKit.ps1';$compiler=if($CompilerPath){$CompilerPath}else{Join-Path $root '.tools\inno-portable\ISCC.exe'};[xml]$project=Get-Content -LiteralPath (Join-Path $src 'Excelopen.csproj') -Raw -Encoding UTF8;$fileVersion=[string]$project.SelectSingleNode('/Project/PropertyGroup/FileVersion').'#text';$zone=[TimeZoneInfo]::FindSystemTimeZoneById('Taipei Standard Time');$today=[TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow,$zone).ToString('yyyyMMdd');$version=& (Join-Path $inst 'PackageVersion.ps1') -Mode Suggest -FileVersion $fileVersion -Date $today -Root $fixture
function Test-Preflight([string]$Name,[string]$Tool) {
    $out=Join-Path $fixture ($Name+'.log')
    $prev=$ErrorActionPreference;$ErrorActionPreference='Continue'
    try {& powershell.exe -NoProfile -NonInteractive -File $package -Preflight -InstallerVersion $version -CompilerPath $Tool *> $out; $code=$LASTEXITCODE} finally {$ErrorActionPreference=$prev}
    return $code
}
function Check([bool]$Condition,[string]$Name){if(-not $Condition){throw "FAILED: $Name"};Write-Host "PASS: $Name"}
Check ((Test-Preflight 'valid' $compiler) -eq 0) 'valid preflight'
Check (@(Get-ChildItem -LiteralPath $setup -Force).Count -eq 0) 'preflight does not create output'
Check ((Test-Preflight 'missing-compiler' (Join-Path $fixture 'missing-ISCC.exe')) -ne 0) 'missing compiler rejected'
$class=Join-Path $src 'Class1.cs';$originalBytes=[IO.File]::ReadAllBytes($class);$original=[IO.File]::ReadAllText($class,[Text.Encoding]::UTF8);$pattern='(public\s+static\s+readonly\s+string\s+Version\s*=\s*)"[^"]+"';if([regex]::Matches($original,$pattern).Count -ne 1){throw '測試版本標記不唯一'};$modified=[regex]::Replace($original,$pattern,'$1"0.0.0+20000101000000"');if($modified -eq $original){throw '測試版本標記未改變'}
[IO.File]::WriteAllText($class,$modified,[Text.UTF8Encoding]::new($true))
Check ((Test-Preflight 'version-mismatch' $compiler) -ne 0) 'component version mismatch rejected'
[IO.File]::WriteAllBytes($class,$originalBytes)
Move-Item -LiteralPath (Join-Path $ref 'Examples-README.md') -Destination (Join-Path $ref 'Examples-README.saved')
Check ((Test-Preflight 'missing-reference' $compiler) -ne 0) 'missing reference rejected'
Check (@(Get-ChildItem -LiteralPath $setup -Force).Count -eq 0) 'failure paths preserve empty output'
Move-Item -LiteralPath (Join-Path $ref 'Examples-README.saved') -Destination (Join-Path $ref 'Examples-README.md')
[IO.File]::WriteAllText((Join-Path $setup ('ExcelKit-Setup-'+$version+'.manifest.json')),'fixture')
Check ((Test-Preflight 'partial-collision' $compiler) -ne 0) 'partial output collision rejected'
Write-Host "EVIDENCE: $fixture"