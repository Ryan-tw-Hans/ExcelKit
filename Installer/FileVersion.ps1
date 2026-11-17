[CmdletBinding()]
param([ValidateSet('Reserve','Check')][string]$Mode='Check',[string]$Root='')
$ErrorActionPreference='Stop'
if (-not $Root) { $Root = Split-Path -Parent $PSScriptRoot }
$root=[IO.Path]::GetFullPath($Root).TrimEnd('\')
$mutexHash=[Security.Cryptography.SHA256]::Create();try {$mutexKey=([BitConverter]::ToString($mutexHash.ComputeHash([Text.Encoding]::UTF8.GetBytes($root.ToUpperInvariant())))).Replace('-','')}finally{$mutexHash.Dispose()}
$mutex=[Threading.Mutex]::new($false,('Global\ExcelKit.FileVersion.'+$mutexKey));$locked=$false
try {
 try {$locked=$mutex.WaitOne(0)}catch [Threading.AbandonedMutexException] {$locked=$true}
 if(-not $locked){throw 'Another file version operation is running'}
$ledger=Join-Path $root 'Installer\FileVersionLedger.json'
$project=Join-Path $root 'src\Excelopen\Excelopen.csproj'
$setup=Join-Path $root 'publish\Installer\Setup'
[xml]$xml=Get-Content -LiteralPath $project -Raw -Encoding UTF8
$component=[string]$xml.SelectSingleNode('/Project/PropertyGroup/Version').'#text'
$file=[string]$xml.SelectSingleNode('/Project/PropertyGroup/FileVersion').'#text'
$baseline=(Get-Content -LiteralPath (Join-Path $root 'Installer\FormalBaseline.txt') -Raw).Trim()
if ($component -notmatch ('^'+[regex]::Escape($baseline)+'\+\d{14}$')) {throw 'Invalid component version'}
if ($file -notmatch '^\d+\.\d+\.\d+\.\d+$') {throw 'Invalid four-part file version'}
$version=[Version]$file
if ($file -ne "$baseline.$($version.Revision)" -or $version.Revision -lt 1 -or $version.Revision -gt 65534 -or @(@($version.Major,$version.Minor,$version.Build) | Where-Object {$_ -gt 65534}).Count -gt 0) {throw 'File version must be baseline.revision within 0..65534; revision must be positive'}
$files=Get-ChildItem -LiteralPath (Join-Path $root 'src') -Recurse -File | Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' } | Sort-Object FullName
if(-not $files){throw 'No source files'}
$lines=@($files|ForEach-Object { $_.FullName.Substring($root.Length).Replace('\','/').ToLowerInvariant()+'|'+(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash })
$sha=[Security.Cryptography.SHA256]::Create();try {$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($lines -join "`n"))))).Replace('-','')}finally{$sha.Dispose()}
$entries=@();if(Test-Path -LiteralPath $ledger){$doc=Get-Content -LiteralPath $ledger -Encoding UTF8 -Raw|ConvertFrom-Json;if($doc.schema -ne 1){throw 'Invalid version ledger schema'};$entries=@($doc.entries)}
$seen=@{};$previous=$null
foreach($e in $entries){
 if([string]$e.file_version -notmatch '^\d+\.\d+\.\d+\.\d+$' -or [string]$e.source_sha256 -notmatch '^[A-F0-9]{64}$' -or -not $e.component_version){throw 'Invalid version ledger entry'}
 $v=[Version]([string]$e.file_version);if($seen.ContainsKey($v.ToString()) -or ($previous -and $v -le $previous)){throw 'Ledger is duplicate or out of order'};$seen[$v.ToString()]=$true;$previous=$v
}
if(Test-Path -LiteralPath $setup){foreach($f in Get-ChildItem -LiteralPath $setup -Filter '*.build.json' -File){
 try{$r=Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8|ConvertFrom-Json;if([string]$r.file_version -notmatch '^\d+\.\d+\.\d+\.\d+$' -or -not $r.component_version){throw 'Invalid report'};$v=[Version]([string]$r.file_version)}catch{throw "Unreadable version report: $($f.FullName)"}
 if($v -gt $version -or ($v -eq $version -and [string]$r.component_version -ne $component)){throw "File version conflicts with report: $($f.Name)"}
 if($v -eq $version){foreach($currentFile in $files){$relative=$currentFile.FullName.Substring($root.Length).TrimStart('\');$prior=@($r.source_inputs|Where-Object {$_.path -eq $relative});if($prior.Count -ne 1 -or $prior[0].sha256 -ne (Get-FileHash -LiteralPath $currentFile.FullName -Algorithm SHA256).Hash){throw "Reported file version belongs to different source: $($f.Name)"}}}

}}
# Inspect both user and machine registry views; unknown or missing registered DLL is not silently ignored.
$key='SOFTWARE\Classes\CLSID\{87654321-4321-4321-4321-210987654321}\InprocServer32'
foreach($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryHive]::LocalMachine)) {foreach($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {
 $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,$view);try{$k=$base.OpenSubKey($key);if($k){try{
  $code=[string]$k.GetValue('CodeBase');if(-not $code){throw "COM registration has no CodeBase ($hive/$view)"};$uri=[Uri]$code
  if(-not $uri.IsFile -or $uri.IsUnc -or -not (Test-Path -LiteralPath $uri.LocalPath -PathType Leaf)){throw "Registered COM DLL unavailable ($hive/$view)"}
  $info=[Diagnostics.FileVersionInfo]::GetVersionInfo($uri.LocalPath);$installed=[Version]$info.FileVersion
  if($installed -gt $version -or ($installed -eq $version -and -not $info.ProductVersion.StartsWith($component))){throw "File version conflicts with registered DLL ($hive/$view)"}
  if($installed -eq $version){
   $matched=$false
   foreach($reportFile in Get-ChildItem -LiteralPath $setup -Filter '*.build.json' -File){
    $report=Get-Content -LiteralPath $reportFile.FullName -Raw -Encoding UTF8|ConvertFrom-Json
    if($report.file_version -ne $file -or $report.component_version -ne $component){continue}
    $manifestPath=$reportFile.FullName -replace '\.build\.json$','.manifest.json'
    if(-not(Test-Path -LiteralPath $manifestPath -PathType Leaf)){continue}
    if((Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash -ne $report.manifest_sha256){continue}
    $manifest=Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8|ConvertFrom-Json
    $payload=@($manifest.files|Where-Object {$_.path -eq 'bin\net472\ExcelKit.dll'})
    if($payload.Count -eq 1 -and $payload[0].sha256 -eq (Get-FileHash -LiteralPath $uri.LocalPath -Algorithm SHA256).Hash){$matched=$true;break}
   }
   if(-not $matched){throw "Installed file version is already in use without a matching manifest ($hive/$view)"}
  }
 }finally{$k.Dispose()}}}finally{$base.Dispose()}
}}
if($Mode -eq 'Reserve') {
 if($previous -and $version -le $previous){throw 'New reservation must exceed the ledger high-water mark'}
 if(@($entries | Where-Object { $_.component_version -eq $component }).Count){throw 'Component version already reserved'}
 if($seen.ContainsKey($version.ToString())){throw 'File version already reserved'}
 $entries+= [pscustomobject]@{file_version=$file;component_version=$component;source_sha256=$hash}
 $body=[ordered]@{schema=1;entries=@($entries)}|ConvertTo-Json -Depth 5
 $tmp=$ledger+'.new-'+[Guid]::NewGuid().ToString('N');[IO.File]::WriteAllText($tmp,$body,[Text.UTF8Encoding]::new($false))
 if(Test-Path -LiteralPath $ledger){$backup=$ledger+'.backup-'+[Guid]::NewGuid().ToString('N');[IO.File]::Replace($tmp,$ledger,$backup)}else{[IO.File]::Move($tmp,$ledger)}
 Write-Output "Reserved $file for $component source=$hash"
} else {
 $current=@($entries|Where-Object { $_.file_version -eq $file })
 if($current.Count -ne 1 -or $current[0].component_version -ne $component -or $current[0].source_sha256 -ne $hash -or $previous -ne $version){throw 'File version not reserved for exact current source or is below ledger high-water mark'}
 Write-Output "Verified $file for $component source=$hash"
}
} finally {if($locked){$mutex.ReleaseMutex()};$mutex.Dispose()}
