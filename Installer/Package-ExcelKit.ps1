[CmdletBinding()]
param(
    [string]$InstallerVersion = '',
    [string]$CompilerPath = '',
    [switch]$Preflight
)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd('\')
$setupDir = Join-Path $root 'publish\Installer\Setup'
$workParent = Join-Path $root '.tools\package-work'
function Assert-File([string]$Path) {if(-not(Test-Path -LiteralPath $Path -PathType Leaf)){throw "缺少必要檔案：$Path"};$item=Get-Item -LiteralPath $Path -Force;if($item.Attributes -band [IO.FileAttributes]::ReparsePoint){throw "不可使用重新解析點：$Path"}}
function Assert-Folder([string]$Path) {if(-not(Test-Path -LiteralPath $Path -PathType Container)){throw "缺少必要資料夾：$Path"};$item=Get-Item -LiteralPath $Path -Force;if($item.Attributes -band [IO.FileAttributes]::ReparsePoint){throw "不可使用重新解析點：$Path"}}
function Assert-NoSetupRunning {
    $found=@(Get-Process -ErrorAction Stop | Where-Object { $_.ProcessName -match '^ExcelKit-Setup-' })
    if($found.Count){$detail=($found|ForEach-Object {try {"$($_.ProcessName) PID=$($_.Id) Path=$($_.Path)"}catch{"$($_.ProcessName) PID=$($_.Id) Path=unknown"}}) -join '; ';throw "ExcelKit 安裝程式仍在執行，請先關閉：$detail"}
}
function Get-TreeFiles([string]$Folder) {
    Assert-Folder $Folder
    $items=@(Get-ChildItem -LiteralPath $Folder -Recurse -Force)
    foreach($item in $items){if($item.Attributes -band [IO.FileAttributes]::ReparsePoint){throw "來源包含重新解析點：$($item.FullName)"}}
    return @($items | Where-Object {-not $_.PSIsContainer} | Sort-Object FullName)
}
function Copy-Tree([string]$From,[string]$To) {
    $base=[IO.Path]::GetFullPath($From).TrimEnd('\')
    foreach($f in Get-TreeFiles $base){$rel=$f.FullName.Substring($base.Length).TrimStart('\');$dest=Join-Path $To $rel;[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($dest))|Out-Null;[IO.File]::Copy($f.FullName,$dest,$false)}
}
function Write-Utf8([string]$Path,[string]$Text) {[IO.File]::WriteAllText($Path,$Text,[Text.UTF8Encoding]::new($false))}
function Input-Entry([string]$Relative) {
    $p=Join-Path $root $Relative; Assert-File $p
    [pscustomobject]@{path=$Relative.Replace('/','\');sha256=(Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash;length=(Get-Item -LiteralPath $p).Length}
}
function Check-Inputs {
    $script:baselinePath=Join-Path $root 'Installer\FormalBaseline.txt';Assert-File $baselinePath
    $script:baseline=(Get-Content -LiteralPath $baselinePath -Raw).Trim()
    if($baseline -notmatch '^\d+\.\d+\.\d+$'){throw '正式基準格式無效'}
    foreach($dir in @($root,(Join-Path $root '.tools'),(Join-Path $root 'publish'),(Join-Path $root 'publish\Installer'),$setupDir)){Assert-Folder $dir}
    if(Test-Path -LiteralPath $workParent){Assert-Folder $workParent}
    $project=Join-Path $root 'src\Excelopen\Excelopen.csproj';$class=Join-Path $root 'src\Excelopen\Class1.cs';Assert-File $project;Assert-File $class
    [xml]$xml=Get-Content -LiteralPath $project -Raw -Encoding UTF8
    $script:componentVersion=[string]$xml.SelectSingleNode('/Project/PropertyGroup/Version').'#text'
    $script:assemblyVersion=[string]$xml.SelectSingleNode('/Project/PropertyGroup/AssemblyVersion').'#text'
    $script:fileVersion=[string]$xml.SelectSingleNode('/Project/PropertyGroup/FileVersion').'#text'
    $informational=[string]$xml.SelectSingleNode('/Project/PropertyGroup/InformationalVersion').'#text'
    & (Join-Path $root 'Installer\FileVersion.ps1') -Mode Check -Root $root
    $text=Get-Content -LiteralPath $class -Raw -Encoding UTF8
    $matches=[regex]::Matches($text,'public\s+static\s+readonly\s+string\s+Version\s*=\s*"([^"]+)"')
    if($matches.Count -ne 1 -or $matches[0].Groups[1].Value -ne $componentVersion -or $informational -ne $componentVersion){throw 'csproj 與 Class1.cs 的產品版本不一致'}
    if($componentVersion -notmatch '^\d+\.\d+\.\d+(\+\d{14})?$' -or $componentVersion.Split('+')[0] -ne $baseline -or $assemblyVersion -notmatch '^\d+\.\d+\.\d+\.\d+$' -or $fileVersion -notmatch '^\d+\.\d+\.\d+\.\d+$'){throw '元件版本格式或正式基準不一致'}
    $zone=[TimeZoneInfo]::FindSystemTimeZoneById('Taipei Standard Time');$today=[TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow,$zone).ToString('yyyyMMdd')
    $versionTool=Join-Path $root 'Installer\PackageVersion.ps1';Assert-File $versionTool
    if(-not $script:InstallerVersion){$script:InstallerVersion=& $versionTool -Mode Suggest -FileVersion $fileVersion -Date $today -Root $root}
    if($script:InstallerVersion -notmatch ('^'+[regex]::Escape($fileVersion)+'-(?<date>\d{8})-\d{2,}$')){throw '安裝器版本必須使用 DLL 檔案版本-日期-序號'}
    $script:packageDate=$Matches.date
    if($script:packageDate -ne $today){throw '安裝器版本日期與目前台北日期不同，請重新預檢'}
    $null=& $versionTool -Mode Check -FileVersion $fileVersion -Date $packageDate -InstallerVersion $script:InstallerVersion -Root $root
    $script:artifactName='ExcelKit-Setup-'+$script:InstallerVersion+'.exe'
    $script:reportName=[IO.Path]::GetFileNameWithoutExtension($script:artifactName)+'.build.json'
    $script:manifestName=[IO.Path]::GetFileNameWithoutExtension($script:artifactName)+'.manifest.json'
    foreach($n in @($script:artifactName,$script:reportName,$script:manifestName)){if(Test-Path -LiteralPath (Join-Path $setupDir $n)){throw "輸出已存在，不會覆寫：$n"}}
    $dotnet=Get-Command dotnet.exe -ErrorAction SilentlyContinue;if(-not $dotnet){throw '找不到 dotnet.exe'};$script:dotnetPath=$dotnet.Source
    $toolchainPath=Join-Path $root 'Installer\Toolchain.json';Assert-File $toolchainPath
    $toolchain=Get-Content -LiteralPath $toolchainPath -Encoding UTF8 -Raw | ConvertFrom-Json
    if($toolchain.dotnet_sdk -notmatch '^\d+\.\d+\.\d+$' -or $toolchain.inno_setup -notmatch '^\d+\.\d+\.\d+$'){throw '工具鏈版本設定無效'}
    $script:sdkVersion=[string]$toolchain.dotnet_sdk
    $sdks=@(& $script:dotnetPath --list-sdks);if($LASTEXITCODE -ne 0 -or -not @($sdks|Where-Object { $_ -like ($script:sdkVersion+' *') }).Count){throw "找不到指定 .NET SDK：$script:sdkVersion"}
    Assert-File (Join-Path $root 'src\Excelopen\packages.lock.json')
    if(-not $script:CompilerPath){$candidate=Join-Path $root '.tools\inno-portable\ISCC.exe';if(Test-Path -LiteralPath $candidate){$script:CompilerPath=$candidate}else{$cmd=Get-Command ISCC.exe -ErrorAction SilentlyContinue;if($cmd){$script:CompilerPath=$cmd.Source}}}
    if(-not $script:CompilerPath){throw '找不到 Inno Setup ISCC.exe；請指定 -CompilerPath'}
    $script:CompilerPath=[IO.Path]::GetFullPath($script:CompilerPath);Assert-File $script:CompilerPath
    $script:compilerVersion=(& $script:CompilerPath --version | Select-Object -Last 1);if($LASTEXITCODE -ne 0 -or $script:compilerVersion -ne $toolchain.inno_setup){throw "Inno Setup 版本不符，預期 $($toolchain.inno_setup)，實際 $script:compilerVersion"}
    $script:installerFiles=@('ExcelKit.iss','Startup.iss','ComTransaction.ps1','LibreOffice-下載連結.txt','Verify-ExcelKit.ps1','FormalBaseline.txt','Toolchain.json','PackageVersion.ps1','FileVersion.ps1','FileVersionLedger.json','Package-ExcelKit.ps1','Rebuild-Setup.ps1')
    foreach($f in $script:installerFiles){Assert-File (Join-Path $root ('Installer\'+$f))}
    $script:referenceDirs=@('Manual','Media','VBS','PowerShell','LabVIEW')
    foreach($d in $script:referenceDirs){$folder=Join-Path $root ('Reference\'+$d);$files=@(Get-TreeFiles $folder | Where-Object { -not ($d -eq 'Media' -and $_.FullName -ieq (Join-Path $folder 'Video_2026-08-15_233419.mp4')) });if($files.Count -eq 0){throw "素材資料夾排除後為空：Reference\$d"}}
    Assert-File (Join-Path $root 'Reference\Examples-README.md')
    Assert-NoSetupRunning
}
Check-Inputs
Write-Host "Installer=$InstallerVersion; Component=$componentVersion; Assembly=$assemblyVersion; File=$fileVersion; SDK=$sdkVersion; Inno=$compilerVersion"
Write-Host "Output=$(Join-Path $setupDir $artifactName)"
if($Preflight){Write-Host '預檢完成；未保留版本或建立 staging、發布檔案或執行安裝。請以顯示的版本執行打包。';return}
$null=& (Join-Path $root 'Installer\PackageVersion.ps1') -Mode Reserve -FileVersion $fileVersion -Date $packageDate -InstallerVersion $InstallerVersion -Root $root
$runId=$InstallerVersion.Replace('+','-')+'-'+[Guid]::NewGuid().ToString('N')
$work=Join-Path $workParent $runId
if(Test-Path -LiteralPath $work){throw '獨立 staging 已存在，停止'}
[IO.Directory]::CreateDirectory($work)|Out-Null
try {
    $sourceBase=Join-Path $root 'src';$sourceSnapshot=Join-Path $work 'source\src'
    $inputEntries=[System.Collections.Generic.List[object]]::new()
    foreach($f in Get-TreeFiles $sourceBase){$rel=$f.FullName.Substring($sourceBase.Length).TrimStart('\');if($rel -match '(^|\\)(bin|obj)(\\|$)'){continue};$dest=Join-Path $sourceSnapshot $rel;[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($dest))|Out-Null;[IO.File]::Copy($f.FullName,$dest,$false);$entry=Input-Entry ('src\'+$rel);if((Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash -ne $entry.sha256){throw "來源複製時變更：src\$rel"};$inputEntries.Add($entry)}
    $bundle=Join-Path $work 'bundle';$config=Join-Path $bundle 'Installer\Config';[IO.Directory]::CreateDirectory($config)|Out-Null
    foreach($f in $installerFiles){$src=Join-Path $root ('Installer\'+$f);[IO.File]::Copy($src,(Join-Path $config $f),$false);$entry=Input-Entry ('Installer\'+$f);if((Get-FileHash -LiteralPath (Join-Path $config $f) -Algorithm SHA256).Hash -ne $entry.sha256){throw "來源複製時變更：Installer\$f"};$inputEntries.Add($entry)}
    foreach($d in $referenceDirs){$from=Join-Path $root ('Reference\'+$d);$to=switch($d){'VBS'{Join-Path $bundle 'Examples\VBS'} 'PowerShell'{Join-Path $bundle 'Examples\PowerShell'} 'LabVIEW'{Join-Path $bundle 'Examples\LabVIEW'} default {Join-Path $bundle $d}};foreach($f in Get-TreeFiles $from){$local=$f.FullName.Substring($from.Length).TrimStart('\');if($d -eq 'Media' -and $local -ieq 'Video_2026-08-15_233419.mp4'){continue};$rel=$f.FullName.Substring((Join-Path $root 'Reference').Length).TrimStart('\');$entry=Input-Entry ('Reference\'+$rel);$copied=Join-Path $to $local;[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($copied))|Out-Null;[IO.File]::Copy($f.FullName,$copied,$false);if((Get-FileHash -LiteralPath $copied -Algorithm SHA256).Hash -ne $entry.sha256){throw "素材複製時變更：Reference\$rel"};$inputEntries.Add($entry)}}
    [IO.Directory]::CreateDirectory((Join-Path $bundle 'Examples'))|Out-Null
    [IO.File]::Copy((Join-Path $root 'Reference\Examples-README.md'),(Join-Path $bundle 'Examples\README.md'),$false);$entry=Input-Entry 'Reference\Examples-README.md';if((Get-FileHash -LiteralPath (Join-Path $bundle 'Examples\README.md') -Algorithm SHA256).Hash -ne $entry.sha256){throw '範例說明複製時變更'};$inputEntries.Add($entry)
    $snapshotProject=Join-Path $sourceSnapshot 'Excelopen\Excelopen.csproj';$buildLog=Join-Path $work 'dotnet-build.log'
    $sourceRoot=Join-Path $work 'source'
    Write-Utf8 (Join-Path $sourceRoot 'global.json') ('{"sdk":{"version":"'+$sdkVersion+'","rollForward":"disable"}}')
    $restoreLog=Join-Path $work 'dotnet-restore.log'
    Push-Location $sourceRoot
    try {
        & $dotnetPath restore $snapshotProject --locked-mode --ignore-failed-sources '-p:NuGetAudit=false' *> $restoreLog
        if($LASTEXITCODE -ne 0){throw "dotnet restore 失敗，詳見 $restoreLog"}
        & $dotnetPath build $snapshotProject -c Release --no-restore '-p:PublishToFolder=false' *> $buildLog
        if($LASTEXITCODE -ne 0){throw "dotnet build 失敗，詳見 $buildLog"}
    } finally { Pop-Location }
    foreach($tfm in @('net472','net10.0')){$from=Join-Path $sourceSnapshot ('Excelopen\bin\Release\'+$tfm);Assert-Folder $from;$dest=Join-Path $bundle ('bin\'+$tfm);[IO.Directory]::CreateDirectory($dest)|Out-Null;$files=@(Get-ChildItem -LiteralPath $from -File | Where-Object {$_.Extension -in @('.dll','.json')});if($files.Count -eq 0){throw "$tfm 建置結果為空"};foreach($f in $files){[IO.File]::Copy($f.FullName,(Join-Path $dest $f.Name),$false)}}
    foreach($tfm in @('net472','net10.0')) {
        $dll=Join-Path $bundle ('bin\'+$tfm+'\ExcelKit.dll');Assert-File $dll
        $actualAssembly=[Reflection.AssemblyName]::GetAssemblyName($dll).Version.ToString();$info=[Diagnostics.FileVersionInfo]::GetVersionInfo($dll)
        if($actualAssembly -ne $assemblyVersion -or $info.FileVersion -ne $fileVersion -or -not ($info.ProductVersion -eq $componentVersion -or $info.ProductVersion.StartsWith($componentVersion+'.'))){throw "$tfm DLL 版本不一致：Assembly=$actualAssembly File=$($info.FileVersion) Product=$($info.ProductVersion)"}
    }
    $dll=Join-Path $bundle 'bin\net472\ExcelKit.dll';$hash=(Get-FileHash -LiteralPath $dll -Algorithm SHA256).Hash
    $packagedFileVersion=[Diagnostics.FileVersionInfo]::GetVersionInfo($dll).FileVersion
    if($packagedFileVersion -ne $fileVersion){throw 'net472 DLL file version metadata mismatch'}
    $meta=@("#define PackageAssemblyVersion `"$actualAssembly`"","#define PackageAssemblyHash `"$hash`"","#define PackageComponentVersion `"$componentVersion`"","#define PackageFileVersion `"$packagedFileVersion`"")
    [IO.File]::WriteAllLines((Join-Path $config 'PackageMetadata.iss'),[string[]]$meta,[Text.Encoding]::ASCII)
    $iss=Join-Path $config 'ExcelKit.iss';$issText=[IO.File]::ReadAllText($iss,[Text.Encoding]::UTF8)
    $versionPattern='(?m)^#define ProductVersion "[^"]+"[ \t]*\r?$';$outputPattern='(?m)^OutputBaseFilename=[^\r\n]+\r?$'
    if(([regex]::Matches($issText,$versionPattern)).Count -ne 1 -or ([regex]::Matches($issText,$outputPattern)).Count -ne 1){throw '安裝器版本設定不是唯一項目'}
    $issText=[regex]::Replace($issText,$versionPattern,('#define ProductVersion "'+$InstallerVersion+'"'))
    $issText=[regex]::Replace($issText,$outputPattern,('OutputBaseFilename='+[IO.Path]::GetFileNameWithoutExtension($artifactName)))
    Write-Utf8 $iss $issText
    $files=[System.Collections.Generic.List[object]]::new()
    foreach($map in @(@('bin\net472','required'),@('bin\net10.0','net10'),@('Manual','manual'),@('Media','manual'),@('Examples','examples'))){$dir=Join-Path $bundle $map[0];foreach($f in Get-TreeFiles $dir){$rel=$f.FullName.Substring($bundle.Length).TrimStart('\');$files.Add([pscustomobject]@{path=$rel;group=$map[1];sha256=(Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash})}}
    foreach($pair in @(@('ComTransaction.ps1','Installer\ComTransaction.ps1'),@('Verify-ExcelKit.ps1','Installer\Verify-ExcelKit.ps1'),@('LibreOffice-下載連結.txt','LibreOffice-下載連結.txt'))){$f=Join-Path $config $pair[0];$files.Add([pscustomobject]@{path=$pair[1];group='required';sha256=(Get-FileHash -LiteralPath $f -Algorithm SHA256).Hash})}
    $manifest=[ordered]@{schema=1;installer_version=$InstallerVersion;component_version=$componentVersion;assembly_version=$assemblyVersion;file_version=$fileVersion;files=@($files|Sort-Object path)}
    $manifestPath=Join-Path $config 'InstallManifest.json';Write-Utf8 $manifestPath ($manifest|ConvertTo-Json -Depth 8)
    $compilerLog=Join-Path $work 'inno-build.log';$compilerError=Join-Path $work 'inno-build.stderr.log'
    $compilerProcess=Start-Process -FilePath $CompilerPath -ArgumentList @('"'+$iss+'"') -Wait -PassThru -NoNewWindow -RedirectStandardOutput $compilerLog -RedirectStandardError $compilerError
    if($compilerProcess.ExitCode -ne 0){throw "Inno Setup 編譯失敗，詳見 $compilerLog 與 $compilerError"}
    $exe=Join-Path $bundle ('Installer\Setup\'+$artifactName);Assert-File $exe
    $product=([Diagnostics.FileVersionInfo]::GetVersionInfo($exe).ProductVersion).Trim();if($product -ne $InstallerVersion){throw "安裝檔版本不符：$product"}
    foreach($e in $inputEntries){$current=Input-Entry $e.path;if($current.sha256 -ne $e.sha256 -or $current.length -ne $e.length){throw "打包期間來源已變更：$($e.path)"}}
    Assert-NoSetupRunning
    Assert-Folder $setupDir
    $manifestHash=(Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash;$exeHash=(Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash
    $commit=(& git -C $root rev-parse HEAD);if($LASTEXITCODE -ne 0){throw 'Cannot resolve source commit'};$report=[ordered]@{schema=1;source_commit=$commit;github_run_id=$env:GITHUB_RUN_ID;github_run_attempt=$env:GITHUB_RUN_ATTEMPT;release_type='test';excluded_files=@('Reference\Media\Video_2026-08-15_233419.mp4');created_at=(Get-Date).ToString('o');installer_version=$InstallerVersion;component_version=$componentVersion;assembly_version=$assemblyVersion;file_version=$fileVersion;sdk_version=$sdkVersion;inno_version=$compilerVersion;compiler_sha256=(Get-FileHash -LiteralPath $CompilerPath -Algorithm SHA256).Hash;installer_sha256=$exeHash;manifest_sha256=$manifestHash;setup_script_sha256=(Get-FileHash -LiteralPath $iss -Algorithm SHA256).Hash;package_metadata_sha256=(Get-FileHash -LiteralPath (Join-Path $config 'PackageMetadata.iss') -Algorithm SHA256).Hash;nuget_lock_sha256=(Get-FileHash -LiteralPath (Join-Path $sourceSnapshot 'Excelopen\packages.lock.json') -Algorithm SHA256).Hash;source_inputs=@($inputEntries|Sort-Object path);staging=$work;validation='build-and-static-only';installation_test='not_run'}
    $stagedReport=Join-Path $work $reportName;Write-Utf8 $stagedReport ($report|ConvertTo-Json -Depth 8)
    [IO.File]::Move($exe,(Join-Path $setupDir $artifactName));[IO.File]::Copy($manifestPath,(Join-Path $setupDir $manifestName),$false);[IO.File]::Move($stagedReport,(Join-Path $setupDir $reportName))
    "SUCCESS: $(Join-Path $setupDir $artifactName)";"SHA256: $exeHash";"MANIFEST_SHA256: $manifestHash";"REPORT: $(Join-Path $setupDir $reportName)";"STAGING: $work"
} catch {Write-Error "打包失敗；保留 staging 供檢查：$work`n$($_.Exception.Message)";exit 1}