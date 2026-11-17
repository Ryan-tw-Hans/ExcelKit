[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$InstallDir,
    [string]$ExpectedManifestSha256 = '',
    [switch]$Functional
)
$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()
function Add-Result([string]$Name,[string]$Status,[string]$Detail) {
    $results.Add([pscustomobject]@{ name=$Name; status=$Status; detail=$Detail })
}
function Safe-Child([string]$Root,[string]$Relative) {
    if ([string]::IsNullOrWhiteSpace($Relative) -or [IO.Path]::IsPathRooted($Relative) -or $Relative -match '(^|[\\/])\.\.?(?:[\\/]|$)' -or $Relative -match ':') { throw "Invalid relative path: $Relative" }
    $base=[IO.Path]::GetFullPath($Root).TrimEnd('\')
    $full=[IO.Path]::GetFullPath((Join-Path $base $Relative))
    if (-not $full.StartsWith($base+'\',[StringComparison]::OrdinalIgnoreCase)) { throw "Path escapes install directory: $Relative" }
    $parts=$Relative -split '[\\/]'; $current=$base
    foreach($part in $parts) { $current=Join-Path $current $part; if(Test-Path -LiteralPath $current) { $item=Get-Item -LiteralPath $current -Force; if($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Reparse point not allowed: $current" } } }
    return $full
}
function Read-Selection([string]$Path) {
    $map=@{}
    foreach($line in Get-Content -LiteralPath $Path -Encoding UTF8) {
        if ($line -notmatch '^([a-z0-9_]+)=(.*)$') { throw 'Invalid selection record' }
        $key=$Matches[1]; if($map.ContainsKey($key)) { throw "Duplicate selection key: $key" }; $map[$key]=$Matches[2]
    }
    foreach($key in @('installer_version','manifest_sha256','transaction','com32','com64','manual','examples','net10')) { if(-not $map.ContainsKey($key)) { throw "Missing selection key: $key" } }
    foreach($key in @('com32','com64','manual','examples','net10')) { if($map[$key] -notin @('0','1')) { throw "Invalid selection value: $key" } }
    if($map.com32 -eq '0' -and $map.com64 -eq '0') { throw 'No COM architecture recorded' }
    if($map.manifest_sha256 -notmatch '^[0-9A-Fa-f]{64}$') { throw 'Invalid manifest hash' }
    return $map
}
function Read-ComView([string]$View,[string]$ExpectedPath,[string]$ExpectedAssembly) {
    $v=if($View -eq 'x86'){[Microsoft.Win32.RegistryView]::Registry32}else{[Microsoft.Win32.RegistryView]::Registry64}
    $u=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$v)
    $b=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$v)
    $key='SOFTWARE\Classes\CLSID\{87654321-4321-4321-4321-210987654321}\InprocServer32'
    try {
        foreach($name in @('SOFTWARE\Classes\CLSID\{87654321-4321-4321-4321-210987654321}','SOFTWARE\Classes\ExcelKit.ExcelOperator')) { $x=$u.OpenSubKey($name); if($x){$x.Dispose();return @{status='unknown';detail='使用者層 COM 覆寫，需人工確認'} } }
        $p=$b.OpenSubKey('SOFTWARE\Classes\ExcelKit.ExcelOperator\CLSID'); $k=$b.OpenSubKey($key)
        try {
            if(-not $p -or -not $k) {return @{status='fail';detail='找不到必要的 ProgID 或 CLSID 註冊'}}
            if([string]$p.GetValue('') -ne '{87654321-4321-4321-4321-210987654321}') {return @{status='fail';detail='ProgID 指向不同 CLSID'}}
            $asm=[string]$k.GetValue('Assembly'); $url=[string]$k.GetValue('CodeBase'); $sub=@($k.GetSubKeyNames())
            if($sub.Count -ne 1 -or $sub[0] -ne $ExpectedAssembly) { return @{status='unknown';detail='COM 登錄包含多個或不一致的版本'} }
            if($asm -notmatch ('^ExcelKit, Version='+[regex]::Escape($ExpectedAssembly)+'(?:,|$)')) {return @{status='fail';detail="登錄 Assembly 不符：$asm"}}
            $uri=[Uri]$url; if(-not $uri.IsFile -or $uri.IsUnc) {return @{status='unknown';detail='CodeBase 不是可確認的本機檔案'}}
            if([IO.Path]::GetFullPath($uri.LocalPath) -ine [IO.Path]::GetFullPath($ExpectedPath)) {return @{status='fail';detail="CodeBase 不符：$url"}}
            return @{status='pass';detail="Assembly $ExpectedAssembly；$($uri.LocalPath)"}
        } finally {if($p){$p.Dispose()};if($k){$k.Dispose()}}
    } finally {$b.Dispose();$u.Dispose()}
}
try {
    $root=[IO.Path]::GetFullPath($InstallDir).TrimEnd('\')
    if(-not(Test-Path -LiteralPath $root -PathType Container)) { throw "找不到安裝資料夾：$root" }
    if((Get-Item -LiteralPath $root -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {throw '安裝根目錄不可為重新解析點'}
    $manifestPath=Safe-Child $root 'Installer\InstallManifest.json'
    $selectionPath=Safe-Child $root 'Installer\install-selection.txt'
    if(-not(Test-Path -LiteralPath $manifestPath -PathType Leaf) -or -not(Test-Path -LiteralPath $selectionPath -PathType Leaf)) {throw '缺少安裝清單或安裝選項紀錄；舊版安裝無法使用本腳本確認'}
    $selection=Read-Selection $selectionPath
    $manifestHash=(Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
    if($manifestHash -ine $selection.manifest_sha256) {throw '安裝清單雜湊與安裝選項紀錄不一致'}
    if($ExpectedManifestSha256) {
        if($ExpectedManifestSha256 -notmatch '^[0-9a-fA-F]{64}$' -or $manifestHash -ine $ExpectedManifestSha256) {throw '安裝清單與外部指定雜湊不一致'}
        Add-Result 'manifest-authentication' 'pass' '與外部提供的清單 SHA-256 一致'
    } else {Add-Result 'manifest-authentication' 'not_run' '未提供外部清單雜湊；已比對安裝紀錄，但不能證明清單來源'}
    $manifest=Get-Content -LiteralPath $manifestPath -Encoding UTF8 -Raw | ConvertFrom-Json
    if($manifest.schema -ne 1 -or $manifest.installer_version -ne $selection.installer_version -or -not $manifest.files -or $manifest.assembly_version -notmatch '^\d+\.\d+\.\d+\.\d+$' -or $manifest.component_version -notmatch '^\d+\.\d+\.\d+(\+[0-9]{14})?$') {throw '安裝清單格式或版本無效'}
    $seen=@{}; $groups=@{required=0;manual=0;examples=0;net10=0}
    foreach($f in $manifest.files) {
        $relative=[string]$f.path; $group=[string]$f.group; $hash=[string]$f.sha256
        if(-not $groups.ContainsKey($group) -or $hash -notmatch '^[0-9a-fA-F]{64}$' -or $seen.ContainsKey($relative.ToLowerInvariant())) {throw "清單項目無效或重複：$relative"}
        $seen[$relative.ToLowerInvariant()]=$true; $groups[$group]++
        $file=Safe-Child $root $relative
        if($group -ne 'required' -and $selection[$group] -ne '1') {continue}
        if(-not(Test-Path -LiteralPath $file -PathType Leaf)) {Add-Result "file:$relative" 'fail' '檔案不存在或無法讀取'; continue}
        try { $actual=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash; if($actual -ieq $hash){Add-Result "file:$relative" 'pass' 'SHA-256 相符'}else{Add-Result "file:$relative" 'fail' 'SHA-256 不符'} } catch {Add-Result "file:$relative" 'unknown' '無法讀取檔案或計算雜湊'}
    }
    foreach($g in @('required','manual','examples','net10')) {if($groups[$g] -eq 0) {throw "清單沒有必要群組：$g"}}
    foreach($f in @('bin\net472\excelkit.dll','installer\comtransaction.ps1','installer\verify-excelkit.ps1')) {if(-not $seen.ContainsKey($f)) {throw "清單缺少必要檔案：$f"}}
    $dll=Safe-Child $root 'bin\net472\ExcelKit.dll'
    if(Test-Path -LiteralPath $dll -PathType Leaf) {
        try { $actualAssembly=[Reflection.AssemblyName]::GetAssemblyName($dll).Version.ToString(); $info=[Diagnostics.FileVersionInfo]::GetVersionInfo($dll); if($actualAssembly -eq $manifest.assembly_version -and $info.FileVersion -eq $manifest.file_version -and ($info.ProductVersion -eq $manifest.component_version -or $info.ProductVersion.StartsWith($manifest.component_version+'.'))) {Add-Result 'dll-version' 'pass' "Assembly=$actualAssembly; File=$($info.FileVersion); Product=$($info.ProductVersion)"} else {Add-Result 'dll-version' 'fail' "Assembly=$actualAssembly; File=$($info.FileVersion); Product=$($info.ProductVersion)"} } catch {Add-Result 'dll-version' 'unknown' '無法讀取 DLL 版本'}
    } else {Add-Result 'dll-version' 'fail' 'DLL 遺失'}
    $txRoot=[IO.Path]::GetFullPath((Join-Path $env:ProgramData 'ExcelKit\Transactions')).TrimEnd('\')
    $tx=[IO.Path]::GetFullPath($selection.transaction)
    if(-not $tx.StartsWith($txRoot+'\',[StringComparison]::OrdinalIgnoreCase)) {throw '交易紀錄路徑不在 ExcelKit 交易資料夾'}
    try { $relative=$tx.Substring($txRoot.Length+1); $txSafe=Safe-Child $txRoot ($relative+'\state.txt'); $state=(Get-Content -LiteralPath $txSafe -Encoding UTF8 -Raw).Trim(); if($state -eq 'complete'){Add-Result 'transaction' 'pass' 'COM 交易已完成'}else{Add-Result 'transaction' 'unknown' "交易狀態：$state"} } catch {Add-Result 'transaction' 'unknown' '無法讀取 COM 交易狀態；可能需要管理員權限'}
    foreach($pair in @(@('com32','x86'),@('com64','x64'))) {
        $key=$pair[0];$view=$pair[1]
        if($selection[$key] -ne '1'){Add-Result "com:$view" 'not_run' '本次未選擇';continue}
        if($view -eq 'x64' -and -not [Environment]::Is64BitOperatingSystem){Add-Result "com:$view" 'fail' '作業系統不支援 64 位元';continue}
        try {$r=Read-ComView $view $dll $manifest.assembly_version;Add-Result "com:$view" $r.status $r.detail} catch {Add-Result "com:$view" 'unknown' '無法讀取 COM 登錄'}
    }
    $versionKey='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{62953C31-A16D-469E-8866-6218DF2B3D2A}_is1';$matches=0
    foreach($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {if($view -eq [Microsoft.Win32.RegistryView]::Registry64 -and -not [Environment]::Is64BitOperatingSystem){continue};$b=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$view);try{$k=$b.OpenSubKey($versionKey);if($k){try{$loc=[string]$k.GetValue('InstallLocation');$ver=[string]$k.GetValue('DisplayVersion');if($loc -and [IO.Path]::GetFullPath($loc).TrimEnd('\') -ieq $root){$matches++;if($ver -ne $manifest.installer_version){Add-Result 'installer-version' 'fail' "已安裝版本 $ver 與預期 $($manifest.installer_version) 不同"}else{Add-Result 'installer-version' 'pass' $ver}}}finally{$k.Dispose()}}}finally{$b.Dispose()}}
    if($matches -eq 0){Add-Result 'installer-version' 'unknown' '未找到符合資料夾的安裝器紀錄'}elseif($matches -gt 1){Add-Result 'installer-version-ambiguity' 'unknown' '兩個登錄視圖均有安裝紀錄，需人工確認'}
    if($Functional) {
        foreach($pair in @(@('com32','x86'),@('com64','x64'))) {
            $key=$pair[0];$view=$pair[1]; if($selection[$key] -ne '1'){continue}
            $check=@($results|Where-Object {$_.name -eq "com:$view"})
            if(@($results|Where-Object {$_.status -in @('fail','unknown')}).Count -gt 0) {Add-Result "functional:$view" 'not_run' '唯讀驗證有失敗或未確認項目';continue}
            if($check.Count -ne 1 -or $check[0].status -ne 'pass') {Add-Result "functional:$view" 'not_run' 'COM 登錄未通過安全檢查';continue}
            $ps=if($view -eq 'x86' -and [Environment]::Is64BitOperatingSystem){Join-Path $env:windir 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'}elseif($view -eq 'x64' -and -not [Environment]::Is64BitProcess){Join-Path $env:windir 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'}else{Join-Path $env:windir 'System32\WindowsPowerShell\v1.0\powershell.exe'}
            $code='$ErrorActionPreference="Stop"; $o=New-Object -ComObject ExcelKit.ExcelOperator; try {$o.SetLogEnabled($false); $v=$o.GetVersionInfo(); if($v -notlike "*'+$manifest.component_version+'*"){throw "Version mismatch: $v"}; $o.CreateNewWorkbook(); $o.SetCellValue("Sheet1","A1","ExcelKit verification"); if($o.GetCellValue("Sheet1","A1") -ne "ExcelKit verification"){throw "Workbook mismatch"}; Write-Output $v} finally {if($o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}'
            $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($code))
            try {$output=& $ps -NoProfile -NonInteractive -EncodedCommand $encoded 2>&1; if($LASTEXITCODE -eq 0){Add-Result "functional:$view" 'pass' ($output -join ' ')}else{Add-Result "functional:$view" 'fail' (($output -join ' ') + " (exit=$LASTEXITCODE)")}} catch {Add-Result "functional:$view" 'unknown' $_.Exception.Message}
        }
    } else {Add-Result 'functional' 'not_run' '未啟用 COM 功能測試'}
} catch {Add-Result 'verification' 'unknown' $_.Exception.Message}
$failed=@($results|Where-Object {$_.status -eq 'fail'}).Count; $unknown=@($results|Where-Object {$_.status -eq 'unknown'}).Count
$overall=if($failed){'fail'}elseif($unknown){'unknown'}else{'pass'}
[pscustomobject]@{ overall=$overall; install_dir=$InstallDir; checked_at=(Get-Date).ToString('o'); functional_requested=[bool]$Functional; results=@($results) } | ConvertTo-Json -Depth 8
if($overall -eq 'fail'){exit 1}elseif($overall -eq 'unknown'){exit 2}else{exit 0}