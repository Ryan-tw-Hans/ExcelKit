$ErrorActionPreference = 'Stop'
if ($env:USERNAME -ne 'WDAGUtilityAccount') { throw 'Sandbox only' }
$out = 'C:\SandboxOutput\optional-20261116091614'
New-Item -ItemType Directory -Path $out -Force | Out-Null
Start-Transcript (Join-Path $out 'transcript.txt') | Out-Null
$app = 'C:\Program Files (x86)\ExcelKit'
function Check([bool]$ok, [string]$name) {
    if (-not $ok) { throw "FAILED: $name" }
    "PASS: $name" | Add-Content (Join-Path $out 'checks.txt')
}
function Install([string]$name, [string]$tasks, [string]$lang = 'english') {
    $p = Start-Process C:\SandboxInput\OptionalInstaller.exe -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /LANG=$lang /DIR=`"$app`" /TASKS=$tasks /LOG=`"$out\$name.log`"" -Wait -PassThru
    Check ($p.ExitCode -eq 0) "$name exit=$($p.ExitCode)"
}
function Uninstall([string]$name) {
    $p = Start-Process (Join-Path $app 'unins000.exe') -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=`"$out\$name.log`"" -Wait -PassThru
    Check ($p.ExitCode -eq 0) "$name exit=$($p.ExitCode)"
}
try {
    $comScript = Join-Path $out 'com.ps1'
    @('$ErrorActionPreference="Stop"', '$o=New-Object -ComObject ExcelKit.ExcelOperator', '$o.CreateNewWorkbook()', '$o.SetCellValue("Sheet1","A1","smoke")', 'if($o.GetCellValue("Sheet1","A1") -ne "smoke"){exit 2}') | Set-Content $comScript -Encoding ASCII
    for ($i = 0; $i -lt 8; $i++) {
        $tasks = @('com32','com64')
        if ($i -band 1) { $tasks += 'manual' }
        if ($i -band 2) { $tasks += 'examples' }
        if ($i -band 4) { $tasks += 'net10' }
        Install "install-$i" ($tasks -join ',')
        foreach ($item in @(@('Manual',1),@('Media',1),@('Examples',2),@('bin\net10.0',4))) {
            $selected = [bool]($i -band $item[1]); $dest = Join-Path $app $item[0]
            Check ((Test-Path $dest) -eq $selected) "case $i $($item[0]) presence"
            if ($selected) {
                $source = Join-Path 'C:\Published' $item[0]
                foreach ($file in Get-ChildItem $source -File -Recurse) {
                    $target = Join-Path $dest $file.FullName.Substring($source.Length).TrimStart('\')
                    Check ((Test-Path $target) -and ((Get-FileHash $file.FullName).Hash -eq (Get-FileHash $target).Hash)) "case $i hash $($file.Name)"
                }
            }
        }
        $shell = New-Object -ComObject WScript.Shell
        $links = @(Get-ChildItem (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\ExcelKit') -Filter '*.lnk' -ErrorAction SilentlyContinue | ForEach-Object { $shell.CreateShortcut($_.FullName).TargetPath })
        Check (($links -contains (Join-Path $app 'Manual')) -eq [bool]($i -band 1)) "case $i manual shortcut"
        Check (($links -contains (Join-Path $app 'Examples')) -eq [bool]($i -band 2)) "case $i examples shortcut"
        foreach ($exe in @("$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe", "$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe")) {
            & $exe -NoProfile -ExecutionPolicy Bypass -File $comScript
            Check ($LASTEXITCODE -eq 0) "case $i COM $exe"
        }
        Uninstall "uninstall-$i"
        Check (-not (Test-Path (Join-Path $app 'bin\net472\ExcelKit.dll'))) "case $i COM DLL removed"
    }
    Install 'reinstall-initial' 'com32,com64,manual,examples,net10' 'chinesetraditional'
    Set-Content (Join-Path $app 'Examples\user-created.txt') 'preserve'
    Install 'reinstall-reduced' 'com32,com64'
    foreach ($file in @('Manual\ExcelKit_User_Manual.pdf','Examples\README.md','bin\net10.0\ExcelKit.dll')) { Check (Test-Path (Join-Path $app $file)) "reinstall retained $file" }
    Uninstall 'uninstall-reinstall'
    Check (Test-Path (Join-Path $app 'Examples\user-created.txt')) 'uninstall retained user data'
    'PASS' | Set-Content (Join-Path $out 'result.txt')
} catch { ($_ | Out-String) | Set-Content (Join-Path $out 'result.txt'); throw } finally { Stop-Transcript | Out-Null }
