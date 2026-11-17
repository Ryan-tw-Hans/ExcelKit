$ErrorActionPreference='Stop'
if($env:USERNAME -ne 'WDAGUtilityAccount'){throw 'Sandbox only'}
$out=Join-Path C:\SandboxOutput ('upgrade-'+(Get-Date -Format yyyyMMdd-HHmmss))
New-Item -ItemType Directory $out|Out-Null
Start-Transcript (Join-Path $out 'transcript.txt')|Out-Null
$app='C:\Program Files (x86)\ExcelKit'
$keys=@('ExcelKit.ExcelOperator','CLSID\{87654321-4321-4321-4321-210987654321}','Interface\{12345678-1234-1234-1234-123456789012}','TypeLib\{5F4B3D3E-6C2F-4F77-9D65-4B8A25E1B9A1}')
function Snapshot{
 $lines=@();foreach($v in '32','64'){foreach($k in $keys){$lines+="$v/$k";$prior=$ErrorActionPreference;$ErrorActionPreference='Continue';$text=& reg.exe query "HKLM\SOFTWARE\Classes\$k" /s "/reg:$v" 2>&1;$code=$LASTEXITCODE;$ErrorActionPreference=$prior;if($code -eq 0){$lines+=@($text|ForEach-Object{"$_"})}else{$lines+='ABSENT'}}};return ,@($lines|Sort-Object)
}
function Check([bool]$condition,[string]$name){if(-not $condition){throw "FAILED: $name"};"PASS: $name"|Tee-Object -FilePath (Join-Path $out 'checks.txt') -Append}
function Install([string]$name,[string]$tasks,[string]$extra='', [bool]$fail=$false){
 $exe=if($extra -match 'TESTFAIL'){'C:\SandboxInput\Installer-FaultTest.exe'}else{'C:\SandboxInput\Installer.exe'}
 $p=Start-Process $exe -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /LANG=english /DIR=`"$app`" /TASKS=$tasks /LOG=`"$out\$name.log`" $extra" -Wait -PassThru
 Check (($p.ExitCode -ne 0) -eq $fail) "$name exit=$($p.ExitCode)"
}
function Uninstall([string]$name,[string]$extra='', [bool]$fail=$false){
 $p=Start-Process (Join-Path $app 'unins000.exe') -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=`"$out\$name.log`" $extra" -Wait -PassThru
 Check (($p.ExitCode -ne 0) -eq $fail) "$name exit=$($p.ExitCode)"
}
function ComCheck([string]$view){
 $script=Join-Path $out "com-$view.ps1"
 @('$ErrorActionPreference="Stop"','$o=New-Object -ComObject ExcelKit.ExcelOperator','$o.CreateNewWorkbook()','$o.SetCellValue("Sheet1","A1","smoke")','if($o.GetCellValue("Sheet1","A1") -ne "smoke"){exit 2}','$o.GetVersionInfo()') | Set-Content -LiteralPath $script -Encoding ASCII
 $exe=if($view -eq 'x86'){"$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"}else{"$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe"}
 $value=& $exe -NoProfile -ExecutionPolicy Bypass -File $script
 Check ($LASTEXITCODE -eq 0 -and "$value" -like '*1.2.4+20261115075927*') "COM $view $value"
}

try{
 $empty=Snapshot
 $p=Start-Process C:\SandboxInput\PreviousInstaller.exe -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LANG=english /DIR=`"$app`" /TASKS=com32,com64 /LOG=`"$out\previous.log`"" -Wait -PassThru
 Check ($p.ExitCode -eq 0) 'previous version installed'
 $before=Snapshot;$dll=Join-Path $app 'bin\net472\ExcelKit.dll';$hash=(Get-FileHash $dll).Hash
 Install 'upgrade-fault' 'com32,com64' '/TESTFAIL=after-x86' $true
 Check (-not(Compare-Object $before (Snapshot))) 'upgrade fault registry restored'
 Check ((Get-FileHash $dll).Hash -eq $hash) 'upgrade fault original DLL restored'
 Install 'upgrade-success' 'com32,com64'
 ComCheck x86;ComCheck x64
 Uninstall 'upgrade-uninstall'
 Check (-not(Compare-Object $empty (Snapshot))) 'upgrade uninstall cleaned COM'
 'PASS'|Set-Content (Join-Path $out 'result.txt')
}catch{($_|Out-String)|Set-Content (Join-Path $out 'result.txt');throw}finally{Stop-Transcript|Out-Null}
