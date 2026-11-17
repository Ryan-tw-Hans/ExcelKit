$ErrorActionPreference='Stop'
if($env:USERNAME -ne 'WDAGUtilityAccount'){throw 'Sandbox only'}
$out=Join-Path C:\SandboxOutput ('integration-'+(Get-Date -Format yyyyMMdd-HHmmss))
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
 if(Test-Path "$app\unins000.exe"){Uninstall "cleanup-prior-run"}
 $empty=Snapshot
 foreach($case in @(@('x86','com32'),@('x64','com64'),@('both','com32,com64'))){Install ('fresh-'+$case[0]) $case[1];if($case[0] -ne 'x64'){ComCheck 'x86'};if($case[0] -ne 'x86'){ComCheck 'x64'};Uninstall ('uninstall-'+$case[0]);Check (-not(Compare-Object $empty (Snapshot))) ('clean after '+$case[0])}
 $legacy=Join-Path $env:TEMP 'LegacyFixture';New-Item -ItemType Directory $legacy -Force|Out-Null
 $dll=Join-Path $legacy 'ExcelKit.dll';Copy-Item C:\SandboxInput\LegacyExcelKit.dll $dll
 foreach($view in 'Framework','Framework64'){& "$env:windir\Microsoft.NET\$view\v4.0.30319\RegAsm.exe" $dll /codebase "/tlb:$legacy\$view.tlb" /nologo;if($LASTEXITCODE -ne 0){throw 'legacy fixture registration failed'}}
 Remove-Item -LiteralPath "$legacy\Framework.tlb","$legacy\Framework64.tlb"
 $old=Snapshot
 Install 'decline-legacy' 'com32,com64' '/MIGRATELEGACY=no' $true
 Check (-not(Compare-Object $old (Snapshot))) 'decline leaves old registry'
 foreach($fault in 'after-remove','after-x86'){
  Install ('fault-'+$fault) 'com32,com64' "/MIGRATELEGACY=yes /TESTFAIL=$fault" $true
  Check (-not(Compare-Object $old (Snapshot))) ('rollback '+$fault)
  Uninstall ('cleanup-failed-'+$fault)
  Check (-not(Compare-Object $old (Snapshot))) ('failed setup cleanup preserved legacy '+$fault)
 }
 Install 'migrate' 'com32,com64' '/MIGRATELEGACY=yes'
 ComCheck x86;ComCheck x64
 Check (Test-Path -LiteralPath $dll) 'old DLL preserved'
 Install 'reinstall' 'com32,com64'
 ComCheck x86;ComCheck x64
 $p=Start-Process C:\SandboxInput\Installer-FaultTest.exe -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LANG=english /DIR=`"$app`" /TASKS=com32,com64 /LOG=`"$out\test-uninstaller.log`"" -Wait -PassThru
 Check ($p.ExitCode -eq 0) 'test-uninstaller installed'
 $owned=Snapshot
 Uninstall 'uninstall-fault' '/TESTFAIL=uninstall' $true
 Check (-not(Compare-Object $owned (Snapshot))) 'uninstall failure registry rollback'
 Check (Test-Path "$app\bin\net472\ExcelKit.dll") 'uninstall failure kept DLL'
 ComCheck x86;ComCheck x64
 Uninstall 'uninstall-final'
 Check (-not(Compare-Object $empty (Snapshot))) 'final COM registry clean'
 'PASS'|Set-Content (Join-Path $out 'result.txt')
}catch{($_|Out-String)|Set-Content (Join-Path $out 'result.txt');throw}finally{Stop-Transcript|Out-Null}
