$ErrorActionPreference='Stop'
if($env:USERNAME -ne 'WDAGUtilityAccount'){throw 'Sandbox only'}
$out=Join-Path C:\SandboxOutput ('safety-'+(Get-Date -Format yyyyMMdd-HHmmss))
New-Item -ItemType Directory $out|Out-Null
$app='C:\ExcelKitSafetyTest'
$keys=@('ExcelKit.ExcelOperator','CLSID\{87654321-4321-4321-4321-210987654321}','Interface\{12345678-1234-1234-1234-123456789012}','TypeLib\{5F4B3D3E-6C2F-4F77-9D65-4B8A25E1B9A1}')
function Snapshot{
 $lines=@();foreach($v in '32','64'){foreach($k in $keys){$lines+="$v/$k";$prior=$ErrorActionPreference;$ErrorActionPreference='Continue';$text=& reg.exe query "HKLM\SOFTWARE\Classes\$k" /s "/reg:$v" 2>&1;$code=$LASTEXITCODE;$ErrorActionPreference=$prior;if($code -eq 0){$lines+=@($text|ForEach-Object{"$_"})}else{$lines+='ABSENT'}}};return ,@($lines|Sort-Object)
}
function Check([bool]$condition,[string]$name){if(-not $condition){throw "FAILED: $name"};"PASS: $name"|Tee-Object -FilePath (Join-Path $out 'checks.txt') -Append}

$ps="$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe"
$helper='C:\SandboxInput\ComTransaction.ps1'
$tx=Join-Path $env:ProgramData ('ExcelKit\Transactions\safety-'+[guid]::NewGuid())
function Call([string]$action,[string]$transaction){
 $prior=$ErrorActionPreference;$ErrorActionPreference='Continue'
 $extra=@();if($transaction){$extra=@('-Transaction',$transaction)}
 & $ps -NoProfile -ExecutionPolicy Bypass -File $helper -Action $action -AppDir $app -Architectures x86,x64 @extra *> (Join-Path $out ($action+'-'+[guid]::NewGuid()+'.log'))
 $code=$LASTEXITCODE;$ErrorActionPreference=$prior;return $code
}
try{
 $before=Snapshot
 Check ((Call 'Begin' $tx) -eq 0) 'backup created'
 Check ((Call 'Begin' ($tx+'-other')) -ne 0) 'unfinished transaction blocked'
 $file=Join-Path $tx 'snapshot.xml';$bytes=[IO.File]::ReadAllBytes($file)
 Add-Content -LiteralPath $file 'tampered'
 Check ((Call 'Apply' $tx) -ne 0) 'corrupt backup blocked before changes'
 Check (-not(Compare-Object $before (Snapshot))) 'corrupt backup left registry unchanged'
 [IO.File]::WriteAllBytes($file,$bytes)
 Check ((Call 'Rollback' $tx) -eq 0) 'interrupted prepared transaction recovered'
 Check (-not(Compare-Object $before (Snapshot))) 'recovery matches baseline'
 $key='HKLM:\SOFTWARE\Classes\CLSID\{87654321-4321-4321-4321-210987654321}\InprocServer32'
 New-Item -Path $key -Force|Out-Null;New-ItemProperty -Path $key -Name CodeBase -Value 'file:///C:/Unknown.dll'|Out-Null
 $unknown=Snapshot
 Check ((Call 'Inspect' '') -ne 0) 'unknown registration blocked'
 Check (-not(Compare-Object $unknown (Snapshot))) 'unknown registration untouched'
 Remove-Item 'HKLM:\SOFTWARE\Classes\CLSID\{87654321-4321-4321-4321-210987654321}' -Recurse
 $user='HKCU:\Software\Classes\ExcelKit.ExcelOperator';New-Item $user -Force|Out-Null
 Check ((Call 'Inspect' '') -ne 0) 'per-user override blocked'
 Check (Test-Path $user) 'per-user registration retained'
 Remove-Item $user -Recurse
 Check (-not(Compare-Object $before (Snapshot))) 'safety tests left COM baseline unchanged'
 'PASS'|Set-Content (Join-Path $out 'result.txt')
}catch{($_|Out-String)|Set-Content (Join-Path $out 'result.txt');throw}

