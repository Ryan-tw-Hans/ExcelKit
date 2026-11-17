$ErrorActionPreference='Stop'
if($env:USERNAME -ne 'WDAGUtilityAccount'){throw 'Sandbox only'}
$out=Join-Path C:\SandboxOutput ('examples-'+(Get-Date -Format yyyyMMdd-HHmmss));New-Item -ItemType Directory $out|Out-Null
Start-Transcript (Join-Path $out transcript.txt)|Out-Null
$app='C:\Program Files (x86)\ExcelKit'
function Check([bool]$ok,[string]$message){if(-not $ok){throw $message};"PASS: $message"|Tee-Object -FilePath (Join-Path $out checks.txt) -Append}
try{
 $p=Start-Process C:\SandboxInput\ExamplesInstaller.exe -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LANG=chinesetraditional /TASKS=com32,com64 /DIR=`"$app`" /LOG=`"$out\install.log`"" -Wait -PassThru
 Check ($p.ExitCode -eq 0) 'installer succeeded'
 foreach($category in @('Manual','Examples')){
  $root=Join-Path C:\Published $category
  foreach($file in Get-ChildItem -LiteralPath $root -File -Recurse){$relative=$file.FullName.Substring($root.Length).TrimStart('\');$target=Join-Path (Join-Path $app $category) $relative;Check ((Test-Path -LiteralPath $target) -and (Get-FileHash -LiteralPath $file.FullName).Hash -eq (Get-FileHash -LiteralPath $target).Hash) "$category/$relative installed with matching hash"}
 }
 foreach($view in @('x86','x64')){
  $exe=if($view -eq 'x86'){"$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"}else{"$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe"}
  $dest=Join-Path $out $view
  $result=& $exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $app 'Examples\PowerShell\QuickStart.ps1') -OutputDirectory $dest
  $result|Set-Content (Join-Path $out "example-$view.txt")
  Check ($LASTEXITCODE -eq 0 -and (Test-Path (Join-Path $dest 'ExcelKit-QuickStart.xlsx'))) "$view sample created workbook"
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip=[IO.Compression.ZipFile]::OpenRead((Join-Path $dest 'ExcelKit-QuickStart.xlsx'))
  try{$entry=$zip.GetEntry('xl/worksheets/sheet1.xml');Check ($null -ne $entry) "$view workbook has worksheet"}finally{$zip.Dispose()}
 }
 $shell=New-Object -ComObject WScript.Shell
 $linkRoot=Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\ExcelKit'
 $targets=@();foreach($file in Get-ChildItem -LiteralPath $linkRoot -Filter '*.lnk'){$targets+=$shell.CreateShortcut($file.FullName).TargetPath}
 Check ($targets -contains (Join-Path $app 'Manual')) 'manual folder shortcut points to installed manuals'
 Check ($targets -contains (Join-Path $app 'Examples')) 'example folder shortcut points to installed samples'
 Check ($targets -contains (Join-Path $app 'Manual\ExcelKit_User_Manual.pdf')) 'PDF shortcut retained'
 $sentinel=Join-Path $app 'Examples\user-created.txt';'preserve me'|Set-Content $sentinel
 $p=Start-Process (Join-Path $app unins000.exe) -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=`"$out\uninstall.log`"" -Wait -PassThru
 Check ($p.ExitCode -eq 0) 'uninstaller succeeded'
 Check (Test-Path $sentinel) 'user-created example file retained'
 Check (-not(Test-Path (Join-Path $app 'Manual\ExcelKit_User_Manual.pdf'))) 'installed PDF removed on uninstall'
 Check (-not(Test-Path (Join-Path $app 'Examples\PowerShell\QuickStart.ps1'))) 'installed sample removed on uninstall'
 'PASS'|Set-Content (Join-Path $out result.txt)
}catch{($_|Out-String)|Set-Content (Join-Path $out result.txt);throw}finally{Stop-Transcript|Out-Null}
