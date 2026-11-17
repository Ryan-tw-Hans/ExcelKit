param(
 [Parameter(Mandatory=$true)][ValidateSet('Inspect','Begin','Apply','Finish','Rollback','Uninstall','UninstallCheck')][string]$Action,
 [Parameter(Mandatory=$true)][string]$AppDir,
 [string]$Architectures='x86,x64', [string]$Transaction='', [switch]$AllowLegacy,
 [string]$Fault=''
)
$ErrorActionPreference='Stop'
$AppDir=[IO.Path]::GetFullPath($AppDir).TrimEnd('\')
$ids=@('ExcelKit.ExcelOperator','CLSID\{87654321-4321-4321-4321-210987654321}','Interface\{12345678-1234-1234-1234-123456789012}','TypeLib\{5F4B3D3E-6C2F-4F77-9D65-4B8A25E1B9A1}')
$ownerPath='SOFTWARE\ExcelKitInstaller\{62953C31-A16D-469E-8866-6218DF2B3D2A}'
$views=@('x86'); if([Environment]::Is64BitOperatingSystem){$views+='x64'}
$selected=@($Architectures.Split(',') | Where-Object {$_})
if(-not $selected.Count -or @($selected | Where-Object {$_ -notin $views}).Count){throw 'Invalid COM architectures'}
if($Fault -and $env:USERNAME -ne 'WDAGUtilityAccount'){throw 'Fault injection requires Windows Sandbox'}
function Base([string]$view,[bool]$user=$false){
 $hive=if($user){[Microsoft.Win32.RegistryHive]::CurrentUser}else{[Microsoft.Win32.RegistryHive]::LocalMachine}
 $v=if($view -eq 'x86'){[Microsoft.Win32.RegistryView]::Registry32}else{[Microsoft.Win32.RegistryView]::Registry64}
 return [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,$v)
}
function Read-Tree($key){
 if($null -eq $key){return $null}
 $values=@();foreach($n in $key.GetValueNames()){$values+=@{Name=$n;Kind=[int]$key.GetValueKind($n);Data=$key.GetValue($n,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)}}
 $children=@();foreach($n in $key.GetSubKeyNames()){$k=$key.OpenSubKey($n);try{$children+=@{Name=$n;Tree=(Read-Tree $k)}}finally{$k.Dispose()}}
 return @{Values=$values;Children=$children;Dacl=$key.GetAccessControl([Security.AccessControl.AccessControlSections]::Access).GetSecurityDescriptorSddlForm([Security.AccessControl.AccessControlSections]::Access)}
}
function Write-Tree($key,$tree){foreach($v in $tree.Values){$key.SetValue([string]$v.Name,$v.Data,[Microsoft.Win32.RegistryValueKind][int]$v.Kind)};foreach($c in $tree.Children){$k=$key.CreateSubKey([string]$c.Name);try{Write-Tree $k $c.Tree}finally{$k.Dispose()}};if($tree.Dacl){$security=[Security.AccessControl.RegistrySecurity]::new();$security.SetSecurityDescriptorSddlForm([string]$tree.Dacl,[Security.AccessControl.AccessControlSections]::Access);$key.SetAccessControl($security)}}
function Get-Snapshot{
 $snap=@();foreach($view in $views){$b=Base $view;try{foreach($path in (@($ids|ForEach-Object{'SOFTWARE\Classes\'+$_})+@($ownerPath))){$k=$b.OpenSubKey($path);try{$snap+=@{View=$view;Path=$path;Tree=(Read-Tree $k)}}finally{if($k){$k.Dispose()}}}}finally{$b.Dispose()}};return ,$snap
}
function Restore($snap){
 foreach($s in $snap){$b=Base $s.View;try{$b.DeleteSubKeyTree([string]$s.Path,$false);if($null -ne $s.Tree){$k=$b.CreateSubKey([string]$s.Path);try{Write-Tree $k $s.Tree}finally{$k.Dispose()}}}finally{$b.Dispose()}}
}
function Flatten-Tree($tree,[string]$prefix){
 if($null -eq $tree){"$prefix|ABSENT";return}
 "$prefix|KEY|$($tree.Dacl)"
 foreach($v in @($tree.Values|Sort-Object Name)){"$prefix|VALUE|$($v.Name)|$($v.Kind)|$($v.Data|ConvertTo-Json -Compress -Depth 10)"}
 foreach($c in @($tree.Children|Sort-Object Name)){Flatten-Tree $c.Tree ($prefix+'\'+$c.Name)}
}
function Canonical($snap){return ((@($snap|Sort-Object View,Path|ForEach-Object{Flatten-Tree $_.Tree ($_.View+'|'+$_.Path)})) -join "`n")}
function Inspect{
 $states=@();foreach($view in $views){
  $u=Base $view $true;try{foreach($id in $ids){$k=$u.OpenSubKey('SOFTWARE\Classes\'+$id);if($k){$k.Dispose();throw 'Per-user COM overrides require manual review'}}}finally{$u.Dispose()}
  $b=Base $view;try{
   $p=$b.OpenSubKey('SOFTWARE\Classes\'+$ids[0]+'\CLSID');$c=$b.OpenSubKey('SOFTWARE\Classes\'+$ids[1]+'\InprocServer32');$ob=Base 'x86';try{$o=$ob.OpenSubKey($ownerPath)}finally{$ob.Dispose()}
   try{
    if(-not $c){if($p -and $p.GetValue('') -ne '{87654321-4321-4321-4321-210987654321}'){throw 'Unknown ProgID'};$z=$b.OpenSubKey('SOFTWARE\Classes\'+$ids[1]);if($z){$z.Dispose();throw 'Incomplete CLSID'};$states+=@{View=$view;Kind='Empty'};continue}
    if(-not $p -or -not $c -or $p.GetValue('') -ne '{87654321-4321-4321-4321-210987654321}'){throw 'Unknown or incomplete COM registration'}
    $url=[Uri][string]$c.GetValue('CodeBase');if(-not $url.IsFile -or $url.IsUnc){throw 'COM DLL must be a local file'}
    $dll=$url.LocalPath
    if(-not(Test-Path -LiteralPath $dll -PathType Leaf)){throw 'Existing COM DLL is missing'}
    $assembly=[string]$c.GetValue('Assembly')
    $own=$o -and ([string]$o.GetValue($view) -eq $AppDir) -and ($dll -eq (Join-Path $AppDir 'bin\net472\ExcelKit.dll'))
    if($own){$kind='Owned'}elseif((Get-FileHash -LiteralPath $dll -Algorithm SHA256).Hash -eq 'B2C2FC06817EA615CED43C90D910F259E516E4A10BEE3941DC701EB1339FA84E' -and $assembly -eq 'ExcelKit, Version=1.2.0.0, Culture=neutral, PublicKeyToken=null'){$kind='Legacy'}else{throw "Unrecognized COM DLL: $dll"}
    $versions=@($c.GetSubKeyNames());if($versions.Count -ne 1){throw 'Multiple COM assembly versions require manual review'}
    $typelib=$b.OpenSubKey('SOFTWARE\Classes\'+$ids[3]);if($typelib){try{if(@($typelib.GetSubKeyNames()|Where-Object{$_ -ne '1.2'}).Count){throw 'Shared typelib has other versions'}}finally{$typelib.Dispose()}}
    $iface=$b.OpenSubKey('SOFTWARE\Classes\'+$ids[2]+'\TypeLib')
    if($iface){try{if($iface.GetValue('') -ne '{5F4B3D3E-6C2F-4F77-9D65-4B8A25E1B9A1}' -or $iface.GetValue('Version') -ne '1.2'){throw 'Interface belongs to another typelib'}}finally{$iface.Dispose()}}
    if($view -notin $selected){throw "Keep existing $view selected"}
    $states+=@{View=$view;Kind=$kind;Dll=$dll;Assembly=$assembly}
   }finally{foreach($k in @($p,$c,$o)){if($k){$k.Dispose()}}}
  }finally{$b.Dispose()}
 }; return ,$states
}
function Regasm([string]$view,[string]$dll,[bool]$remove){
 $f=if($view -eq 'x86'){'Framework'}else{'Framework64'}
 $exe=Join-Path $env:WINDIR "Microsoft.NET\$f\v4.0.30319\RegAsm.exe"
 if($remove){$args=@($dll,'/unregister','/nologo')}else{$tlb=Join-Path $AppDir "bin\net472\ExcelKit-$view.tlb";$args=@($dll,'/codebase',"/tlb:$tlb",'/nologo')}
 & $exe @args | Out-String | Add-Content -LiteralPath (Join-Path $Transaction 'regasm.log')
 if($LASTEXITCODE -ne 0){throw "RegAsm failed ($view/$remove): $LASTEXITCODE"}
}
function Clear-Com{
 foreach($v in $views){$b=Base $v;try{foreach($id in $ids){$b.DeleteSubKeyTree('SOFTWARE\Classes\'+$id,$false)}}finally{$b.Dispose()}}
}
function Load-Journal{
 if(-not $Transaction.StartsWith((Join-Path $env:ProgramData 'ExcelKit\Transactions\'),[StringComparison]::OrdinalIgnoreCase)){throw 'Invalid transaction path'}
 $file=Join-Path $Transaction 'snapshot.xml';$expected=(Get-Content -LiteralPath (Join-Path $Transaction 'snapshot.sha256') -Raw).Trim()
 if((Get-FileHash -LiteralPath $file).Hash -ne $expected){throw 'Backup checksum mismatch'}
 $script:journal=Import-Clixml -LiteralPath $file
 if($journal.AppDir -ne $AppDir){throw 'Backup belongs to another installation'}
}
function Rollback{
 Load-Journal
 if((Get-Content -LiteralPath (Join-Path $Transaction 'state.txt') -Raw).Trim() -eq 'complete'){throw 'Completed transactions cannot be rolled back'}
 foreach($f in $journal.Files){if((Get-FileHash -LiteralPath (Join-Path $Transaction $f.Backup)).Hash -ne $f.Hash){throw 'DLL backup checksum mismatch'}}
 Restore $journal.Registry
 if((Canonical (Get-Snapshot)) -ne (Canonical $journal.Registry)){throw 'Registry rollback verification failed'}
 foreach($f in $journal.Files){[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName([string]$f.Target))|Out-Null;Copy-Item -LiteralPath (Join-Path $Transaction $f.Backup) -Destination $f.Target -Force}
 Set-Content -LiteralPath (Join-Path $Transaction 'state.txt') 'rolledback'
}
$mutex=[Threading.Mutex]::new($false,'Global\ExcelKitInstallerCOM');$locked=$false
try{
 $locked=$mutex.WaitOne(0);if(-not $locked){throw 'Another ExcelKit COM operation is running'}
 if(-not [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'Administrator privileges required'}
 switch($Action){
 'Inspect' { $states=Inspect;$states|ConvertTo-Json -Depth 5; if(@($states|Where-Object Kind -eq 'Legacy').Count){exit 10};exit 0 }
 'Begin' {
  $states=Inspect;if(@($states|Where-Object Kind -eq 'Legacy').Count -and -not $AllowLegacy){throw 'Legacy replacement consent required'}
  $parent=Join-Path $env:ProgramData 'ExcelKit'
  $root=Join-Path $parent 'Transactions'
  foreach($dir in @($parent,$root)){if((Test-Path -LiteralPath $dir) -and ((Get-Item -LiteralPath $dir -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Backup path cannot be a link'}}
  [IO.Directory]::CreateDirectory($root)|Out-Null
  foreach($dir in @($parent,$root)){
   $acl=[Security.AccessControl.DirectorySecurity]::new();$acl.SetAccessRuleProtection($true,$false)
   $admins=[Security.Principal.SecurityIdentifier]::new('S-1-5-32-544');$system=[Security.Principal.SecurityIdentifier]::new('S-1-5-18');$acl.SetOwner($admins)
   foreach($sid in @($admins,$system)){$rule=[Security.AccessControl.FileSystemAccessRule]::new($sid,'FullControl','ContainerInherit,ObjectInherit','None','Allow');$acl.AddAccessRule($rule)}
   [IO.Directory]::SetAccessControl($dir,$acl)
  }
  if(-not $Transaction.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Invalid transaction directory'}
  if(Test-Path -LiteralPath $Transaction){throw 'Transaction already exists'}
  foreach($prior in Get-ChildItem -LiteralPath $root -Directory){$stateFile=Join-Path $prior.FullName 'state.txt';if(Test-Path -LiteralPath $stateFile){$state=(Get-Content -LiteralPath $stateFile -Raw).Trim();if($state -in @('prepared','applying','applied')){throw ('Unfinished COM transaction requires recovery: '+$prior.FullName)}}}
  New-Item -ItemType Directory -Path $Transaction|Out-Null
  $files=@();$i=0
  if(@($states|Where-Object Kind -eq 'Owned').Count){foreach($f in Get-ChildItem -LiteralPath (Join-Path $AppDir 'bin\net472') -File){$name="file-$i.bak";Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $Transaction $name);$files+=@{Target=$f.FullName;Backup=$name;Hash=(Get-FileHash -LiteralPath $f.FullName).Hash};$i++}}
  $journal=@{AppDir=$AppDir;Registry=(Get-Snapshot);States=$states;Files=$files;Selected=$selected}
  $file=Join-Path $Transaction 'snapshot.xml';$journal|Export-Clixml -LiteralPath $file -Depth 100
  (Get-FileHash -LiteralPath $file).Hash|Set-Content -LiteralPath (Join-Path $Transaction 'snapshot.sha256')
  Set-Content -LiteralPath (Join-Path $Transaction 'state.txt') 'prepared'
 }
 'Apply' {
  Load-Journal
  if((Canonical (Get-Snapshot)) -ne (Canonical $journal.Registry)){throw 'COM changed after preflight'}
  try{
   Set-Content -LiteralPath (Join-Path $Transaction 'state.txt') 'applying'
   foreach($s in $journal.States){if($s.Kind -eq 'Legacy'){Regasm $s.View $s.Dll $true}}
   if(@($journal.States|Where-Object Kind -eq 'Legacy').Count){Clear-Com}
   if($Fault -eq 'after-remove'){throw 'Injected failure after old removal'}
   foreach($v in $journal.Selected){Regasm $v (Join-Path $AppDir 'bin\net472\ExcelKit.dll') $false;if($Fault -eq 'after-x86' -and $v -eq 'x86'){throw 'Injected failure after x86 registration'}}
   $b=Base 'x86';try{$o=$b.CreateSubKey($ownerPath);foreach($v in $journal.Selected){$o.SetValue($v,$AppDir)};$o.Dispose()}finally{$b.Dispose()}
   Set-Content -LiteralPath (Join-Path $Transaction 'state.txt') 'applied'
  }catch{$errorText=$_|Out-String;Rollback;throw "COM operation failed; previous registry restored. $errorText"}
 }
 'Finish' {Load-Journal;Set-Content -LiteralPath (Join-Path $Transaction 'state.txt') 'complete'}
 'Rollback' {Rollback}
 'UninstallCheck' {
  $states=Inspect
  if(@($states|Where-Object Kind -eq 'Owned').Count){exit 10}
  if(@($states|Where-Object {$_.Dll -eq (Join-Path $AppDir 'bin\net472\ExcelKit.dll')}).Count){throw 'Unowned COM DLL is inside installation folder'}
  exit 0
 }
 'Uninstall' {
  Load-Journal
  $states=Inspect;if(@($states|Where-Object Kind -eq 'Legacy').Count){throw 'Uninstaller does not own this registration'}
  $snapshot=$journal.Registry
  try{
   foreach($s in $states){if($s.Kind -eq 'Owned'){Regasm $s.View $s.Dll $true}}
   if($Fault -eq 'uninstall'){throw 'Injected uninstall failure'}
   Clear-Com
   $b=Base 'x86';try{$b.DeleteSubKeyTree($ownerPath,$false)}finally{$b.Dispose()}
   Set-Content -LiteralPath (Join-Path $Transaction 'state.txt') 'complete'
  }catch{Rollback;throw}
 }
 }
 exit 0
}catch{if($Transaction -and (Test-Path -LiteralPath $Transaction)){($_|Out-String)|Add-Content -LiteralPath (Join-Path $Transaction 'error.txt')};Write-Error ($_|Out-String) -ErrorAction Continue;exit 1}finally{if($locked){$mutex.ReleaseMutex()};$mutex.Dispose()}


