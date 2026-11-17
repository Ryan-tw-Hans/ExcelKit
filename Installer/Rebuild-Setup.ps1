[CmdletBinding()]
param([string]$CompilerPath = '', [string]$InstallerVersion = '', [switch]$Preflight)
$ErrorActionPreference = 'Stop'
$localRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$configRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
$root = if(Test-Path -LiteralPath (Join-Path $localRoot 'Installer\Package-ExcelKit.ps1')){$localRoot}else{$configRoot}
$package=Join-Path $root 'Installer\Package-ExcelKit.ps1'
if(-not(Test-Path -LiteralPath $package -PathType Leaf)){throw '找不到 Installer\Package-ExcelKit.ps1；請從完整專案執行。'}
& $package -CompilerPath $CompilerPath -InstallerVersion $InstallerVersion -Preflight:$Preflight