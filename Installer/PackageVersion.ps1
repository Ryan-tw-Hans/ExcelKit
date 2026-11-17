[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][ValidateSet('Suggest','Check','Reserve')][string]$Mode,
 [Parameter(Mandatory=$true)][string]$FileVersion,
 [Parameter(Mandatory=$true)][string]$Date,
 [string]$InstallerVersion='',
 [string]$Root=''
)
$ErrorActionPreference='Stop'
if(-not $Root){$Root=Split-Path -Parent $PSScriptRoot}
$root=[IO.Path]::GetFullPath($Root).TrimEnd('\')
if($FileVersion -notmatch '^\d+\.\d+\.\d+\.\d+$' -or ([Version]$FileVersion).ToString() -ne $FileVersion){throw 'Invalid DLL file version'}
$dt=[datetime]::MinValue
if(-not [datetime]::TryParseExact($Date,'yyyyMMdd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$dt)){throw 'Invalid package date'}
# Fail closed if the protection configured for the global namespace has changed.
try {$rule=Invoke-RestMethod -Uri 'https://api.github.com/repos/Ryan-tw-Hans/ExcelKit/rulesets/27431700' -Headers @{ 'User-Agent'='ExcelKit-package'; 'Accept'='application/vnd.github+json' } -TimeoutSec 30} catch {throw 'Unable to verify GitHub reservation protection'}
$types=@($rule.rules | ForEach-Object {$_.type})
if (([datetime]$rule.updated_at).ToUniversalTime() -ne ([datetime]'2026-11-17T04:26:25.021Z').ToUniversalTime()) {throw 'Reservation ruleset changed; review and update its trusted baseline'}
if($rule.target -ne 'tag' -or $rule.enforcement -ne 'active' -or @($rule.conditions.ref_name.exclude).Count -ne 0 -or @($rule.conditions.ref_name.include) -cnotcontains 'refs/tags/excelkit-package/**/*' -or $types -notcontains 'deletion' -or $types -notcontains 'update'){throw 'GitHub reservation protection is missing or changed'}
$remote=(& git -C $root remote get-url origin 2>$null);if($LASTEXITCODE -ne 0 -or $remote -notmatch '^https://github\.com/Ryan-tw-Hans/ExcelKit(?:\.git)?$'){throw 'Unexpected Git origin'}
$refs=@(& git -C $root ls-remote --refs origin 'refs/tags/excelkit-package/*' 2>$null);if($LASTEXITCODE -ne 0){throw 'Unable to read GitHub reservations'}
$max=0;$legacy=$false
foreach($line in $refs){
 if($line -notmatch '^[0-9a-f]{40}\s+refs/tags/excelkit-package/(?<day>\d{8})/(?<seq>\d{2,})$'){throw 'Unrecognized remote reservation'}
 $day=$Matches.day;$seqText=$Matches.seq;$seq=0
 if(-not [int]::TryParse($seqText,[ref]$seq) -or $seq -lt 1 -or $seq -gt 999999 -or $seqText -cne $seq.ToString('D2')){throw 'Invalid remote reservation'}
 if($day -eq '20261117' -and $seq -eq 1){$legacy=$true}
 if($day -eq $Date -and $seq -gt $max){$max=$seq}
}
if(-not $legacy){throw 'Historical reservation 20261117-01 is missing; migration required'}
# Existing local outputs and legacy reservations still raise the high-water mark.
$reserveDir=Join-Path $root '.tools\package-versions';$setup=Join-Path $root 'publish\Installer\Setup'
if(Test-Path -LiteralPath $reserveDir){foreach($f in Get-ChildItem -LiteralPath $reserveDir -File -Force){
 if($f.Name -notmatch '^(?<day>\d{8})-(?<seq>\d{2,})\.json$'){throw 'Unrecognized local reservation'}
 $day=$Matches.day;$seqText=$Matches.seq;$seq=0;if(-not [int]::TryParse($seqText,[ref]$seq) -or $seq -lt 1 -or $seq -gt 999999 -or $seqText -cne $seq.ToString('D2')){throw 'Invalid local reservation'}
 if($day -eq $Date -and $seq -gt $max){$max=$seq}
}}
if(Test-Path -LiteralPath $setup){foreach($f in Get-ChildItem -LiteralPath $setup -File -Force){
 if($f.Name -match '^ExcelKit-Setup-\d+\.\d+\.\d+\.\d+-(?<day>\d{8})-(?<seq>\d{2,})\.(?:exe|manifest\.json|build\.json)$'){
 $day=$Matches.day;$seqText=$Matches.seq;$seq=0;if(-not [int]::TryParse($seqText,[ref]$seq) -or $seq -lt 1 -or $seq -gt 999999 -or $seqText -cne $seq.ToString('D2')){throw 'Invalid local artifact sequence'}
 if($day -eq $Date -and $seq -gt $max){$max=$seq}
 }} }
if($max -ge 999999){throw 'Package sequence exhausted'}
$next=$max+1;$expected=$FileVersion+'-'+$Date+'-'+$next.ToString('D2')
if($Mode -eq 'Suggest'){Write-Output $expected;return}
if($InstallerVersion -cne $expected){throw "Package version is stale or conflicts; expected $expected"}
if($Mode -eq 'Check'){Write-Output $expected;return}
# Reservation records the exact commit. Ref creation is atomic; never overwrite or retry an uncertain result.
& git -C $root diff --quiet HEAD -- src Installer
if($LASTEXITCODE -ne 0){throw 'Commit source and installer changes before reserving a package version'}
$untracked=@(& git -C $root ls-files --others --exclude-standard -- src);if($LASTEXITCODE -ne 0 -or $untracked.Count -ne 0){throw 'Untracked source files are not allowed for a package reservation'}
$commit=(& git -C $root rev-parse --verify HEAD);if($LASTEXITCODE -ne 0 -or $commit -notmatch '^[0-9a-f]{40}$'){throw 'Unable to determine commit'}
$ref='refs/tags/excelkit-package/'+$Date+'/'+$next.ToString('D2')
$oldPrompt=$env:GIT_TERMINAL_PROMPT;$env:GIT_TERMINAL_PROMPT='0'
try {$oldError=$ErrorActionPreference;$ErrorActionPreference='Continue';& git -C $root push origin ($commit+':'+$ref) 2>&1 | Out-Null;$code=$LASTEXITCODE} finally {$ErrorActionPreference=$oldError;$env:GIT_TERMINAL_PROMPT=$oldPrompt}
if($code -ne 0){throw 'Reservation failed or result is unknown; inspect remote ref before retrying. Never reuse a reserved number.'}
$confirmed=@(& git -C $root ls-remote --refs origin $ref 2>$null)
if($LASTEXITCODE -ne 0 -or $confirmed.Count -ne 1 -or $confirmed[0] -notmatch ('^'+$commit+'\s+'+[regex]::Escape($ref)+'$')){throw 'Reservation result cannot be verified; inspect remote ref before retrying'}
Write-Output $expected
