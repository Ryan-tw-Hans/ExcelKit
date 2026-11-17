#include "PackageMetadata.iss"
#if LowerCase(GetSHA256OfFile(AddBackslash(SourcePath) + "..\..\bin\net472\ExcelKit.dll")) != LowerCase(PackageAssemblyHash)
  #error Packaged DLL does not match PackageMetadata.iss; run Rebuild-Setup.ps1
#endif
#define PackageManifestHash GetSHA256OfFile(AddBackslash(SourcePath) + "InstallManifest.json")
#define ProductVersion "1.2.4+20261116121333"
[Setup]
AppId={{62953C31-A16D-469E-8866-6218DF2B3D2A}
AppName=ExcelKit
AppVersion={#ProductVersion}
AppPublisher=ExcelKit
DefaultDirName={autopf}\ExcelKit
DefaultGroupName=ExcelKit
PrivilegesRequired=admin
OutputDir=..\Setup
OutputBaseFilename=ExcelKit-Setup-1.2.4-20261116121333-TEST
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
DisableWelcomePage=no
DisableProgramGroupPage=yes
UsePreviousTasks=yes
LanguageDetectionMethod=uilanguage
UsePreviousLanguage=no
ShowLanguageDialog=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesetraditional"; MessagesFile: "compiler:Languages\ChineseTraditional.isl"

[CustomMessages]
english.Com32=Register COM for 32-bit (x86) applications
chinesetraditional.Com32=註冊供 32 位元（x86）應用程式使用的 COM 元件
english.Com64=Register COM for 64-bit (x64) applications
chinesetraditional.Com64=註冊供 64 位元（x64）應用程式使用的 COM 元件
english.ComGroup=Choose COM client architecture:
chinesetraditional.ComGroup=請選擇使用 COM 元件的應用程式位元數：
english.Manual=ExcelKit Manual
chinesetraditional.Manual=ExcelKit 使用手冊
english.FrameworkRequired=Microsoft .NET Framework 4.7.2 or later is required.
chinesetraditional.FrameworkRequired=需要 Microsoft .NET Framework 4.7.2 或更新版本。
english.SelectArchitecture=Select at least one COM client architecture.
chinesetraditional.SelectArchitecture=請至少選擇一種 COM 應用程式位元數。
english.Regasm32Missing=32-bit RegAsm.exe is missing. Installation stopped.
chinesetraditional.Regasm32Missing=找不到 32 位元 RegAsm.exe，安裝已停止。
english.Regasm64Missing=64-bit RegAsm.exe is missing. Installation stopped.
chinesetraditional.Regasm64Missing=找不到 64 位元 RegAsm.exe，安裝已停止。
english.Conflict32=An existing 32-bit ExcelKit COM registration belongs to another installation or cannot be verified. Setup stopped to protect it. Contact the administrator before changing the existing registration.
chinesetraditional.Conflict32=偵測到既有的 32 位元 ExcelKit COM 註冊屬於其他安裝或無法確認來源。為保護現有設定，安裝已停止。變更既有註冊前，請聯絡管理員確認。
english.Conflict64=An existing 64-bit ExcelKit COM registration belongs to another installation or cannot be verified. Setup stopped to protect it. Contact the administrator before changing the existing registration.
chinesetraditional.Conflict64=偵測到既有的 64 位元 ExcelKit COM 註冊屬於其他安裝或無法確認來源。為保護現有設定，安裝已停止。變更既有註冊前，請聯絡管理員確認。
english.KeepArchitectures=Keep all previously installed COM architectures selected when upgrading. To remove one, uninstall first.
chinesetraditional.KeepArchitectures=升級時請保留先前已安裝的所有 COM 位元數。若要移除其中一項，請先解除安裝。
english.Register32Failed=32-bit COM registration failed. Installation cannot complete.
chinesetraditional.Register32Failed=32 位元 COM 註冊失敗，無法完成安裝。
english.Register64Failed=64-bit COM registration failed. Installation cannot complete. Check the setup log and COM registration.
chinesetraditional.Register64Failed=64 位元 COM 註冊失敗，無法完成安裝。請檢查安裝記錄與 COM 註冊狀態。
english.Owner32Failed=Could not save 32-bit ownership information. Check setup log.
chinesetraditional.Owner32Failed=無法儲存 32 位元安裝擁有者資訊。請檢查安裝記錄。
english.Owner64Failed=Could not save 64-bit ownership information. Check setup log.
chinesetraditional.Owner64Failed=無法儲存 64 位元安裝擁有者資訊。請檢查安裝記錄。
english.Unregister32Failed=Could not unregister 32-bit COM. Uninstallation is stopping.
chinesetraditional.Unregister32Failed=無法解除 32 位元 COM 註冊，解除安裝已停止。
english.Unregister64Failed=Could not unregister 64-bit COM. Uninstallation is stopping.
chinesetraditional.Unregister64Failed=無法解除 64 位元 COM 註冊，解除安裝已停止。

english.ExistingDetail=%1: %2%nSource: %3
chinesetraditional.ExistingDetail=%1：%2%n來源：%3
english.UnknownValue=Unknown
chinesetraditional.UnknownValue=無法確認
english.LegacySafety=This test installer cannot safely remove a previous registration yet. No registration has been changed.
chinesetraditional.LegacySafety=此測試版尚無法安全移除舊註冊，未變更任何註冊設定。
english.View32Machine=32-bit system registration
chinesetraditional.View32Machine=32 位元系統註冊
english.View64Machine=64-bit system registration
chinesetraditional.View64Machine=64 位元系統註冊
english.View32User=32-bit user registration
chinesetraditional.View32User=32 位元使用者註冊
english.View64User=64-bit user registration
chinesetraditional.View64User=64 位元使用者註冊

english.MigratePrompt=Remove the old ExcelKit COM registration and install this version? Other applications using the old COM registration will use this version. Old DLL files are kept. A backup will be saved before changes.
chinesetraditional.MigratePrompt=是否先移除舊版 ExcelKit COM 註冊再安裝新版？使用舊 COM 的其他程式將改用新版。舊 DLL 檔案會保留，變更前會保存備份。
english.OperationFailed=COM operation failed. See the transaction folder for backup and logs:
chinesetraditional.OperationFailed=COM 作業失敗。請查看以下交易資料夾的備份及記錄：
english.ConsentNeeded=Legacy COM replacement was declined. No COM changes were made.
chinesetraditional.ConsentNeeded=未同意取代舊版 COM 註冊，未變更 COM 設定。

english.ManualFolder=Operation manuals folder
chinesetraditional.ManualFolder=操作手冊資料夾
english.ExampleFolder=Example programs folder
chinesetraditional.ExampleFolder=範例程式資料夾
english.OptionalGroup=Choose optional content:
chinesetraditional.OptionalGroup=請選擇要安裝的附加內容：
english.InstallManual=Operation manuals and media
chinesetraditional.InstallManual=操作手冊及媒體
english.InstallExamples=Example programs
chinesetraditional.InstallExamples=範例程式
english.InstallNet10=.NET 10.0 components
chinesetraditional.InstallNet10=.NET 10.0 元件

english.OpenInstallFolder=Open the ExcelKit installation folder
chinesetraditional.OpenInstallFolder=開啟 ExcelKit 安裝資料夾
english.LibreOfficeInstall=Open the LibreOffice installer download page
chinesetraditional.LibreOfficeInstall=開啟 LibreOffice 安裝版下載頁
english.LibreOfficePortable=Open the LibreOffice portable download page
chinesetraditional.LibreOfficePortable=開啟 LibreOffice 免安裝版下載頁

english.StartupUserOverride=Per-user override detected; status requires review.
chinesetraditional.StartupUserOverride=偵測到使用者層覆寫，狀態需人工確認。
english.StartupNotFound=No complete machine registration found, or unreadable.
chinesetraditional.StartupNotFound=找不到完整系統註冊，或無法讀取。
english.StartupUnknown=Unknown or unreadable
chinesetraditional.StartupUnknown=無法確認或讀取
english.StartupMultiple=Multiple or incomplete registration versions; review required.
chinesetraditional.StartupMultiple=註冊版本多筆或不完整，需人工確認。
english.StartupFileVersion=Currently registered DLL file version:
chinesetraditional.StartupFileVersion=目前已登錄 DLL 檔案版本：
english.StartupMissingOrUnreadable=DLL path invalid, missing or unreadable; content unknown.
chinesetraditional.StartupMissingOrUnreadable=DLL 路徑無效、檔案遺失或無法讀取；內容未知。
english.StartupHashSame=DLL SHA-256 matches packaged DLL.
chinesetraditional.StartupHashSame=DLL SHA-256 與套件一致。
english.StartupHashDifferent=DLL SHA-256 differs from packaged DLL.
chinesetraditional.StartupHashDifferent=DLL SHA-256 與套件不同。
english.StartupHashUnknown=DLL SHA-256 could not be checked.
chinesetraditional.StartupHashUnknown=無法檢查 DLL SHA-256。
english.StartupFolderUnknown=Installed ExcelKit folder not found or cannot be verified.
chinesetraditional.StartupFolderUnknown=未偵測到可確認的 ExcelKit 安裝資料夾。
english.StartupOpenFailed=Could not open the selected location.
chinesetraditional.StartupOpenFailed=無法開啟選擇的位置。
english.StartupInstallerVersion=This installer:
chinesetraditional.StartupInstallerVersion=本安裝程式版本：
english.StartupInstalledSetupVersion=Installed installer version:
chinesetraditional.StartupInstalledSetupVersion=已安裝的安裝程式版本：
english.StartupPackageComponent=Packaged DLL product version:
chinesetraditional.StartupPackageComponent=套件 DLL 產品版本：
english.StartupPackageFile=Packaged DLL file version:
chinesetraditional.StartupPackageFile=套件 DLL 檔案版本：

[Tasks]
Name: "com32"; Description: "{cm:Com32}"; GroupDescription: "{cm:ComGroup}"
Name: "com64"; Description: "{cm:Com64}"; GroupDescription: "{cm:ComGroup}"; Check: IsWin64
Name: "manual"; Description: "{cm:InstallManual}"; GroupDescription: "{cm:OptionalGroup}"; Flags: unchecked
Name: "examples"; Description: "{cm:InstallExamples}"; GroupDescription: "{cm:OptionalGroup}"; Flags: unchecked
Name: "net10"; Description: "{cm:InstallNet10}"; GroupDescription: "{cm:OptionalGroup}"; Flags: unchecked

[Files]
Source: "LibreOffice-下載連結.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "InstallManifest.json"; DestDir: "{app}\Installer"; Flags: ignoreversion
Source: "Verify-ExcelKit.ps1"; DestDir: "{app}\Installer"; Flags: ignoreversion
Source: "ComTransaction.ps1"; Flags: dontcopy
Source: "ComTransaction.ps1"; DestDir: "{app}\Installer"; Flags: ignoreversion

Source: "..\..\bin\net472\*"; DestDir: "{app}\bin\net472"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\..\bin\net10.0\*"; DestDir: "{app}\bin\net10.0"; Tasks: net10; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\..\Manual\*"; DestDir: "{app}\Manual"; Tasks: manual; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\..\Examples\*"; DestDir: "{app}\Examples"; Tasks: examples; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\..\Media\*"; DestDir: "{app}\Media"; Tasks: manual; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{cm:ManualFolder}"; Filename: "{app}\Manual"; Tasks: manual
Name: "{group}\{cm:ExampleFolder}"; Filename: "{app}\Examples"; Tasks: examples
Name: "{group}\{cm:Manual}"; Filename: "{app}\Manual\ExcelKit_User_Manual.pdf"; Tasks: manual

[Run]
Filename: "{app}"; Description: "{cm:OpenInstallFolder}"; Flags: shellexec postinstall unchecked skipifsilent; Check: CanOfferDownloads
Filename: "https://www.libreoffice.org/download/"; Description: "{cm:LibreOfficeInstall}"; Flags: shellexec postinstall unchecked skipifsilent; Check: CanOfferDownloads
Filename: "https://portableapps.com/apps/office/libreoffice_portable"; Description: "{cm:LibreOfficePortable}"; Flags: shellexec postinstall unchecked skipifsilent; Check: CanOfferDownloads

[UninstallDelete]
Type: files; Name: "{app}\Installer\install-selection.txt"
Type: files; Name: "{app}\bin\net472\ExcelKit-x86.tlb"
Type: files; Name: "{app}\bin\net472\ExcelKit-x64.tlb"

[Code]
var Transaction, Architectures: String;
    Prepared, Finished, Consent, ApplyFailed: Boolean;

#include "Startup.iss"

function CanOfferDownloads(): Boolean;
begin
 Result := not ApplyFailed;
end;

function HasRequiredFramework(): Boolean;
var Release: Cardinal;
begin
 Result := RegQueryDWordValue(HKLM32, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release);
 if Result then Result := Release >= 461808;
end;

function PSPath(): String;
begin
 if IsWin64 then Result := ExpandConstant('{sysnative}\WindowsPowerShell\v1.0\powershell.exe')
 else Result := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');
end;

function InvokeHelper(Action, HelperPath: String): Integer;
var Params, Fault: String; Code: Integer;
begin
 Params := '-NoProfile -ExecutionPolicy Bypass -File "' + HelperPath + '" -Action ' + Action +
   ' -AppDir "' + ExpandConstant('{app}') + '" -Architectures "' + Architectures + '"';
 if Transaction <> '' then Params := Params + ' -Transaction "' + Transaction + '"';
 if Consent then Params := Params + ' -AllowLegacy';
 #ifdef TestBuild
 Fault := ExpandConstant('{param:TESTFAIL|}');
 Log('Fault injection requested=' + Fault);
 if Fault <> '' then Params := Params + ' -Fault "' + Fault + '"';
 #endif
 if Exec(PSPath(), Params, '', SW_HIDE, ewWaitUntilTerminated, Code) then Result := Code else Result := -1;
 Log('COM helper ' + Action + ' exit=' + IntToStr(Result));
end;

function ExistingDetails(): String;
var AssemblyName, Source: String;
begin
 Result := '';
 if RegQueryStringValue(HKLM32, 'Software\Classes\CLSID\{87654321-4321-4321-4321-210987654321}\InprocServer32', 'Assembly', AssemblyName) then begin
   RegQueryStringValue(HKLM32, 'Software\Classes\CLSID\{87654321-4321-4321-4321-210987654321}\InprocServer32', 'CodeBase', Source);
   Result := Result + #13#10#13#10 + 'x86: ' + AssemblyName + #13#10 + Source;
 end;
 if IsWin64 then begin
   if RegQueryStringValue(HKLM64, 'Software\Classes\CLSID\{87654321-4321-4321-4321-210987654321}\InprocServer32', 'Assembly', AssemblyName) then begin
     RegQueryStringValue(HKLM64, 'Software\Classes\CLSID\{87654321-4321-4321-4321-210987654321}\InprocServer32', 'CodeBase', Source);
     Result := Result + #13#10#13#10 + 'x64: ' + AssemblyName + #13#10 + Source;
   end;
 end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var Status: Integer; HelperPath: String;
begin
 Result := '';
 if not HasRequiredFramework() then begin Result := CustomMessage('FrameworkRequired'); Exit; end;
 if Prepared then Exit;
 Architectures := '';
 if WizardIsTaskSelected('com32') then Architectures := 'x86';
 if WizardIsTaskSelected('com64') then begin
   if Architectures <> '' then Architectures := Architectures + ',';
   Architectures := Architectures + 'x64';
 end;
 if Architectures = '' then begin Result := CustomMessage('SelectArchitecture'); Exit; end;
 ExtractTemporaryFile('ComTransaction.ps1');
 HelperPath := ExpandConstant('{tmp}\ComTransaction.ps1');
 Status := InvokeHelper('Inspect', HelperPath);
 if Status = 10 then begin
   if WizardSilent then Consent := CompareText(ExpandConstant('{param:MIGRATELEGACY|no}'), 'yes') = 0
   else Consent := MsgBox(CustomMessage('MigratePrompt') + ExistingDetails(), mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES;
   if not Consent then begin Result := CustomMessage('ConsentNeeded'); Exit; end;
 end else if Status <> 0 then begin Result := CustomMessage('Conflict32'); Exit; end;
 Transaction := ExpandConstant('{commonappdata}\ExcelKit\Transactions\') + GetDateTimeString('yyyymmdd-hhnnss-zzz', '-', ':') + '-' + IntToStr(Random(1000000));
 if InvokeHelper('Begin', HelperPath) <> 0 then begin Result := CustomMessage('OperationFailed') + #13#10 + Transaction; Exit; end;
 Prepared := True;
end;

function SaveInstallSelection(): Boolean;
var Lines: TArrayOfString; ManifestFile: String;
begin
 Result := False;
 ManifestFile := ExpandConstant('{app}\Installer\InstallManifest.json');
 if not FileExists(ManifestFile) then Exit;
 try
   if CompareText(GetSHA256OfFile(ManifestFile), '{#PackageManifestHash}') <> 0 then Exit;
 except Exit; end;
 SetArrayLength(Lines, 8);
 Lines[0] := 'installer_version={#ProductVersion}';
 Lines[1] := 'manifest_sha256={#PackageManifestHash}';
 Lines[2] := 'transaction=' + Transaction;
 if WizardIsTaskSelected('com32') then Lines[3] := 'com32=1' else Lines[3] := 'com32=0';
 if WizardIsTaskSelected('com64') then Lines[4] := 'com64=1' else Lines[4] := 'com64=0';
 if WizardIsTaskSelected('manual') then Lines[5] := 'manual=1' else Lines[5] := 'manual=0';
 if WizardIsTaskSelected('examples') then Lines[6] := 'examples=1' else Lines[6] := 'examples=0';
 if WizardIsTaskSelected('net10') then Lines[7] := 'net10=1' else Lines[7] := 'net10=0';
 Result := SaveStringsToUTF8File(ExpandConstant('{app}\Installer\install-selection.txt'), Lines, False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
 if CurStep = ssPostInstall then begin
   if InvokeHelper('Apply', ExpandConstant('{tmp}\ComTransaction.ps1')) <> 0 then ApplyFailed := True
   else if not SaveInstallSelection() then ApplyFailed := True;
   if ApplyFailed then
     SuppressibleMsgBox(CustomMessage('OperationFailed') + #13#10 + Transaction, mbError, MB_OK, IDOK);
 end;
 if (CurStep = ssDone) and not ApplyFailed and not Finished then begin
   if InvokeHelper('Finish', ExpandConstant('{tmp}\ComTransaction.ps1')) <> 0 then ApplyFailed := True
   else Finished := True;
 end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var I: Integer;
begin
 Result := True;
 if (CurPageID = wpFinished) and Prepared and not Finished and not ApplyFailed then begin
   if InvokeHelper('Finish', ExpandConstant('{tmp}\ComTransaction.ps1')) <> 0 then begin
     ApplyFailed := True;
     SuppressibleMsgBox(CustomMessage('OperationFailed') + #13#10 + Transaction, mbError, MB_OK, IDOK);
   end else Finished := True;
 end;
 if (CurPageID = wpFinished) and ApplyFailed then
   for I := 0 to WizardForm.RunList.Items.Count - 1 do WizardForm.RunList.Checked[I] := False;
end;

procedure DeinitializeSetup;
begin
 if Prepared and not Finished then begin
   if InvokeHelper('Rollback', ExpandConstant('{tmp}\ComTransaction.ps1')) <> 0 then
     MsgBox(CustomMessage('OperationFailed') + #13#10 + Transaction, mbError, MB_OK);
 end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var CheckResult: Integer;
begin
 if CurUninstallStep <> usUninstall then Exit;
 Architectures := 'x86'; if IsWin64 then Architectures := 'x86,x64';
 CheckResult := InvokeHelper('UninstallCheck', ExpandConstant('{app}\Installer\ComTransaction.ps1'));
 if CheckResult = 0 then Exit;
 if CheckResult <> 10 then RaiseException(CustomMessage('OperationFailed'));
 Transaction := ExpandConstant('{commonappdata}\ExcelKit\Transactions\') + GetDateTimeString('yyyymmdd-hhnnss-zzz', '-', ':') + '-uninstall-' + IntToStr(Random(1000000));
 if InvokeHelper('Begin', ExpandConstant('{app}\Installer\ComTransaction.ps1')) <> 0 then
   RaiseException(CustomMessage('OperationFailed') + #13#10 + Transaction);
 if InvokeHelper('Uninstall', ExpandConstant('{app}\Installer\ComTransaction.ps1')) <> 0 then
   RaiseException(CustomMessage('OperationFailed') + #13#10 + Transaction);
end;

function GetCustomSetupExitCode(): Integer;
begin
 if ApplyFailed then Result := 1 else Result := 0;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
 if (CurPageID = wpFinished) and ApplyFailed then
   WizardForm.FinishedLabel.Caption := CustomMessage('OperationFailed') + #13#10 + Transaction;
end;
