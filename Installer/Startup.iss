var StartupFolder: String;

function PathCreateFromUrlW(Url, Path: String; var PathLength: Cardinal; Flags: Cardinal): LongInt;
  external 'PathCreateFromUrlW@shlwapi.dll stdcall';

function LocalPathFromCodeBase(const Url: String): String;
var Buffer: String; N: Cardinal;
begin
 Result := '';
 if CompareText(Copy(Url, 1, 8), 'file:///') <> 0 then Exit;
 N := 32767;
 Buffer := StringOfChar(#0, N);
 if PathCreateFromUrlW(Url, Buffer, N, 0) <> 0 then Exit;
 SetLength(Buffer, N);
 if (Length(Buffer) < 3) or (Buffer[2] <> ':') or (Buffer[3] <> '\') then Exit;
 Result := Buffer;
end;

function AssemblyVersion(const Assembly: String): String;
var P, Q: Integer; S: String;
begin
 Result := '';
 P := Pos('Version=', Assembly);
 if P = 0 then Exit;
 S := Copy(Assembly, P + 8, Length(Assembly));
 Q := Pos(',', S);
 if Q > 0 then S := Copy(S, 1, Q - 1);
 Result := Trim(S);
end;

function InspectView(const MachineRoot, UserRoot: Integer; const ViewName: String): String;
var Key, Assembly, CodeBase, Dll, FileVer, RegVer, Hash: String;
    Names: TArrayOfString;
begin
 Key := 'Software\Classes\CLSID\{87654321-4321-4321-4321-210987654321}\InprocServer32';
 Result := ViewName + ': ';
 if RegKeyExists(UserRoot, 'Software\Classes\CLSID\{87654321-4321-4321-4321-210987654321}') or
    RegKeyExists(UserRoot, 'Software\Classes\ExcelKit.ExcelOperator') then begin
   Result := Result + CustomMessage('StartupUserOverride') + #13#10; Exit;
 end;
 if not RegKeyExists(MachineRoot, Key) then begin
   Result := Result + CustomMessage('StartupNotFound') + #13#10; Exit;
 end;
 if not RegGetSubkeyNames(MachineRoot, Key, Names) then begin
   Result := Result + CustomMessage('StartupUnknown') + #13#10; Exit;
 end;
 if GetArrayLength(Names) <> 1 then begin
   Result := Result + CustomMessage('StartupMultiple') + #13#10; Exit;
 end;
 if not RegQueryStringValue(MachineRoot, Key, 'Assembly', Assembly) or
    not RegQueryStringValue(MachineRoot, Key, 'CodeBase', CodeBase) then begin
   Result := Result + CustomMessage('StartupUnknown') + #13#10; Exit;
 end;
 RegVer := AssemblyVersion(Assembly);
 if (CompareText(Copy(Assembly, 1, 18), 'ExcelKit, Version=') <> 0) or
    (RegVer = '') or (CompareText(Names[0], RegVer) <> 0) then begin
   Result := Result + CustomMessage('StartupUnknown') + #13#10; Exit;
 end;
 Result := Result + #13#10 + '  ' + CodeBase + #13#10;
 Dll := LocalPathFromCodeBase(CodeBase);
 if (Dll = '') or not FileExists(Dll) then begin
   Result := Result + '  ' + CustomMessage('StartupMissingOrUnreadable') + #13#10; Exit;
 end;
 if GetVersionNumbersString(Dll, FileVer) then
   Result := Result + '  ' + CustomMessage('StartupFileVersion') + ' ' + FileVer + #13#10
 else Result := Result + '  ' + CustomMessage('StartupFileVersion') + ' ' + CustomMessage('StartupUnknown') + #13#10;
 try
   Hash := GetSHA256OfFile(Dll);
   if CompareText(Hash, '{#PackageAssemblyHash}') = 0 then
     Result := Result + '  ' + CustomMessage('StartupHashSame') + #13#10
   else Result := Result + '  ' + CustomMessage('StartupHashDifferent') + #13#10;
 except
   Result := Result + '  ' + CustomMessage('StartupHashUnknown') + #13#10;
 end;
end;

function ResolveInstalledFolder(var InstalledVersion: String): String;
var Key, A, B, VA, VB: String; KA, KB: Boolean;
begin
 Result := ''; InstalledVersion := CustomMessage('StartupUnknown');
 Key := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{62953C31-A16D-469E-8866-6218DF2B3D2A}_is1';
 KA := RegKeyExists(HKLM32, Key); KB := False;
 if IsWin64 then KB := RegKeyExists(HKLM64, Key);
 if not KA and not KB then Exit;
 if KA and not RegQueryStringValue(HKLM32, Key, 'InstallLocation', A) then Exit;
 if KB and not RegQueryStringValue(HKLM64, Key, 'InstallLocation', B) then Exit;
 if KA then RegQueryStringValue(HKLM32, Key, 'DisplayVersion', VA);
 if KB then RegQueryStringValue(HKLM64, Key, 'DisplayVersion', VB);
 if KA and KB then begin
   if CompareText(RemoveBackslashUnlessRoot(A), RemoveBackslashUnlessRoot(B)) <> 0 then Exit;
   if CompareText(VA, VB) <> 0 then Exit;
 end;
 if KA then begin Result := A; InstalledVersion := VA; end
 else begin Result := B; InstalledVersion := VB; end;
 if (Length(Result) < 3) or (Result[2] <> ':') or (Result[3] <> '\') or not DirExists(Result) then begin
   Result := ''; InstalledVersion := CustomMessage('StartupUnknown');
 end;
 if InstalledVersion = '' then InstalledVersion := CustomMessage('StartupUnknown');
end;

procedure StartupLinkClick(Sender: TObject; const Link: String; LinkType: TSysLinkType);
var Target: String; ErrorCode: Integer;
begin
 Target := '';
 if LinkType <> sltID then Exit;
 if Link = 'install' then Target := 'https://www.libreoffice.org/download/'
 else if Link = 'portable' then Target := 'https://portableapps.com/apps/office/libreoffice_portable'
 else if Link = 'folder' then Target := ResolveInstalledFolder(StartupFolder);
 if Target = '' then begin MsgBox(CustomMessage('StartupFolderUnknown'), mbInformation, MB_OK); Exit; end;
 if not ShellExecAsOriginalUser('open', Target, '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode) then
   MsgBox(CustomMessage('StartupOpenFailed'), mbError, MB_OK);
end;

procedure InitializeWizard;
var Links: TNewLinkLabel; Status: TNewMemo; V, S, Folder: String; LeftPos, TopPos, Width: Integer;
begin
 WizardForm.WelcomeLabel2.Visible := False;
 LeftPos := WizardForm.WelcomeLabel2.Left;
 Width := WizardForm.WelcomeLabel2.Width;
 TopPos := WizardForm.WelcomeLabel1.Top + WizardForm.WelcomeLabel1.Height + ScaleY(4);
 Links := TNewLinkLabel.Create(WizardForm);
 Links.Parent := WizardForm.WelcomePage;
 Links.Left := LeftPos; Links.Top := TopPos; Links.Width := Width;
 Links.AutoSize := False;
 Folder := ResolveInstalledFolder(V);
 S := '<a id="install">' + CustomMessage('LibreOfficeInstall') + '</a>' + #13#10 +
      '<a id="portable">' + CustomMessage('LibreOfficePortable') + '</a>' + #13#10;
 if Folder <> '' then S := S + '<a id="folder">' + CustomMessage('OpenInstallFolder') + '</a>'
 else S := S + CustomMessage('StartupFolderUnknown');
 Links.Caption := S; Links.OnLinkClick := @StartupLinkClick; Links.UseVisualStyle := True; Links.AdjustHeight;
 Status := TNewMemo.Create(WizardForm);
 Status.Parent := WizardForm.WelcomePage;
 Status.Left := LeftPos; Status.Top := Links.Top + Links.Height + ScaleY(8); Status.Width := Width;
 Status.Height := WizardForm.WelcomePage.ClientHeight - Status.Top - ScaleY(8);
 Status.ReadOnly := True; Status.ScrollBars := ssVertical; Status.WordWrap := True; Status.TabStop := True;
 S := CustomMessage('StartupInstallerVersion') + ' {#ProductVersion}' + #13#10 +
      CustomMessage('StartupInstalledSetupVersion') + ' ' + V + #13#10 +
      CustomMessage('StartupPackageComponent') + ' {#PackageComponentVersion}' + #13#10 +
      CustomMessage('StartupPackageFile') + ' {#PackageFileVersion}' + #13#10#13#10 +
      InspectView(HKLM32, HKCU32, 'x86');
 if IsWin64 then S := S + #13#10 + InspectView(HKLM64, HKCU64, 'x64');
 Status.Text := S;
end;
