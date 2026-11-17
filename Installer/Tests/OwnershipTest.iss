[Setup]
AppName=ExcelKit registration decision test
AppVersion=1.0
CreateAppDir=no
Uninstallable=no
PrivilegesRequired=lowest
DisableWelcomePage=yes
DisableReadyPage=yes
DisableFinishedPage=yes
OutputDir=..\..\.tools\installer-tests
OutputBaseFilename=ownership-test
[Code]
#include "..\Ownership.iss"
function InitializeSetup: Boolean;
var UserConflict, MachineConflict, ConfirmedOwner, Expected, Actual: Boolean; I: Integer;
begin
  for I := 0 to 7 do begin
    UserConflict := (I and 1) <> 0;
    MachineConflict := (I and 2) <> 0;
    ConfirmedOwner := (I and 4) <> 0;
    Expected := (I = 0) or (I = 6);
    Actual := AllowRegistrationState(UserConflict, MachineConflict, ConfirmedOwner);
    if Actual <> Expected then begin
      Log('FAIL state=' + IntToStr(I));
      Result := False;
      Exit;
    end;
  end;
  Log('PASS: all 8 registration ownership states');
  Result := True;
end;
