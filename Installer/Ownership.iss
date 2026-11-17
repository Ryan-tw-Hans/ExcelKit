// Pure decision used by the installer and isolated test harness.
function AllowRegistrationState(UserConflict, MachineConflict, ConfirmedOwner: Boolean): Boolean;
begin
  Result := (not UserConflict) and ((not MachineConflict and not ConfirmedOwner) or (MachineConflict and ConfirmedOwner));
end;
