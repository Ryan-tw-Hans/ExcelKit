Option Explicit

Dim fso, shell, scriptFolder, folder, file, dllPath, cmd, exitCode
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Set folder = fso.GetFolder(scriptFolder)

dllPath = ""
If WScript.Arguments.Count > 0 Then
    dllPath = Trim(WScript.Arguments(0))
End If

For Each file In folder.Files
    If LCase(Right(file.Name, 4)) = ".vbs" Then
        If file.Name <> "_TestCommon.vbs" And file.Name <> "Run_All_Ordered_Tests.vbs" Then
            If Mid(file.Name, 3, 1) = "_" Then
                cmd = "cscript //nologo """ & file.Path & """"
                If Len(dllPath) > 0 Then
                    cmd = cmd & " """ & dllPath & """"
                End If

                WScript.Echo "Running: " & file.Name
                exitCode = shell.Run(cmd, 1, True)
                If exitCode <> 0 Then
                    WScript.Echo "Stopped due to failure in: " & file.Name
                    WScript.Quit exitCode
                End If
            End If
        End If
    End If
Next

WScript.Echo "All ordered test scripts completed."
Dim librePath
librePath = GetLibreOfficePath()

orderedScripts = Array( _
    "01_Reference_COM_Component.vbs", _
    "02_Read_DLL_Version.vbs", _
    "03_Create_WorkBook.vbs", _
    "08_Save_WorkBook.vbs", _
    "04_Create_WorkSheet.vbs", _
    "05_Set_Cell_By_Name.vbs", _
    "06_Set_Cell_By_Row_Column.vbs", _
    "07_Read_Cell_Value.vbs", _
    "09_Export_To_PDF.vbs", _
    "10_Print_WorkBook.vbs", _
    "11_Open_WorkBook.vbs", _
    "12_Delete_Sheet.vbs", _
    "13_Clear_Cell.vbs", _
    "14_Insert_Row.vbs", _
    "15_Delete_Row.vbs", _
    "16_Set_Get_Row_Height.vbs", _
    "17_Insert_Delete_Column.vbs", _
    "18_Set_Get_Column_Width.vbs", _
    "19_Get_All_Sheet_Names.vbs", _
    "20_Config_Functions.vbs" _
)

For i = 0 To UBound(orderedScripts)
    scriptPath = fso.BuildPath(scriptFolder, orderedScripts(i))

    If fso.FileExists(scriptPath) Then
        cmd = "cscript //nologo """ & scriptPath & """"
        If Len(dllPath) > 0 Then
            cmd = cmd & " """ & dllPath & """"
        End If
        If Len(librePath) > 0 Then
            cmd = cmd & " """ & librePath & """"
        End If

        WScript.Echo "Running: " & orderedScripts(i)
        exitCode = shell.Run(cmd, 1, True)
        If exitCode <> 0 Then
            WScript.Echo "Stopped due to failure in: " & orderedScripts(i)
            WScript.Quit exitCode
        End If
    End If
Next

