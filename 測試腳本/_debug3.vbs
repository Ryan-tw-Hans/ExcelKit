Dim fso: Set fso=CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName),"_TestCommon.vbs"),1).ReadAll
WScript.Echo "ok"


