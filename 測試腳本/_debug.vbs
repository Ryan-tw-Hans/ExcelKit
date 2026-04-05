Option Explicit
Dim fso,p
Set fso=CreateObject("Scripting.FileSystemObject")
p=fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName),"_TestCommon.vbs")
WScript.Echo p
WScript.Echo fso.FileExists(p)
Dim t:Set t=fso.OpenTextFile(p,1)
WScript.Echo "opened"
t.Close


