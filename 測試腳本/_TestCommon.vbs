Option Explicit

Function GetFso()
    Set GetFso = CreateObject("Scripting.FileSystemObject")
End Function

Function GetComVersionToken()
    Dim op, versionText
    versionText = ""

    On Error Resume Next
    Set op = CreateObject("ExcelKit.ExcelOperator")
    If Err.Number = 0 Then
        versionText = Trim(CStr(op.GetVersionInfo()))
    End If
    If Err.Number <> 0 Then
        versionText = ""
        Err.Clear
    End If
    On Error GoTo 0

    If Len(versionText) = 0 Then
        GetComVersionToken = ""
    Else
        GetComVersionToken = SanitizeFileToken(versionText)
    End If
End Function

Function SanitizeFileToken(value)
    Dim token
    token = Trim(CStr(value))
    token = Replace(token, "\\", "_")
    token = Replace(token, "/", "_")
    token = Replace(token, ":", "_")
    token = Replace(token, "*", "_")
    token = Replace(token, "?", "_")
    token = Replace(token, """", "_")
    token = Replace(token, "<", "_")
    token = Replace(token, ">", "_")
    token = Replace(token, "|", "_")
    token = Replace(token, "(", "_")
    token = Replace(token, ")", "_")
    token = Replace(token, " ", "_")
    token = Replace(token, ".", "_")

    Do While InStr(token, "__") > 0
        token = Replace(token, "__", "_")
    Loop

    If Len(token) = 0 Then
        token = "UnknownVersion"
    End If

    SanitizeFileToken = token
End Function

Function W(hexCode)
    W = ChrW(CLng("&H" & hexCode))
End Function

Function GetShell()
    Set GetShell = CreateObject("WScript.Shell")
End Function

Function GetScriptFolder()
    Dim fso
    Set fso = GetFso()
    GetScriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
End Function

Function BuildPath(basePath, childPath)
    Dim fso
    Set fso = GetFso()
    BuildPath = fso.BuildPath(basePath, childPath)
End Function

Sub EnsureFolder(path)
    Dim fso
    Set fso = GetFso()
    If Not fso.FolderExists(path) Then
        fso.CreateFolder(path)
    End If
End Sub

Sub WriteUtf8TextFile(path, content, append)
    Dim stream, fso
    Set fso = GetFso()
    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open

    If append And fso.FileExists(path) Then
        stream.LoadFromFile path
        stream.Position = stream.Size
    End If

    stream.WriteText content, 0
    stream.SaveToFile path, 2
    stream.Close
End Sub

Function GetArgOrDefault(index, defaultValue)
    If WScript.Arguments.Count > index Then
        GetArgOrDefault = Trim(WScript.Arguments(index))
    Else
        GetArgOrDefault = defaultValue
    End If
End Function

Function GetDynamicValue(prefix)
    Randomize
    GetDynamicValue = prefix & "_" & Replace(Replace(Replace(CStr(Now), "/", ""), ":", ""), " ", "_") & "_" & CStr(Int((Rnd() * 900000) + 100000))
End Function

Function ColumnNumberToLetter(columnNumber)
    Dim dividend, modulo, columnName
    dividend = CLng(columnNumber)
    columnName = ""

    Do While dividend > 0
        modulo = (dividend - 1) Mod 26
        columnName = Chr(65 + modulo) & columnName
        dividend = Int((dividend - modulo) / 26)
    Loop

    ColumnNumberToLetter = columnName
End Function

Function GetCellAddressByRC(row, column)
    GetCellAddressByRC = ColumnNumberToLetter(column) & CStr(row)
End Function

Function ReadCellValueByAdo(workbookPath, sheetName, cellAddress)
    Dim conn, rs, connStr, sql, value
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & workbookPath & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=No;IMEX=1"";"

    conn.Open connStr
    sql = "SELECT * FROM [" & sheetName & "$" & cellAddress & ":" & cellAddress & "]"
    rs.Open sql, conn, 0, 1

    If rs.EOF Then
        value = ""
    ElseIf IsNull(rs.Fields(0).Value) Then
        value = ""
    Else
        value = CStr(rs.Fields(0).Value)
    End If

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    ReadCellValueByAdo = value
End Function

Function FindLocalLibreOfficePath()
    Dim fso, scriptFolder, candidates, i
    Set fso = GetFso()
    scriptFolder = GetScriptFolder()

    candidates = Array( _
        BuildPath(scriptFolder, "LibreOfficePortable\\App\\libreoffice\\program\\soffice.exe"), _
        BuildPath(scriptFolder, "LibreOfficePortable\\program\\soffice.exe"), _
        BuildPath(scriptFolder, "LibreOffice\\program\\soffice.exe"), _
        BuildPath(scriptFolder, "libreoffice\\program\\soffice.exe"), _
        BuildPath(scriptFolder, "LibreOffice\\soffice.exe"), _
        BuildPath(scriptFolder, "libreoffice\\soffice.exe"), _
        BuildPath(scriptFolder, "soffice.exe") _
    )

    For i = 0 To UBound(candidates)
        If fso.FileExists(candidates(i)) Then
            FindLocalLibreOfficePath = candidates(i)
            Exit Function
        End If
    Next

    FindLocalLibreOfficePath = ""
End Function

Function GetLibreOfficePath()
    GetLibreOfficePath = GetArgOrDefault(1, FindLocalLibreOfficePath())
End Function

Function GetDllPath()
    Dim fso, defaultPath
    Set fso = GetFso()
    defaultPath = BuildPath(GetScriptFolder(), "..\\Excelopen\\bin\\Debug\\net472\\Excelopen.dll")
    defaultPath = fso.GetAbsolutePathName(defaultPath)
    GetDllPath = GetArgOrDefault(0, defaultPath)
End Function

Function GetDllVersion()
    Dim fso, dllPath, version
    Set fso = GetFso()
    dllPath = GetDllPath()

    On Error Resume Next
    If fso.FileExists(dllPath) Then
        version = fso.GetFileVersion(dllPath)
    End If
    If Err.Number <> 0 Then
        version = "UnknownVersion"
        Err.Clear
    End If
    On Error GoTo 0

    If Len(Trim(version)) = 0 Then
        version = "UnknownVersion"
    End If

    version = Replace(version, ".", "_")
    version = Replace(version, " ", "_")
    GetDllVersion = version
End Function

Function GetDataFolder()
    Dim path
    path = BuildPath(GetScriptFolder(), "TestData")
    EnsureFolder path
    GetDataFolder = path
End Function

Function GetReportFolder()
    Dim path
    path = BuildPath(GetScriptFolder(), "TestReports")
    EnsureFolder path
    GetReportFolder = path
End Function

Function GetWorkbookPath()
    GetWorkbookPath = BuildPath(GetDataFolder(), "ExcelKit_AutoTest.xlsx")
End Function

Function GetPdfPath()
    GetPdfPath = BuildPath(GetDataFolder(), "ExcelKit_AutoTest.pdf")
End Function

Function GetReportPath()
    Dim fileName, reportVersion
    reportVersion = GetComVersionToken()
    If Len(reportVersion) = 0 Then
        reportVersion = GetDllVersion()
    End If

    fileName = "TestReport_COM_" & reportVersion & ".md"
    GetReportPath = BuildPath(GetReportFolder(), fileName)
End Function

Sub EnsureReportHeader()
    Dim fso, reportPath, header
    Set fso = GetFso()
    reportPath = GetReportPath()

    If Not fso.FileExists(reportPath) Then
        header = "# ExcelKit COM " & W("6E2C") & W("8A66") & W("5831") & W("544A") & vbCrLf & _
                 vbCrLf & _
                 "- DLL Path: `" & GetDllPath() & "`" & vbCrLf & _
                 "- DLL Version: `" & GetDllVersion() & "`" & vbCrLf & _
                 "- " & W("5EFA") & W("7ACB") & W("6642") & W("9593") & ": `" & Now & "`" & vbCrLf & _
                 vbCrLf & _
                 "| " & W("6E2C") & W("8A66") & W("7DE8") & W("865F") & " | " & W("529F") & W("80FD") & " | " & W("72C0") & W("614B") & " | " & W("6642") & W("9593") & " | " & W("8A0A") & W("606F") & " |" & vbCrLf & _
                 "|---|---|---|---|---|" & vbCrLf

        WriteUtf8TextFile reportPath, header, False
    End If
End Sub

Function GetTestDisplayName(testNo)
    Select Case CStr(testNo)
        Case "01": GetTestDisplayName = W("5F15") & W("7528") & " COM " & W("5143") & W("4EF6")
        Case "02": GetTestDisplayName = W("8B80") & W("53D6") & " DLL " & W("7248") & W("672C")
        Case "03": GetTestDisplayName = W("5EFA") & W("7ACB") & " Excel " & W("6A94") & W("6848") & " (WorkBook)"
        Case "04": GetTestDisplayName = W("5EFA") & W("7ACB") & " Sheet (WorkSheet)"
        Case "05": GetTestDisplayName = W("8A2D") & W("5B9A") & W("5132") & W("5B58") & W("683C") & W("8CC7") & W("6599") & " by cell name"
        Case "06": GetTestDisplayName = W("8A2D") & W("5B9A") & W("5132") & W("5B58") & W("683C") & W("8CC7") & W("6599") & " by row column"
        Case "07": GetTestDisplayName = W("8B80") & W("53D6") & W("5132") & W("5B58") & W("683C") & W("8CC7") & W("6599")
        Case "08": GetTestDisplayName = W("5132") & W("5B58") & W("6A94") & W("6848")
        Case "09": GetTestDisplayName = W("532F") & W("51FA") & W("6210") & " PDF"
        Case "10": GetTestDisplayName = W("5217") & W("5370")
        Case "11": GetTestDisplayName = W("958B") & W("555F") & W("6A94") & W("6848")
        Case "12": GetTestDisplayName = W("522A") & W("9664") & " Sheet"
        Case "13": GetTestDisplayName = W("6E05") & W("9664") & W("5132") & W("5B58") & W("683C")
        Case "14": GetTestDisplayName = W("63D2") & W("5165") & W("5217")
        Case "15": GetTestDisplayName = W("522A") & W("9664") & W("5217")
        Case "16": GetTestDisplayName = W("8A2D") & W("5B9A") & "/" & W("8B80") & W("53D6") & W("884C") & W("9AD8")
        Case "17": GetTestDisplayName = W("63D2") & W("5165") & "/" & W("522A") & W("9664") & W("6B04")
        Case "18": GetTestDisplayName = W("8A2D") & W("5B9A") & "/" & W("8B80") & W("53D6") & W("6B04") & W("5BEC")
        Case "19": GetTestDisplayName = W("53D6") & W("5F97") & W("6240") & W("6709") & " Sheet " & W("540D") & W("7A31")
        Case "20": GetTestDisplayName = W("5176") & W("4ED6") & W("51FD") & W("5F0F") & W("8207") & W("8A2D") & W("5B9A") & W("529F") & W("80FD")
        Case Else: GetTestDisplayName = W("672A") & W("77E5")
    End Select
End Function

Sub LogTestResult(testNo, testName, status, message)
    Dim fso, reportPath, line, displayName, safeMessage, statusText
    Set fso = GetFso()

    EnsureReportHeader
    reportPath = GetReportPath()
    displayName = GetTestDisplayName(testNo)

    Select Case UCase(CStr(status))
        Case "PASS"
            safeMessage = W("901A") & W("904E")
            statusText = W("901A") & W("904E")
        Case "SKIP"
            safeMessage = W("7565") & W("904E")
            statusText = W("7565") & W("904E")
        Case Else
            safeMessage = W("5931") & W("6557") & W("FF1A") & CStr(message)
            statusText = W("5931") & W("6557")
    End Select

    safeMessage = Replace(safeMessage, vbCrLf, " ")
    safeMessage = Replace(safeMessage, vbCr, " ")
    safeMessage = Replace(safeMessage, vbLf, " ")
    safeMessage = Replace(safeMessage, "|", "\\|")
    line = "| " & testNo & " | " & displayName & " | " & statusText & " | " & Now & " | " & safeMessage & " |" & vbCrLf

    WriteUtf8TextFile reportPath, line, True

    WScript.Echo "¡i" & statusText & "¡j" & testNo & " " & displayName & " - " & safeMessage
    WScript.Echo W("5831") & W("8868") & "¡G" & reportPath
End Sub

Function CreateOperator()
    Set CreateOperator = CreateObject("ExcelKit.ExcelOperator")
End Function

Sub ResetTestFiles()
    Dim fso, workbookPath, pdfPath
    Set fso = GetFso()

    workbookPath = GetWorkbookPath()
    pdfPath = GetPdfPath()

    If fso.FileExists(workbookPath) Then fso.DeleteFile workbookPath, True
    If fso.FileExists(pdfPath) Then fso.DeleteFile pdfPath, True
End Sub

Sub SaveMinimalWorkbook(operatorObj)
    Dim workbookPath
    workbookPath = GetWorkbookPath()

    operatorObj.CreateNewWorkbook
    operatorObj.SetCellValue "Sheet1", "A1", GetDynamicValue("Test")
    operatorObj.SaveWorkbook workbookPath
End Sub

