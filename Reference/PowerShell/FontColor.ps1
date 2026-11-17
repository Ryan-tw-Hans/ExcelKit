# Windows PowerShell 5.1; use matching bitness for the registered COM component.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP 'ExcelKitExamples'),
    [string]$DllPath = '' # Optional: test an unregistered net472 build without changing COM registration.
)
$ErrorActionPreference = 'Stop'
[IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null
$excel = $null
try {
    if ($DllPath) {
        Add-Type -Path (Resolve-Path -LiteralPath $DllPath).Path
        $excel = New-Object Excelopen.ExcelOperator
    } else { $excel = New-Object -ComObject ExcelKit.ExcelOperator }
    $excel.SetLogEnabled($false)
    $excel.CreateNewWorkbook()
    $excel.SetCellValue('Sheet1', 'A1', 'Red text')
    $excel.SetCellFontColor('Sheet1', 'A1', 255, 0, 0)
    $excel.SetCellValueByRC('Sheet1', 2, 2, 'Blue text')
    $excel.SetCellFontColorByRC('Sheet1', 2, 2, 0, 0, 255)
    $filePath = Join-Path $OutputDirectory 'ExcelKit-FontColor.xlsx'
    $excel.SaveWorkbook($filePath)
    Write-Output $excel.GetVersionInfo()
    Write-Output "Workbook saved: $filePath"
} finally {
    if ($excel -and [Runtime.InteropServices.Marshal]::IsComObject($excel)) {
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    } elseif ($excel -is [IDisposable]) { $excel.Dispose() }
}
