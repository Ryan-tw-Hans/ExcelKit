# Run using Windows PowerShell 5.1 with the same bitness as the registered COM component.
param([string]$OutputDirectory = (Join-Path $env:TEMP 'ExcelKitExamples'))
$ErrorActionPreference = 'Stop'
[IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null
$excel = $null
try {
    $excel = New-Object -ComObject ExcelKit.ExcelOperator
    $excel.SetLogEnabled($false)
    $excel.CreateNewWorkbook()
    $excel.SetCellValue('Sheet1', 'A1', 'Hello ExcelKit!')
    $excel.SetCellValue('Sheet1', 'A2', 123.45)
    if ($excel.GetCellValue('Sheet1', 'A1') -ne 'Hello ExcelKit!') { throw 'Cell read-back failed.' }
    $filePath = Join-Path $OutputDirectory 'ExcelKit-QuickStart.xlsx'
    $excel.SaveWorkbook($filePath)
    Write-Output $excel.GetVersionInfo()
    Write-Output "Workbook saved: $filePath"
} finally {
    if ($excel -and [Runtime.InteropServices.Marshal]::IsComObject($excel)) {
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
}
