param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.sorted.xlsx",
    [string]$WorksheetName = "de-DE",
    [int]$Row = 263,
    [int]$Column = 1
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $text = [string]$pkg.Workbook.Worksheets[$WorksheetName].Cells.Item($Row, $Column).Text
    [pscustomobject]@{
        Text = $text
        Codes = (($text.ToCharArray() | ForEach-Object { [int][char]$_ }) -join ',')
    } | ConvertTo-Json
}
finally {
    $pkg.Dispose()
}
