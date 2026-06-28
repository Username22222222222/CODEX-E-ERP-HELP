param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [int]$StartRow = 1188,
    [int]$EndRow = 1200,
    [int]$MaxCol = 20
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws) { throw "Worksheet not found: $WorksheetName" }

    for ($r = $StartRow; $r -le $EndRow; $r++) {
        $parts = New-Object System.Collections.Generic.List[string]
        for ($c = 1; $c -le $MaxCol; $c++) {
            $text = [string]$ws.Cells.Item($r, $c).Text
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $parts.Add("C$c=$text")
            }
        }
        [pscustomobject]@{
            Row = $r
            OutlineLevel = [int]$ws.Row($r).OutlineLevel
            Hidden = [bool]$ws.Row($r).Hidden
            Values = ($parts -join ' | ')
        }
    }
}
finally {
    $pkg.Dispose()
}
