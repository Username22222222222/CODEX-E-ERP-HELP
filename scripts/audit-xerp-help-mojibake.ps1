param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $hits = New-Object System.Collections.Generic.List[object]
    $bad1 = [char]195
    $bad2 = [char]194
    for ($r = 1; $r -le $ws.Dimension.End.Row; $r++) {
        for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
            $text = [string]$ws.Cells.Item($r, $c).Text
            if ($text.IndexOf($bad1) -ge 0 -or $text.IndexOf($bad2) -ge 0) {
                $hits.Add([pscustomobject]@{
                    Row = $r
                    Column = $c
                    Text = $text
                })
            }
        }
    }
    [pscustomobject]@{
        MojibakeHits = $hits.Count
        FirstHits = @($hits | Select-Object -First 10)
    } | ConvertTo-Json -Depth 4
}
finally {
    $pkg.Dispose()
}
