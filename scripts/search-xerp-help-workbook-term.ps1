param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$Pattern = "Zusatzfeld|Zusatzfelder|Zusatz-Feld|Zusatz-Felder"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $hits = New-Object System.Collections.Generic.List[object]
    foreach ($ws in $pkg.Workbook.Worksheets) {
        if ($null -eq $ws.Dimension) { continue }
        for ($r = 1; $r -le $ws.Dimension.End.Row; $r++) {
            for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
                $text = [string]$ws.Cells.Item($r, $c).Text
                if ($text -match $Pattern) {
                    $hits.Add([pscustomobject]@{
                        Sheet = $ws.Name
                        Row = $r
                        Column = $c
                        Text = $text
                    })
                }
            }
        }
    }

    [pscustomobject]@{
        Workbook = $WorkbookPath
        Pattern = $Pattern
        Hits = $hits.Count
        First = @($hits | Select-Object -First 20)
    } | ConvertTo-Json -Depth 4
}
finally {
    $pkg.Dispose()
}
