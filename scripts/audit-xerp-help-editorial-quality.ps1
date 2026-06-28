param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\editorial-quality-issues.csv"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$patterns = @(
    @{ Name = 'GENERIC_FIELD'; Regex = 'Beschreibt das Feld' },
    @{ Name = 'GENERIC_REGISTER'; Regex = 'werden die zugehoerigen Informationen|werden die zugehörigen Informationen' },
    @{ Name = 'ASCII_UMLAUT'; Regex = '\bfuer\b|\bFuer\b|Datensaetze|Pruefen|Schaltflaechen|zugehoerig|gehoeren|ergaenz' },
    @{ Name = 'WEAK_INPUT'; Regex = '^Eingabefeld für:|^Eingabefeld fuer:' },
    @{ Name = 'TOO_SHORT'; Regex = '^.{0,30}$' }
)

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) { throw "Worksheet not found or empty: $WorksheetName" }
    $rows = New-Object System.Collections.Generic.List[object]

    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $text = [string]$ws.Cells.Item($r, 8).Text
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        foreach ($p in $patterns) {
            if ($text -match $p.Regex) {
                $rows.Add([pscustomobject]@{
                    Row = $r
                    Issue = $p.Name
                    Level = [int]$ws.Row($r).OutlineLevel
                    Topic = [string]$ws.Cells.Item($r, 1).Text
                    Field = [string]$ws.Cells.Item($r, 6).Text
                    Text = $text
                })
                break
            }
        }
    }

    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputCsv) | Out-Null
    $rows | Export-Csv -LiteralPath $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    $rows | Group-Object Issue | Select-Object Name,Count | ConvertTo-Json
}
finally {
    $pkg.Dispose()
}
