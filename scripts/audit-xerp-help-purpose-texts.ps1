param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\purpose-text-issues.csv"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $issues = New-Object System.Collections.Generic.List[object]

    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $contentType = [string]$ws.Cells.Item($r, 16).Text
        if ($contentType -ne 'View') { continue }
        $topic = [string]$ws.Cells.Item($r, 1).Text
        $text = [string]$ws.Cells.Item($r, 8).Text
        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        $issue = $null
        if ($text -match 'Bearbeitungsansicht unterst|Die Hilfe beschreibt') { $issue = 'OLD_GENERIC_TEXT' }
        elseif ($text -match 'Sie beschreibt, welche Informationen') { $issue = 'WEAK_FALLBACK_TEXT' }
        elseif ($text -match 'pflegt [A-Z][A-Za-z ]+ im Bereich') { $issue = 'TECHNICAL_LABEL_TEXT' }

        if ($issue) {
            $issues.Add([pscustomobject]@{
                Row = $r
                Issue = $issue
                Topic = $topic
                Text = $text
            })
        }
    }

    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputCsv) | Out-Null
    $issues | Export-Csv -LiteralPath $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    [pscustomobject]@{
        Issues = $issues.Count
        OutputCsv = $OutputCsv
        First = @($issues | Select-Object -First 30)
    } | ConvertTo-Json -Depth 4
}
finally {
    $pkg.Dispose()
}
