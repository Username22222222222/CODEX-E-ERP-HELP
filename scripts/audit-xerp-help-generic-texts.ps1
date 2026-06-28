param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\generic-text-issues.csv"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$patterns = @(
    @{ Issue = 'OLD_GENERIC_VIEW_TEXT'; Pattern = 'Bearbeitungsansicht unterst' },
    @{ Issue = 'OLD_HELP_EXPLAINS_TEXT'; Pattern = 'Die Hilfe beschreibt' },
    @{ Issue = 'WEAK_PROCESS_INFORMATION_TEXT'; Pattern = 'Sie beschreibt, welche Informationen' },
    @{ Issue = 'WEAK_USER_SUPPORT_TEXT'; Pattern = 'unterst' + [string][char]0x00FC + 'tzt Anwender beim Anlegen' },
    @{ Issue = 'WEAK_GENERIC_PERSPECTIVE_TEXT'; Pattern = 'fachlicher Sicht' },
    @{ Issue = 'WEAK_GENERIC_MODULE_TEXT'; Pattern = 'im Bereich X-ERP' },
    @{ Issue = 'WEAK_GENERIC_FORWARD_TEXT'; Pattern = 'wie die Daten in X-ERP weiterverwendet werden' }
)

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) {
        throw "Worksheet not found or empty: $WorksheetName"
    }

    $issues = New-Object System.Collections.Generic.List[object]
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $topic = [string]$ws.Cells.Item($r, 1).Text
        $contentType = [string]$ws.Cells.Item($r, 16).Text
        for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
            $text = [string]$ws.Cells.Item($r, $c).Text
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            foreach ($pattern in $patterns) {
                if ($text.Contains($pattern.Pattern)) {
                    $issues.Add([pscustomobject]@{
                        Row = $r
                        Col = $c
                        Header = [string]$ws.Cells.Item(1, $c).Text
                        ContentType = $contentType
                        Issue = $pattern.Issue
                        Topic = $topic
                        Text = $text
                    })
                    break
                }
            }
        }
    }

    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputCsv) | Out-Null
    $issues | Export-Csv -LiteralPath $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    [pscustomobject]@{
        Issues = $issues.Count
        OutputCsv = $OutputCsv
        First = @($issues | Select-Object -First 50)
    } | ConvertTo-Json -Depth 5
}
finally {
    if ($null -ne $pkg) { $pkg.Dispose() }
}
