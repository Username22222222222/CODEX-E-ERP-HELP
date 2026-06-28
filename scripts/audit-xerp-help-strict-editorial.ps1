param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\strict-editorial-issues.csv"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) { throw "Worksheet not found or empty: $WorksheetName" }
    $issues = New-Object System.Collections.Generic.List[object]

    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $level = [int]$ws.Row($r).OutlineLevel
        $topic = [string]$ws.Cells.Item($r, 1).Text
        $field = [string]$ws.Cells.Item($r, 6).Text
        $text = [string]$ws.Cells.Item($r, 8).Text
        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        if ($text -match '\b[A-Z][A-Za-z0-9]+(?:Edit|Wizard)\b') {
            $issues.Add([pscustomobject]@{ Row=$r; Issue='TECHNICAL_VIEW_NAME_IN_TEXT'; Level=$level; Topic=$topic; Field=$field; Text=$text })
        }
        elseif ($text -match '^Der Bereich .+ bündelt die fachlich passenden Einstellungen') {
            $issues.Add([pscustomobject]@{ Row=$r; Issue='GENERIC_REGISTER_TEXT'; Level=$level; Topic=$topic; Field=$field; Text=$text })
        }
        elseif ($text -match '^.+ ist eine fachliche Angabe in diesem Bereich') {
            $issues.Add([pscustomobject]@{ Row=$r; Issue='GENERIC_FIELD_TEXT'; Level=$level; Topic=$topic; Field=$field; Text=$text })
        }
        elseif ($text -match 'Datensatz korrekt zu pflegen, zu finden und weiterzuverarbeiten') {
            $issues.Add([pscustomobject]@{ Row=$r; Issue='GENERIC_FIELD_TEXT'; Level=$level; Topic=$topic; Field=$field; Text=$text })
        }
    }

    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputCsv) | Out-Null
    $issues | Export-Csv -LiteralPath $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    $issues | Group-Object Issue | Select-Object Name,Count | ConvertTo-Json
}
finally {
    $pkg.Dispose()
}
