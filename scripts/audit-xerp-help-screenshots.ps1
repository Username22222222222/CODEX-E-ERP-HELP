param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\screenshot-audit.csv"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) { throw "Worksheet not found or empty: $WorksheetName" }

    $headers = @{}
    for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
        $h = [string]$ws.Cells.Item(1, $c).Text
        if (-not [string]::IsNullOrWhiteSpace($h)) { $headers[$h] = $c }
    }

    foreach ($required in @('Screenshot', 'SCREENSHOT_REL_PATH', 'SCREENSHOT_WEB_PATH', 'IMAGE_ALT', 'IMAGE_CAPTION', 'IMAGE_STATUS')) {
        if (-not $headers.ContainsKey($required)) { throw "Column missing: $required" }
    }

    $rows = New-Object System.Collections.Generic.List[object]
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $status = [string]$ws.Cells.Item($r, $headers['IMAGE_STATUS']).Text
        if ($status -eq 'kein Screenshot') { continue }
        $rows.Add([pscustomobject]@{
            Row = $r
            Topic = [string]$ws.Cells.Item($r, 1).Text
            Original = [string]$ws.Cells.Item($r, 2).Text
            Screenshot = [string]$ws.Cells.Item($r, $headers['Screenshot']).Text
            RelativePath = [string]$ws.Cells.Item($r, $headers['SCREENSHOT_REL_PATH']).Text
            WebPath = [string]$ws.Cells.Item($r, $headers['SCREENSHOT_WEB_PATH']).Text
            Alt = [string]$ws.Cells.Item($r, $headers['IMAGE_ALT']).Text
            Caption = [string]$ws.Cells.Item($r, $headers['IMAGE_CAPTION']).Text
            Status = $status
        })
    }

    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputCsv) | Out-Null
    $rows | Export-Csv -LiteralPath $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    [pscustomobject]@{
        OutputCsv = $OutputCsv
        Rows = $rows.Count
        Existing = @($rows | Where-Object { $_.Status -eq 'vorhanden' }).Count
        Missing = @($rows | Where-Object { $_.Status -eq 'fehlt' }).Count
    } | ConvertTo-Json
}
finally {
    $pkg.Dispose()
}
