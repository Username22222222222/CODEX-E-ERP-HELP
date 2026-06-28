param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\split-word-issues.csv"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$textColumns = @(5, 8, 20, 21, 22, 23, 24)
$patterns = @(
    @{ Name = 'SPLIT_CAPITAL_TERM'; Regex = '\b[A-ZÄÖÜ][a-zäöüß]?\s+[a-zäöüß]{2}\s+[a-zäöüß]{2}\b' },
    @{ Name = 'SPLIT_TECHNICAL_NAME'; Regex = '\b[A-Z][A-Za-z0-9]{0,2}\s+[A-Za-z0-9]{2,}\s+(?:Edit|Wizard|Designer|Reports?)\b' },
    @{ Name = 'SPLIT_UMLAUT_FRAGMENT'; Regex = '\b(?:T\s*ex\s*tb|A\s*nh|S\s*t|R\s*c|E\s*rw|V\s*er|A\s*kt)\b' }
)

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) {
        throw "Worksheet not found or empty: $WorksheetName"
    }

    $issues = New-Object System.Collections.Generic.List[object]
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        foreach ($c in $textColumns) {
            $text = [string]$ws.Cells.Item($r, $c).Text
            if ([string]::IsNullOrWhiteSpace($text)) { continue }

            foreach ($pattern in $patterns) {
                if ($text -match $pattern.Regex) {
                    if ($pattern.Name -eq 'SPLIT_CAPITAL_TERM' -and $text -notmatch '\b\p{L}{1,2}\s+\p{L}{2}\s+\p{L}{2}\s+\p{L}{1,2}\b') {
                        continue
                    }
                    $issues.Add([pscustomobject]@{
                        Row = $r
                        Column = $c
                        Issue = $pattern.Name
                        Level = [int]$ws.Row($r).OutlineLevel
                        Topic = [string]$ws.Cells.Item($r, 1).Text
                        Title = [string]$ws.Cells.Item($r, 5).Text
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
        Workbook = $WorkbookPath
        Issues = $issues.Count
        OutputCsv = $OutputCsv
        First = @($issues | Select-Object -First 20)
    } | ConvertTo-Json -Depth 4
}
finally {
    $pkg.Dispose()
}
