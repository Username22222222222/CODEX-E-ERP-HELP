param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\missing-explanations.csv"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

function Is-WeakExplanation([string]$text, [string]$topic) {
    if ([string]::IsNullOrWhiteSpace($text)) { return $true }
    $t = $text.Trim()
    if ($t.Length -lt 18) { return $true }
    if ($t -eq $topic) { return $true }
    if ($t -match '^Eingabefeld für:') { return $true }
    if ($t -match '^Lagerbezogene Angabe\.$') { return $true }
    if ($t -match '^Gruppenzuordnung\.$') { return $true }
    if ($t -match '^Preisangabe\.$') { return $true }
    return $false
}

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) { throw "Worksheet not found or empty: $WorksheetName" }

    $rows = New-Object System.Collections.Generic.List[object]
    $inViews = $false
    $currentModule = ''
    $currentView = ''
    $currentTab = ''

    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $level = [int]$ws.Row($r).OutlineLevel
        $topic = [string]$ws.Cells.Item($r, 1).Text
        $original = [string]$ws.Cells.Item($r, 2).Text
        $field = [string]$ws.Cells.Item($r, 6).Text
        $body = [string]$ws.Cells.Item($r, 8).Text

        if ($level -eq 0 -and $topic -match 'Ansichten|Views') { $inViews = $true }
        elseif ($level -eq 0 -and $inViews -and $topic -notmatch 'Ansichten|Views') { $inViews = $false }

        # Fallback: in der aktuellen Datei beginnt der Views-Block an den Modulen
        # mit technischen EditView-Unterzeilen. Dadurch wird auch ohne explizite
        # "Ansichten"-Ueberschrift sauber geprueft.
        if ($level -eq 2 -and $topic -match 'Edit$|Wizard$') { $inViews = $true }

        if (-not $inViews) { continue }

        if ($level -eq 1) { $currentModule = $topic; $currentView = ''; $currentTab = '' }
        elseif ($level -eq 2) { $currentView = $topic; $currentTab = '' }
        elseif ($level -eq 3 -and [string]::IsNullOrWhiteSpace($field)) { $currentTab = $topic }

        $isRelevant = $false
        $rowType = ''
        if ($level -eq 2 -and $topic -match 'Edit$|Wizard$') { $isRelevant = $true; $rowType = 'VIEW' }
        elseif ($level -eq 3 -and [string]::IsNullOrWhiteSpace($field) -and -not [string]::IsNullOrWhiteSpace($original)) { $isRelevant = $true; $rowType = 'REGISTER_TAB' }
        elseif ($level -ge 3 -and -not [string]::IsNullOrWhiteSpace($field)) { $isRelevant = $true; $rowType = 'FIELD' }

        if ($isRelevant -and (Is-WeakExplanation $body $topic)) {
            $rows.Add([pscustomobject]@{
                Row = $r
                Level = $level
                Type = $rowType
                Module = $currentModule
                View = $currentView
                Tab = $currentTab
                Topic = $topic
                Original = $original
                Field = $field
                CurrentText = $body
            })
        }
    }

    $outDir = Split-Path -Parent $OutputCsv
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    $rows | Export-Csv -LiteralPath $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    [pscustomobject]@{
        Workbook = $WorkbookPath
        MissingOrWeakExplanations = $rows.Count
        OutputCsv = $OutputCsv
    } | ConvertTo-Json
}
finally {
    $pkg.Dispose()
}
