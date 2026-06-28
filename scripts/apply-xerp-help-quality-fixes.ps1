param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.sorted.xlsx",
    [string]$WorksheetName = "de-DE"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$ue = [char]0x00FC
$ss = [char]0x00DF
$mojibakeStreet = 'Stra' + [char]195 + [char]376 + 'e'
$street = 'Stra' + $ss + 'e'

$replacements = [ordered]@{
    ('Rollen-Men' + $ue) = ('Rollenmen' + $ue)
    'Artikel-Details' = 'Artikeldetails'
    'Artikel-Kategorien' = 'Artikelkategorien'
    'Profit Betrag' = 'Gewinn Betrag'
    'Profit Prozent' = 'Gewinn Prozent'
    'Strasse' = $street
    $mojibakeStreet = $street
    'Mitarbeiter-Gruppe' = 'Mitarbeitergruppe'
    'Kontaktpersonen' = 'Ansprechpartner'
    'Mahnliste' = 'Mahnungsliste'
    'Dashboard-Designer' = 'Dashboarddesigner'
    'Report-Designer' = 'Reportdesigner'
    'Supporteinstellungen' = 'Support-Einstellungen'
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-quality-fixes-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) { throw "Worksheet not found or empty: $WorksheetName" }

    $changed = New-Object System.Collections.Generic.List[object]
    for ($r = 1; $r -le $ws.Dimension.End.Row; $r++) {
        for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
            $cell = $ws.Cells.Item($r, $c)
            $text = [string]$cell.Text
            if ($replacements.Contains($text)) {
                $newText = $replacements[$text]
                $cell.Value = $newText
                $changed.Add([pscustomobject]@{
                    Row = $r
                    Column = $c
                    OldText = $text
                    NewText = $newText
                })
            }
        }
    }

    Close-ExcelPackage $pkg

    [pscustomobject]@{
        Workbook = $WorkbookPath
        Backup = $backup
        ChangedCells = $changed.Count
        Changes = $changed.ToArray()
    } | ConvertTo-Json -Depth 6
}
catch {
    $pkg.Dispose()
    throw
}
