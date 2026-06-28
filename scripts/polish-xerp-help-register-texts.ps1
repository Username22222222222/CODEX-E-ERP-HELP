param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$lowerUe = [char]0x00FC
$upperUe = [char]0x00DC
$lowerOe = [char]0x00F6
$lowerAe = [char]0x00E4

function Clean([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return $text }
    return $text.
        Replace(' gebuendelt', " geb$($lowerUe)ndelt").
        Replace(' buendelt', " b$($lowerUe)ndelt").
        Replace('Flaechen', "Fl$($lowerAe)chen").
        Replace('flaechen', "fl$($lowerAe)chen").
        Replace(" $($upperUe)ber", " $($lowerUe)ber")
}

function RegisterText([string]$tab, [string]$view) {
    switch -Regex ($tab) {
        'Schriftfarbe' { return "Der Abschnitt Schriftfarbe steuert die Textfarbe der Darstellung. Dadurch lassen sich wichtige Elemente visuell unterscheiden und konsistenter gestalten." }
        'Hintergrundfarbe' { return "Der Abschnitt Hintergrundfarbe steuert die Flächenfarbe der Darstellung. Das verbessert Orientierung, Hervorhebung und optische Struktur." }
        'Rahmenfarbe' { return "Der Abschnitt Rahmenfarbe legt die farbliche Begrenzung eines Elements fest. So werden Bereiche, Buttons oder Menüpunkte klarer voneinander getrennt." }
        'Hochgeladene Dateien' { return "Im Register Hochgeladene Dateien verwalten Anwender Dateien, die zum Datensatz gehören und später wieder auffindbar sein müssen." }
        'Zubehör' { return "Im Register Zubehör werden ergänzende Artikel oder Komponenten gepflegt, die zusammen mit dem aktuellen Artikel verwendet, angeboten oder dokumentiert werden." }
        default { return "Der Bereich $tab bündelt die fachlich passenden Einstellungen der Ansicht $view. Anwender finden hier die Felder, die für diesen Abschnitt gemeinsam bearbeitet oder geprüft werden." }
    }
}

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-register-polish-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $changed = 0
    $currentView = ''
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $level = [int]$ws.Row($r).OutlineLevel
        $topic = [string]$ws.Cells.Item($r, 1).Text
        $field = [string]$ws.Cells.Item($r, 6).Text
        $cell = $ws.Cells.Item($r, 8)
        $text = [string]$cell.Text
        if ($level -eq 2 -and -not [string]::IsNullOrWhiteSpace($topic)) { $currentView = $topic }
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        $new = Clean $text
        if ($level -eq 3 -and [string]::IsNullOrWhiteSpace($field) -and $new.Contains('Informationen der Ansicht')) {
            $new = RegisterText $topic $currentView
        }
        if ($new -cne $text) {
            $cell.Value = $new
            $changed++
        }
    }
    Close-ExcelPackage $pkg
    [pscustomobject]@{
        Workbook = $WorkbookPath
        Backup = $backup
        ChangedCells = $changed
    } | ConvertTo-Json
}
catch {
    $pkg.Dispose()
    throw
}
