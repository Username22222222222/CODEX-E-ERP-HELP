param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

function Split-Name([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return '' }
    $n = $name -replace 'Edit$','' -replace 'Wizard$',''
    $n = $n -replace '([a-zäöüß])([A-ZÄÖÜ])','$1 $2'
    $n = $n -replace '([A-ZÄÖÜ]+)([A-ZÄÖÜ][a-zäöüß])','$1 $2'
    return $n.Trim()
}

function Human-View([string]$viewName) {
    $base = Split-Name $viewName
    if ($viewName -match 'Wizard$') { return "$base-Assistent" }
    if ($viewName -match 'Edit$') { return "$base-Bearbeitungsansicht" }
    return $base
}

function Human-Module([string]$module) {
    if ([string]::IsNullOrWhiteSpace($module)) { return 'X-ERP' }
    switch -Regex ($module) {
        '^Admin' { return 'Administration' }
        '^Article$|^Artikel$' { return 'Artikelverwaltung' }
        '^Partner$' { return 'Partnerverwaltung' }
        '^Finance|^Finanz' { return 'Finanzwesen' }
        '^Warehouse|^Lager' { return 'Lagerverwaltung' }
        '^Production|^Produktion' { return 'Produktion' }
        '^Doc|^Dok' { return 'Belegwesen' }
        default { return $module }
    }
}

function Field-Text([string]$topic, [string]$field, [string]$tab, [string]$viewHuman) {
    $t = $topic.Trim()
    $area = if ([string]::IsNullOrWhiteSpace($tab)) { "in der $viewHuman" } else { "im Register $tab" }
    switch -Regex ($t) {
        '^PLZ$' { return "Die PLZ ist Teil der Adresse und unterstützt Suche, regionale Zuordnung und korrekte Anschriften." }
        '^Straße$' { return "Die Straße ist Teil der Adresse und wird für Anschriften, Belege und Kontaktinformationen verwendet." }
        '^Ort$' { return "Der Ort gehört zur Adresse und erleichtert Suche, regionale Zuordnung und eindeutige Kommunikation." }
        '^Land$' { return "Das Land steuert regionale Zuordnung, Formate und länderspezifische Auswertungen." }
        '^Fax$' { return "Die Faxnummer ergänzt die Kontaktdaten und bleibt für Partner relevant, die noch per Fax kommunizieren." }
        '^ID$|^Id$' { return "Die ID ist die eindeutige technische Kennung dieses Datensatzes und dient der sicheren Zuordnung in X-ERP." }
        '^ME$' { return "ME steht für Mengeneinheit und legt fest, in welcher Einheit Mengen erfasst, angezeigt oder berechnet werden." }
        '^EU$' { return "EU kennzeichnet einen Bezug zur Europäischen Union und unterstützt länderspezifische Regeln oder Auswertungen." }
        default { return "$t ist eine fachliche Angabe $area. Sie unterstützt Anwender dabei, den Datensatz korrekt zu pflegen, zu finden und im ERP-Prozess weiterzuverarbeiten." }
    }
}

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-humanize-view-texts-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $changed = 0
    $currentModule = ''
    $currentView = ''
    $currentViewHuman = ''
    $currentTab = ''

    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $level = [int]$ws.Row($r).OutlineLevel
        $topic = [string]$ws.Cells.Item($r, 1).Text
        $field = [string]$ws.Cells.Item($r, 6).Text
        $cell = $ws.Cells.Item($r, 8)
        $text = [string]$cell.Text
        if ($level -eq 1 -and -not [string]::IsNullOrWhiteSpace($topic)) { $currentModule = Human-Module $topic }
        if ($level -eq 2 -and -not [string]::IsNullOrWhiteSpace($topic)) {
            $currentView = $topic
            $currentViewHuman = Human-View $topic
            $currentTab = ''
        }
        if ($level -eq 3 -and [string]::IsNullOrWhiteSpace($field) -and -not [string]::IsNullOrWhiteSpace($topic)) { $currentTab = $topic }
        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        $new = $text
        if ($level -eq 2 -and ($topic -match 'Edit$|Wizard$')) {
            if ($topic -match 'Wizard$') {
                $new = "Der $currentViewHuman führt Anwender Schritt für Schritt durch den zugehörigen Prozess im Bereich $currentModule. Die Hilfe erklärt Zweck, Eingaben und Aktionen so, dass der Ablauf sicher nachvollzogen werden kann."
            } else {
                $new = "Die $currentViewHuman unterstützt Anwender beim Anlegen, Prüfen und Bearbeiten von Datensätzen im Bereich $currentModule. Die Hilfe beschreibt die Register, Felder und Aktionen aus fachlicher Sicht."
            }
        } else {
            if (-not [string]::IsNullOrWhiteSpace($currentView)) {
                $new = $new.Replace($currentView, $currentViewHuman)
            }
            if ($new -match 'Datensatz korrekt zu pflegen, zu finden und weiterzuverarbeiten' -or $new -match '^.+ ist eine fachliche Angabe in diesem Bereich') {
                $new = Field-Text $topic $field $currentTab $currentViewHuman
            }
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
