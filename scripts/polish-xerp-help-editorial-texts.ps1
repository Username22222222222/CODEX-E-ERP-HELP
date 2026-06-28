param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$lowerAe = [char]0x00E4
$upperAe = [char]0x00C4
$lowerOe = [char]0x00F6
$upperOe = [char]0x00D6
$lowerUe = [char]0x00FC
$upperUe = [char]0x00DC
$ss = [char]0x00DF

function U([string]$text) {
    return $text.
        Replace('fuer', "f$($lowerUe)r").
        Replace('Fuer', "F$($lowerUe)r").
        Replace('Pruefen', "Pr$($lowerUe)fen").
        Replace('pruefen', "pr$($lowerUe)fen").
        Replace('Datensaetze', "Datens$($lowerAe)tze").
        Replace('Datensatzes', "Datensatzes").
        Replace('Schaltflaechen', "Schaltfl$($lowerAe)chen").
        Replace('zugehoerigen', "zugeh$($lowerOe)rigen").
        Replace('zugehoerig', "zugeh$($lowerOe)rig").
        Replace('gehoeren', "geh$($lowerOe)ren").
        Replace('gehoert', "geh$($lowerOe)rt").
        Replace("zugeh$($upperOe)rigen", "zugeh$($lowerOe)rigen").
        Replace("geh$($upperOe)rt", "geh$($lowerOe)rt").
        Replace('ergaenzende', "erg$($lowerAe)nzende").
        Replace('ergaenzt', "erg$($lowerAe)nzt").
        Replace('Verfuegbarkeit', "Verf$($lowerUe)gbarkeit").
        Replace('verfuegbar', "verf$($lowerUe)gbar").
        Replace('unterstuetzt', "unterst$($lowerUe)tzt").
        Replace('Unterstuetzt', "Unterst$($lowerUe)tzt").
        Replace('benoetigt', "ben$($lowerOe)tigt").
        Replace('Auspraegung', "Auspr$($lowerAe)gung").
        Replace('Auspraegungen', "Auspr$($lowerAe)gungen").
        Replace('ermoeglicht', "erm$($lowerOe)glicht").
        Replace('gewaehlten', "gew$($lowerAe)hlten").
        Replace('fuehrt', "f$($lowerUe)hrt").
        Replace('fuehren', "f$($lowerUe)hren").
        Replace('zaehlt', "z$($lowerAe)hlt").
        Replace('enthaelt', "enth$($lowerAe)lt").
        Replace('Enthaelt', "Enth$($lowerAe)lt").
        Replace('Uebersicht', "$($upperUe)bersicht").
        Replace('ueber', "$($lowerUe)ber").
        Replace('Ueber', "$($upperUe)ber").
        Replace(" $($upperUe)ber", " $($lowerUe)ber")
}

function Is-Weak([string]$text, [string]$topic) {
    if ([string]::IsNullOrWhiteSpace($text)) { return $true }
    $t = $text.Trim()
    if ($t.Length -lt 45) { return $true }
    if ($t -match 'Beschreibt das Feld') { return $true }
    if ($t -match 'werden die zugehoerigen Informationen|werden die zugehörigen Informationen') { return $true }
    if ($t -match '^Eingabefeld f') { return $true }
    return $false
}

function FieldText([string]$topic, [string]$field, [string]$tab, [string]$view) {
    $t = $topic.Trim()
    if ([string]::IsNullOrWhiteSpace($t)) { $t = $field.Trim() }
    $area = if ([string]::IsNullOrWhiteSpace($tab)) { 'diesem Bereich' } else { "dem Register $tab" }

    switch -Regex ($t) {
        '^Ort$' { return U("Der Ort gehoert zur Adresse des Datensatzes. Er unterstuetzt Anschrift, Suche und regionale Zuordnung.") }
        '^Land$' { return U("Das Land ordnet Adresse oder Datensatz einer Region zu. Es wird fuer Adressen, Formate, Steuern und Auswertungen verwendet.") }
        '^Fax$' { return U("Die Faxnummer ist eine optionale Kontaktangabe. Sie ergaenzt Telefon und E-Mail in der Kommunikation mit Partnern.") }
        '^Tel\.-Land$' { return U("Das Telefon-Land legt die Landesvorwahl bzw. regionale Telefonzuordnung fest und sorgt fuer einheitliche Kontaktdaten.") }
        'E-Mail|Email' { return U("Die E-Mail-Adresse wird fuer elektronische Kommunikation, Benachrichtigungen und eindeutige Kontaktzuordnung verwendet.") }
        'Telefon|Tel\.' { return U("Die Telefonnummer dient der direkten Kontaktaufnahme und wird in Stammdaten, Belegen und Auswertungen angezeigt.") }
        'Name|Bezeichnung|Matchcode' { return U("$t beschreibt den Datensatz so, dass Anwender ihn in Listen, Suchfeldern und Auswertungen schnell wiedererkennen.") }
        'Datum|Termin|Zeit|Gueltig|Gültig' { return U("$t legt einen Zeitpunkt oder Zeitraum fest. Die Angabe ist wichtig fuer Planung, Filterung und Nachvollziehbarkeit.") }
        'Preis|Kosten|Betrag|Aufschlag|Wert' { return U("$t fliesst in Preisfindung, Kalkulation oder betriebswirtschaftliche Auswertungen ein und sollte fachlich nachvollziehbar gepflegt werden.") }
        'Menge|Bestand|Schwelle|Losgr' { return U("$t steuert Mengen, Verfuegbarkeit oder Disposition. Der Wert hilft bei Planung, Lagerung und operativer Bearbeitung.") }
        'Nummer|ID|GTIN|Code|Warencode' { return U("$t ist eine Kennung fuer Suche, Zuordnung oder externe Referenzierung. Sie hilft, Datensaetze eindeutig wiederzufinden.") }
        'Gruppe|Kategorie|Typ|Status|Art' { return U("$t ordnet den Datensatz fachlich ein. Diese Zuordnung erleichtert Auswertung, Steuerung und strukturierte Bearbeitung.") }
        'Lager|Warehouse' { return U("$t ist eine lagerbezogene Angabe. Sie unterstuetzt Bestand, Verfuegbarkeit und logistische Prozesse.") }
        'Text|Bemerkung|Info|Beschreibung' { return U("$t enthaelt ergaenzende Informationen fuer Dokumentation, Druckausgabe oder interne Abstimmung.") }
        'Bild|Datei|Anhang' { return U("$t verknuepft den Datensatz mit visuellen Informationen oder Dateien, damit Anwender relevante Unterlagen direkt finden.") }
        'Faktor|Formel' { return U("$t steuert Berechnungen oder Umrechnungen. Der Wert sollte nachvollziehbar gepflegt werden, weil er Mengen, Preise oder Produktionsdaten beeinflussen kann.") }
        default { return U("$t ist eine fachliche Angabe in $area der Ansicht $view. Der Wert hilft Anwendern, den Datensatz korrekt zu pflegen, zu finden und weiterzuverarbeiten.") }
    }
}

function RegisterText([string]$tab, [string]$view) {
    switch -Regex ($tab) {
        'Schriftfarbe' { return U("Der Abschnitt Schriftfarbe steuert die Textfarbe fuer die Darstellung. So lassen sich wichtige Elemente visuell unterscheiden.") }
        'Hintergrundfarbe' { return U("Der Abschnitt Hintergrundfarbe steuert die Flaechenfarbe fuer die Darstellung und hilft, Elemente visuell zu strukturieren.") }
        'Rahmenfarbe' { return U("Der Abschnitt Rahmenfarbe legt die farbliche Begrenzung eines Elements fest und verbessert die visuelle Orientierung.") }
        'Hochgeladene Dateien' { return U("Im Register Hochgeladene Dateien verwalten Anwender Dateien, die zum Datensatz gehoeren und spaeter wieder auffindbar sein muessen.") }
        default { return U("Der Bereich $tab buendelt die Einstellungen und Informationen, die in der Ansicht $view fuer diesen fachlichen Abschnitt relevant sind.") }
    }
}

function ViewText([string]$view, [string]$module) {
    if ($view -match 'Wizard$') {
        return U("Der Assistent $view fuehrt Anwender durch einen klar abgegrenzten ERP-Prozess im Bereich $module. Die Seite erklaert Zweck, Eingaben und Aktionen Schritt fuer Schritt.")
    }
    return U("Mit der Bearbeitungsansicht $view pflegen Anwender Datensaetze im Bereich $module. Die Ansicht buendelt die relevanten Register, Felder und Aktionen fuer eine strukturierte Bearbeitung.")
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-editorial-polish-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $changed = 0
    $currentModule = ''
    $currentView = ''
    $currentTab = ''

    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $level = [int]$ws.Row($r).OutlineLevel
        $topic = [string]$ws.Cells.Item($r, 1).Text
        $field = [string]$ws.Cells.Item($r, 6).Text
        $text = [string]$ws.Cells.Item($r, 8).Text
        if ($level -eq 1 -and -not [string]::IsNullOrWhiteSpace($topic)) { $currentModule = $topic; $currentView = ''; $currentTab = '' }
        elseif ($level -eq 2 -and -not [string]::IsNullOrWhiteSpace($topic)) { $currentView = $topic; $currentTab = '' }
        elseif ($level -eq 3 -and [string]::IsNullOrWhiteSpace($field) -and -not [string]::IsNullOrWhiteSpace($topic)) { $currentTab = $topic }

        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        $newText = U $text
        if ($level -eq 2 -and ($topic -match 'Edit$|Wizard$') -and (Is-Weak $newText $topic)) {
            $newText = ViewText $topic $currentModule
        }
        elseif ($level -eq 3 -and [string]::IsNullOrWhiteSpace($field) -and (Is-Weak $newText $topic)) {
            $newText = RegisterText $topic $currentView
        }
        elseif (-not [string]::IsNullOrWhiteSpace($field) -and (Is-Weak $newText $topic)) {
            $newText = FieldText $topic $field $currentTab $currentView
        }

        if ($newText -ne $text) {
            $ws.Cells.Item($r, 8).Value = $newText
            $changed++
        }
    }

    Close-ExcelPackage $pkg
    [pscustomobject]@{
        Workbook = $WorkbookPath
        Backup = $backup
        ChangedTextCells = $changed
    } | ConvertTo-Json
}
catch {
    $pkg.Dispose()
    throw
}
