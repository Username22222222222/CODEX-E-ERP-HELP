param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$MissingCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\missing-explanations.csv"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

function Is-UsefulText([string]$text, [string]$topic) {
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    $t = $text.Trim()
    if ($t.Length -lt 18) { return $false }
    if ($t -eq $topic) { return $false }
    if ($t -match '^Eingabefeld fuer:') { return $false }
    if ($t -match '^Eingabefeld für:') { return $false }
    if ($t -match '^Lagerbezogene Angabe\.$') { return $false }
    if ($t -match '^Gruppenzuordnung\.$') { return $false }
    if ($t -match '^Preisangabe\.$') { return $false }
    return $true
}

function Get-ViewExplanation([string]$view, [string]$module) {
    if ($view.EndsWith('Wizard')) {
        return "Der Assistent $view fuehrt Anwender Schritt fuer Schritt durch den zugehoerigen ERP-Prozess im Bereich $module. Die darunterliegenden Register, Felder und Schaltflaechen beschreiben die einzelnen Eingaben und Aktionen."
    }
    return "Die Bearbeitungsansicht $view dient zum Anlegen, Pruefen und Bearbeiten von Datensaetzen im Bereich $module. Die darunterliegenden Registertabs gruppieren die Felder nach fachlichen Aufgaben."
}

function Get-TabExplanation([string]$tab, [string]$view) {
    switch -Regex ($tab) {
        '^Uebersicht$|^Übersicht$' { return "Im Register Uebersicht stehen die wichtigsten Grunddaten und Schluesselinformationen der Ansicht $view." }
        '^Details$' { return "Im Register Details werden erweiterte Stammdaten und fachliche Zusatzangaben zur Ansicht $view gepflegt." }
        '^Verkauf$' { return "Im Register Verkauf verwalten Sie verkaufsrelevante Einstellungen, Mengen, Preise und Zuordnungen." }
        '^Beschaffung$' { return "Im Register Beschaffung verwalten Sie einkaufs- und beschaffungsrelevante Daten wie Einkaufspreise, Lieferantenbezug und Kostenbestandteile." }
        '^Text$' { return "Im Register Text werden beschreibende Texte, Drucktexte und sprachliche Inhalte zur Ansicht gepflegt." }
        '^Bild$' { return "Im Register Bild werden Bilder oder visuelle Informationen zum Datensatz hinterlegt." }
        '^Lagerung$' { return "Im Register Lagerung werden Lagerbestand, Verfuegbarkeit, Bestandsgrenzen und lagerbezogene Einstellungen angezeigt oder gepflegt." }
        '^Set$' { return "Im Register Set verwalten Sie Zusammenstellungen oder Set-Bestandteile, die zu diesem Datensatz gehoeren." }
        '^Makro$' { return "Im Register Makro werden Makro- oder Automatisierungszuordnungen fuer den Datensatz verwaltet." }
        '^Produktion$' { return "Im Register Produktion werden produktionsrelevante Einstellungen und Zuordnungen fuer Fertigung und Arbeitsablaeufe gepflegt." }
        '^Zubehoer$|^Zubehör$' { return "Im Register Zubehoer werden zusaetzliche Artikel oder Komponenten verwaltet, die als Zubehoer zum Datensatz gehoeren." }
        '^Kategorien$' { return "Im Register Kategorien ordnen Sie den Datensatz einer oder mehreren Kategorien zu, damit er in Auswertungen, Navigation und Suche besser strukturiert ist." }
        '^Katalognummern$' { return "Im Register Katalognummern erfassen Sie externe oder partnerbezogene Katalog- und Artikelnummern. Dadurch kann der Datensatz ueber alternative Nummern gefunden und eindeutig zugeordnet werden." }
        '^Produktionsschritt-Ressourcen$' { return "Im Register Produktionsschritt-Ressourcen verwalten Sie Ressourcen, die fuer einzelne Produktionsschritte benoetigt werden." }
        '^Produktionsschritt-Artikel$' { return "Im Register Produktionsschritt-Artikel verwalten Sie Artikel oder Komponenten, die in Produktionsschritten verwendet werden." }
        '^Produktionszeit$' { return "Im Register Produktionszeit werden Zeitvorgaben und produktionszeitrelevante Einstellungen fuer den Datensatz gepflegt." }
        '^Positionsliste$' { return "Im Register Positionsliste sehen Sie Positionen oder Belegzeilen, in denen dieser Datensatz verwendet wird." }
        '^Umsatz$' { return "Im Register Umsatz werden umsatzbezogene Informationen und Auswertungen zum Datensatz angezeigt." }
        '^Vertraege$|^Verträge$' { return "Im Register Vertraege werden Vertragsbeziehungen und vertragliche Verwendungen des Datensatzes angezeigt oder gepflegt." }
        '^Lokal$' { return "Im Register Lokal werden landes-, standort- oder lokalspezifische Angaben zum Datensatz verwaltet." }
        '^Aenderungsprotokoll$|^Änderungsprotokoll$' { return "Im Register Aenderungsprotokoll sehen Sie nachvollziehbare Aenderungen am Datensatz." }
        '^Anhaenge$|^Anhänge$' { return "Im Register Anhaenge verwalten Sie Dateien und Dokumente, die zum Datensatz gehoeren." }
        '^Extra-Felder$' { return "Im Register Extra-Felder werden kundenspezifische Zusatzfelder gepflegt, die ueber den Standardumfang der Ansicht hinausgehen." }
        default { return "Im Register $tab werden die zugehoerigen Informationen der Ansicht $view gebuendelt. Die darunterliegenden Felder beschreiben die einzelnen Eingaben und Anzeigen dieses Bereichs." }
    }
}

function Get-FieldExplanation([string]$topic, [string]$field, [string]$tab, [string]$view, [string]$candidate) {
    if (Is-UsefulText $candidate $topic) { return $candidate }
    $t = $topic.Trim()
    $f = $field.Trim()

    if ($t -match '^Ist ' -or $t -match ' erlaubt$' -or $f -match '^(Is|Has|Can|Show|Use|Active|Favorite)') {
        return "Gibt an, ob die Option '$t' fuer diesen Datensatz aktiv ist. Diese Einstellung beeinflusst die Verarbeitung im Register $tab."
    }
    if ($t -match 'Datum|Termin|Gueltig|Gültig|Zeit') {
        return "Legt das relevante Datum oder den Zeitpunkt fuer '$t' fest. Diese Angabe wird fuer Planung, Filterung und Nachvollziehbarkeit verwendet."
    }
    if ($t -match 'Preis|Kosten|Aufschlag|Betrag|Wert') {
        return "Enthaelt den Wert fuer '$t'. Diese Angabe wird fuer Preisfindung, Kalkulation oder betriebswirtschaftliche Auswertung verwendet."
    }
    if ($t -match 'Menge|Bestand|Losgroesse|Losgröße|Schwelle') {
        return "Enthaelt die Mengenangabe '$t'. Sie wird fuer Verfuegbarkeit, Planung, Lagerung oder Disposition verwendet."
    }
    if ($t -match 'Nummer|ID|GTIN|Code|Warencode') {
        return "Speichert die Kennung '$t'. Diese Kennung dient der eindeutigen Zuordnung, Suche oder externen Referenzierung des Datensatzes."
    }
    if ($t -match 'Name|Bezeichnung|Matchcode') {
        return "Beschreibt den Datensatz ueber '$t'. Der Wert erleichtert Suche, Anzeige und Wiedererkennung in X-ERP."
    }
    if ($t -match 'Gruppe|Kategorie|Typ|Status|Art') {
        return "Ordnet den Datensatz ueber '$t' einer fachlichen Gruppe oder Auspraegung zu. Diese Zuordnung steuert Struktur, Auswertung oder Weiterverarbeitung."
    }
    if ($t -match 'Lager|Warehouse') {
        return "Enthaelt eine lagerbezogene Angabe fuer '$t'. Sie wird fuer Bestand, Verfuegbarkeit und logistische Prozesse verwendet."
    }
    if ($t -match 'Text|Bemerkung|Info|Beschreibung') {
        return "Enthaelt beschreibende Informationen zu '$t'. Diese Texte helfen bei Dokumentation, Druckausgabe oder interner Kommunikation."
    }
    if ($t -match 'Bild|Datei|Anhang') {
        return "Verweist auf eine Datei oder visuelle Information zu '$t'. Dadurch koennen ergaenzende Dokumente oder Bilder am Datensatz hinterlegt werden."
    }
    if ($t -match 'Faktor|Formel') {
        return "Definiert den Rechen- oder Umrechnungswert '$t'. Diese Angabe wird fuer Mengen-, Preis- oder Produktionsberechnungen verwendet."
    }
    return "Beschreibt das Feld '$t' in der Ansicht $view. Der Wert wird im Register $tab erfasst oder angezeigt und unterstuetzt die fachliche Bearbeitung des Datensatzes."
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
if (-not (Test-Path -LiteralPath $MissingCsv)) { throw "Missing audit CSV not found: $MissingCsv" }

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-explanation-fill-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$missing = @(Import-Csv -LiteralPath $MissingCsv -Delimiter ';' -Encoding UTF8)
$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $changed = 0
    foreach ($m in $missing) {
        $r = [int]$m.Row
        $topic = [string]$ws.Cells.Item($r, 1).Text
        $candidate = [string]$ws.Cells.Item($r, 2).Text
        $field = [string]$ws.Cells.Item($r, 6).Text
        $current = [string]$ws.Cells.Item($r, 8).Text
        if (Is-UsefulText $current $topic) { continue }

        $text = ''
        if ($m.Type -eq 'VIEW') { $text = Get-ViewExplanation $topic $m.Module }
        elseif ($m.Type -eq 'REGISTER_TAB') { $text = Get-TabExplanation $topic $m.View }
        elseif ($m.Type -eq 'FIELD') { $text = Get-FieldExplanation $topic $field $m.Tab $m.View $candidate }
        else { continue }

        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $ws.Cells.Item($r, 8).Value = $text
            $changed++
        }
    }
    Close-ExcelPackage $pkg
    [pscustomobject]@{
        Workbook = $WorkbookPath
        Backup = $backup
        ChangedExplanationCells = $changed
    } | ConvertTo-Json
}
catch {
    $pkg.Dispose()
    throw
}
