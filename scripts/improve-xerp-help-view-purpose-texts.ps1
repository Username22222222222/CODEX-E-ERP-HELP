param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$BackupRoot = "C:\Users\micha\Documents\X-ERP-HELP\ARCHIV"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$chr_lower_ae = [string][char]0x00E4
$chr_lower_oe = [string][char]0x00F6
$chr_lower_ue = [string][char]0x00FC
$chr_upper_ae = [string][char]0x00C4
$chr_upper_oe = [string][char]0x00D6
$chr_upper_ue = [string][char]0x00DC
$chr_sz = [string][char]0x00DF

function Split-TechnicalName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $x = $Name -replace '(Edit|Wizard|Designer)$', ''
    $x = $x -creplace '([a-z0-9])([A-Z])', '$1 $2'
    $x = $x -creplace '([A-Z]+)([A-Z][a-z])', '$1 $2'
    $x = $x -replace 'Uo M', 'UoM'
    $x = $x -replace 'E Banking', 'E-Banking'
    $x = $x -replace 'E Invoice', 'E-Rechnung'
    $x = $x -replace 'Ai ', 'KI '
    $x = $x -replace '\s+', ' '
    return $x.Trim()
}

function Get-ModuleName {
    param([string]$Folder, [string]$Topic)
    $source = "$Topic $Folder"
    switch -Regex ($source) {
        'Admin' { return 'Administration' }
        'Article|Artikel' { return 'Artikelverwaltung' }
        'Partner' { return 'Partnerverwaltung' }
        'Finance|Financial|Finanz' { return 'Finanzwesen' }
        'Procurement|Einkauf' { return 'Einkauf und Beschaffung' }
        'Production|Produktion' { return 'Produktion' }
        'Calendar|Kalender' { return 'Kalender' }
        'Email|E-Mail' { return 'E-Mail' }
        'Crm|CRM' { return 'CRM' }
        'Intercom' { return 'Intercom' }
        'CountryPackage' { return 'L' + $chr_lower_ae + 'nderpakete' }
        'Employee|Mitarbeiter' { return 'Mitarbeiterverwaltung' }
        'Support' { return 'Support und Service' }
        'TimeTracking|Zeiterfassung' { return 'Zeiterfassung' }
        'Shipping|Versand' { return 'Versand' }
        'Workspace' { return 'Arbeitsplatz und Rollen' }
        'System' { return 'Systemverwaltung' }
        'Text|Translation' { return 'Texte und ' + $chr_upper_ue + 'bersetzungen' }
        'Dashboard|Report' { return 'Auswertung und Reporting' }
        default { return 'X-ERP' }
    }
}

function Get-ObjectLabel {
    param([string]$Topic)
    $name = Split-TechnicalName $Topic
    $map = [ordered]@{
        'Admin Bank Account E-Banking Type' = 'E-Banking-Typen f' + $chr_lower_ue + 'r Bankkonten'
        'Admin Bank' = 'Bankstammdaten'
        'Admin Company Setting' = 'firmenweite Grundeinstellungen'
        'Admin Country' = 'L' + $chr_lower_ae + 'nderstammdaten'
        'Admin Currency Api' = 'Schnittstellen f' + $chr_lower_ue + 'r W' + $chr_lower_ae + 'hrungskurse'
        'Admin Currency Rate' = 'Wechselkurse'
        'Admin Currency' = 'W' + $chr_lower_ae + 'hrungen'
        'Admin Language' = 'Sprachen'
        'Admin Location' = 'Standorte'
        'Admin Program Setting' = 'Programmeinstellungen'
        'Admin Route' = 'Systemrouten'
        'Admin Svg' = 'SVG-Symbole'
        'Admin Time Zone' = 'Zeitzonen'
        'Admin Year' = 'Gesch' + $chr_lower_ae + 'ftsjahre'
        'Admin Zip' = 'Postleitzahlen'
        'Admin Program Ai Language' = 'KI-Sprachvorgaben'
        'Attachment Status' = 'Statuswerte f' + $chr_lower_ue + 'r Anh' + $chr_lower_ae + 'nge'
        'Attachment' = 'Anh' + $chr_lower_ae + 'nge und Wiedervorlagen'
        'Article Accessory' = 'Artikelzubeh' + $chr_lower_oe + 'r'
        'Article Category Name' = 'mehrsprachige Bezeichnungen f' + $chr_lower_ue + 'r Artikelkategorien'
        'Article Category' = 'Artikelkategorien'
        'Article Group' = 'Artikelgruppen'
        'Article Sales Price List Name' = 'mehrsprachige Bezeichnungen f' + $chr_lower_ue + 'r Verkaufspreislisten'
        'Article Sales Price List' = 'Verkaufspreislisten'
        'Article Price Unit' = 'Preiseinheiten'
        'Article Production Step' = 'Produktionsschritte'
        'Article Template' = 'Artikelvorlagen'
        'Article UoM Conversion' = 'Umrechnungen zwischen Mengeneinheiten'
        'Article UoM' = 'Mengeneinheiten f' + $chr_lower_ue + 'r Artikel'
        'Article Set' = 'Artikelsets'
        'Article Macro' = 'Artikelmakros'
        'Article' = 'Artikelstammdaten'
        'Partner Bank Account' = 'Bankverbindungen von Partnern'
        'Partner Contact Person Email Category Name' = 'mehrsprachige Bezeichnungen f' + $chr_lower_ue + 'r E-Mail-Kategorien von Ansprechpartnern'
        'Partner Contact Person' = 'Ansprechpartner'
        'Partner Delivery Address' = 'Lieferadressen'
        'Partner Billing Address' = 'Rechnungsadressen'
        'Partner Terms Of Payment' = 'Zahlungsbedingungen'
        'Partner Terms Of Delivery' = 'Lieferbedingungen'
        'Partner Discount Group Name' = 'mehrsprachige Bezeichnungen f' + $chr_lower_ue + 'r Rabattgruppen'
        'Partner Discount Group' = 'Rabattgruppen'
        'Partner Group' = 'Partnergruppen'
        'Partner Category Name' = 'mehrsprachige Bezeichnungen f' + $chr_lower_ue + 'r Partnerkategorien'
        'Partner Category' = 'Partnerkategorien'
        'Partner' = 'Kunden, Lieferanten und weitere Gesch' + $chr_lower_ae + 'ftspartner'
        'Finance Account Plan' = 'Kontenpl' + $chr_lower_ae + 'ne'
        'Finance Account Group' = 'Kontengruppen'
        'Finance Account Determination' = 'Kontenfindung'
        'Finance Account Internal Reconciliation' = 'interne Kontoabstimmungen'
        'Finance Open Item Internal Reconciliation' = 'interne Abstimmungen offener Posten'
        'Finance Open Item On Partner Page' = 'offene Posten direkt am Partner'
        'Finance Open Item' = 'offene Posten'
        'Finance Payment' = 'Zahlungen'
        'Finance Posting Period' = 'Buchungsperioden'
        'Finance Posting Template' = 'Buchungsvorlagen'
        'Finance Posting Draft' = 'Buchungsentw' + $chr_lower_ue + 'rfe'
        'Finance Posting' = 'Buchungen'
        'Finance Tax Key' = 'Steuerschl' + $chr_lower_ue + 'ssel'
        'Finance Transfer Position' = 'Bankbewegungen und Zahlungspositionen'
        'Finance Transfer' = $chr_upper_ue + 'berweisungen und Banktransfers'
        'Finance Dunning' = 'Mahnungen'
        'Finance Fixed Asset' = 'Anlagenstammdaten'
        'Calendar Appointment' = 'Termine'
        'Calendar Working Day' = 'Arbeitstage'
        'Calendar Holiday' = 'Feiertage'
        'Calendar Name' = 'Kalenderbezeichnungen'
        'Calendar' = 'Kalender'
        'Email Account' = 'E-Mail-Konten'
        'Email Signature' = 'E-Mail-Signaturen'
        'Email Template' = 'E-Mail-Vorlagen'
        'Email Rule In' = 'Regeln f' + $chr_lower_ue + 'r eingehende E-Mails'
        'Email Rule Out' = 'Regeln f' + $chr_lower_ue + 'r ausgehende E-Mails'
        'Email' = 'E-Mails'
        'Employee Vacation Request Approval' = 'Genehmigungen von Urlaubsantr' + $chr_lower_ae + 'gen'
        'Employee Vacation' = 'Urlaub'
        'Employee Commission Group Name' = 'mehrsprachige Bezeichnungen f' + $chr_lower_ue + 'r Provisionsgruppen'
        'Employee Commission Group' = 'Provisionsgruppen'
        'Employee Group' = 'Mitarbeitergruppen'
        'Employee' = 'Mitarbeiter'
        'Extra Field' = 'Extra-Felder'
        'Extra Table' = 'Extra-Tabellen'
        'Extra Plugin' = 'Plugins'
        'Extra Api' = 'API-Erweiterungen'
        'Output Control Printer Configuration' = 'Drucker- und Ausgabekonfigurationen'
        'Output Control Format' = 'Ausgabeformate'
        'Doc' = 'Belege'
        'Procurement Planning Ask' = 'Beschaffungsanfragen'
        'Procurement Planning Buy' = 'Einkaufsbedarfe'
        'Procurement Planning Make' = 'Fertigungsbedarfe'
        'Bulletin Category' = 'Mitteilungskategorien'
        'Bulletin' = 'Mitteilungen'
        'XX Table' = 'technische Test- und Beispieldaten'
        'Formatted Search' = 'formatierte Suchen'
        'Sticky Note' = 'Notizen'
        'Intercom Recipient Status' = 'Empf' + $chr_lower_ae + 'ngerstatus f' + $chr_lower_ue + 'r interne Nachrichten'
        'Intercom Sender Status' = 'Absenderstatus f' + $chr_lower_ue + 'r interne Nachrichten'
        'Intercom Group' = 'Intercom-Gruppen'
        'Intercom Prio' = 'Priorit' + $chr_lower_ae + 'ten f' + $chr_lower_ue + 'r interne Nachrichten'
        'Intercom' = 'interne Nachrichten'
        'Ai Text Info' = 'KI-Textinformationen'
        'Packing List Box Position' = 'Positionen in Packst' + $chr_lower_ue + 'cken'
        'Packing List Box Type' = 'Packst' + $chr_lower_ue + 'cktypen'
        'Packing List Box' = 'Packst' + $chr_lower_ue + 'cke f' + $chr_lower_ue + 'r Kommissionierlisten'
        'Project' = 'Projekte'
        'Text Block Group' = 'Textbaustein-Gruppen'
        'Shipping Carrier' = 'Frachtf' + $chr_lower_ue + 'hrer'
        'Shipping Rate' = 'Versandtarife'
        'Shipping Type' = 'Versandarten'
        'Shipping Zone' = 'Versandzonen'
        'Workspace Role Desktop' = 'rollenbasierte Desktop-Zugriffe'
        'Workspace Role Menu' = 'rollenbasierte Men' + $chr_lower_ue + 'zugriffe'
        'System View' = 'Systemansichten'
        'System Table Field' = 'Tabellenfelder'
        'Text Block' = 'Textbausteine'
        'Text Variable' = 'Textvariablen'
        'Translation' = $chr_upper_ue + 'bersetzungen'
        'Wiki Category' = 'Wiki-Kategorien'
        'Wiki Status' = 'Wiki-Statuswerte'
        'Wiki' = 'Wiki-Beitr' + $chr_lower_ae + 'ge'
        'Time Tracking' = 'Zeiterfassung'
    }
    foreach ($entry in $map.GetEnumerator()) {
        if ($name.StartsWith($entry.Key, [StringComparison]::OrdinalIgnoreCase)) { return $entry.Value }
    }
    return $name
}

function Get-PurposeText {
    param(
        [string]$Topic,
        [string]$Folder,
        [string[]]$Fields
    )

    $module = Get-ModuleName -Folder $Folder -Topic $Topic
    $label = Get-ObjectLabel $Topic
    $fieldHint = ''
    $topFields = @($Fields | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 4)
    if ($topFields.Count -gt 0) {
        $fieldHint = ' Wichtige Angaben sind ' + (($topFields -join ', ') -replace ', ([^,]+)$', ' und $1') + '.'
    }

    switch -Regex ($Topic) {
        'BankAccountEBankingType' {
            return 'Diese Ansicht legt fest, welche E-Banking-Verfahren f' + $chr_lower_ue + 'r Bankkonten verwendet werden. Damit bleiben Bankanbindung, Zahlungsverkehr und Kontoauszugsimport in X-ERP eindeutig klassifiziert.'
        }
        'ArticleCategoryNameEdit$' {
            return 'Diese Ansicht pflegt die sprachabh' + $chr_lower_ae + 'ngigen Bezeichnungen von Artikelkategorien. Dadurch erscheinen Kategorien in Listen, Belegen und Auswertungen in der passenden Sprache und bleiben trotz mehrsprachiger Oberfl' + $chr_lower_ae + 'che eindeutig zugeordnet.' + $fieldHint
        }
        'ArticleSalesPriceListNameEdit$' {
            return 'Diese Ansicht pflegt die sprachabh' + $chr_lower_ae + 'ngigen Bezeichnungen von Verkaufspreislisten. So k' + $chr_lower_oe + 'nnen Preislisten intern eindeutig verwaltet und in mehrsprachigen Verkaufsprozessen verst' + $chr_lower_ae + 'ndlich angezeigt werden.' + $fieldHint
        }
        'AdminBankEdit$' {
            return 'Diese Ansicht pflegt Banken als Stammdaten f' + $chr_lower_ue + 'r Zahlungsverkehr, Bankverbindungen und Finanzprozesse. BIC, Name und Adresse sorgen daf' + $chr_lower_ue + 'r, dass Bankkonten in Belegen, Zahlungen und Auswertungen eindeutig zugeordnet werden k' + $chr_lower_oe + 'nnen.' + $fieldHint
        }
        'AdminCountryEdit$' {
            return 'Diese Ansicht verwaltet L' + $chr_lower_ae + 'nderstammdaten f' + $chr_lower_ue + 'r Adressen, Steuern, Telefonformate und internationale Abl' + $chr_lower_ae + 'ufe. Sie schafft die Grundlage daf' + $chr_lower_ue + 'r, dass Kunden, Lieferanten, Belege und landesspezifische Regeln korrekt zusammenarbeiten.' + $fieldHint
        }
        'AdminCurrencyEdit$' {
            return 'Diese Ansicht verwaltet W' + $chr_lower_ae + 'hrungen f' + $chr_lower_ue + 'r Angebote, Auftr' + $chr_lower_ae + 'ge, Einkauf, Buchhaltung und Auswertungen. Der W' + $chr_lower_ae + 'hrungscode stellt sicher, dass Betr' + $chr_lower_ae + 'ge im ERP eindeutig bewertet und umgerechnet werden k' + $chr_lower_oe + 'nnen.' + $fieldHint
        }
        'AdminLanguageEdit$' {
            return 'Diese Ansicht verwaltet Sprachen f' + $chr_lower_ue + 'r Oberfl' + $chr_lower_ae + 'che, Belege, Texte und mehrsprachige Stammdaten. Sie ist die Grundlage daf' + $chr_lower_ue + 'r, dass X-ERP Inhalte in der passenden Sprache anzeigen und ausgeben kann.' + $fieldHint
        }
        'AdminLocationEdit$' {
            return 'Diese Ansicht verwaltet Standorte, damit Lager, Organisationseinheiten, Belege und operative Prozesse einem betrieblichen Ort zugeordnet werden k' + $chr_lower_oe + 'nnen.' + $fieldHint
        }
        'AdminProgramAiLanguageEdit$' {
            return 'Diese Ansicht pflegt Sprachvorgaben f' + $chr_lower_ue + 'r KI-gest' + $chr_lower_ue + 'tzte Texte. Damit kann X-ERP steuern, in welcher Sprache automatische Beschreibungen, Hinweise oder redaktionelle Inhalte vorbereitet werden.'
        }
        'AdminSvgEdit$' {
            return 'Diese Ansicht verwaltet SVG-Symbole, die X-ERP in Oberfl' + $chr_lower_ae + 'che, Men' + $chr_lower_ue + 's und Ausgaben verwendet. Eindeutige Namen und sauberer SVG-Code sorgen f' + $chr_lower_ue + 'r konsistente Icons im System.' + $fieldHint
        }
        'AdminTimeZoneEdit$' {
            return 'Diese Ansicht verwaltet Zeitzonen f' + $chr_lower_ue + 'r Termine, Zeitstempel, internationale Benutzer und automatische Prozesse. Die IANA-Zuordnung sorgt daf' + $chr_lower_ue + 'r, dass Zeiten standort' + $chr_lower_ue + 'bergreifend korrekt interpretiert werden.' + $fieldHint
        }
        'AdminYearEdit$' {
            return 'Diese Ansicht verwaltet Gesch' + $chr_lower_ae + 'ftsjahre als zeitliche Grundlage f' + $chr_lower_ue + 'r Buchhaltung, Auswertungen, Perioden und betriebliche Planung.' + $fieldHint
        }
        'AdminZipEdit$' {
            return 'Diese Ansicht verwaltet Postleitzahlen und Orte f' + $chr_lower_ue + 'r Adressen, Versand, Partnerstammdaten und regionale Auswertungen.' + $fieldHint
        }
        'AttachmentStatusEdit$' {
            return 'Diese Ansicht verwaltet Statuswerte f' + $chr_lower_ue + 'r Anh' + $chr_lower_ae + 'nge. Damit k' + $chr_lower_oe + 'nnen Dokumente, Dateien und Wiedervorlagen nachvollziehbar klassifiziert und im Arbeitsablauf bewertet werden.' + $fieldHint
        }
        'AttachmentEdit$' {
            return 'Diese Ansicht verwaltet Anh' + $chr_lower_ae + 'nge, Notizen und Wiedervorlagen zu Datens' + $chr_lower_ae + 'tzen. Sie macht Dateien und Zusatzinformationen dort verf' + $chr_lower_ue + 'gbar, wo Anwender sie im ERP-Prozess ben' + $chr_lower_oe + 'tigen.' + $fieldHint
        }
        'CurrencyRate' {
            return 'Diese Ansicht pflegt Wechselkurse f' + $chr_lower_ue + 'r W' + $chr_lower_ae + 'hrungen, damit Belege, Bewertungen und Auswertungen mit nachvollziehbaren Kursen arbeiten.' + $fieldHint
        }
        'CurrencyApi' {
            return 'Diese Ansicht konfiguriert die Schnittstelle, aus der X-ERP W' + $chr_lower_ae + 'hrungskurse beziehen kann. So lassen sich Kursquellen, Abruf und Aktualisierung zentral steuern.'
        }
        'FinanceOpenItemOnPartnerPageEdit$' {
            return 'Diese Ansicht zeigt und bearbeitet offene Posten direkt im Partnerkontext. Anwender sehen damit Forderungen, Verbindlichkeiten und Zahlungsstatus dort, wo Kunden- oder Lieferanteninformationen gepflegt werden.' + $fieldHint
        }
        'CalendarNameEdit$' {
            return 'Diese Ansicht pflegt Kalenderbezeichnungen f' + $chr_lower_ue + 'r mehrsprachige oder organisatorisch getrennte Kalender. Dadurch bleiben Arbeits-, Termin- und Planungskalender f' + $chr_lower_ue + 'r Anwender eindeutig benannt.' + $fieldHint
        }
        'EmployeeCommissionGroupNameEdit$' {
            return 'Diese Ansicht pflegt sprachabh' + $chr_lower_ae + 'ngige Bezeichnungen von Provisionsgruppen. Sie sorgt daf' + $chr_lower_ue + 'r, dass Verg' + $chr_lower_ue + 'tungs- und Vertriebszuordnungen in mehrsprachigen Oberfl' + $chr_lower_ae + 'chen verst' + $chr_lower_ae + 'ndlich bleiben.' + $fieldHint
        }
        'EmployeeVacationRequestApprovalEdit$' {
            return 'Diese Ansicht unterst' + $chr_lower_ue + 'tzt die Pr' + $chr_lower_ue + 'fung und Genehmigung von Urlaubsantr' + $chr_lower_ae + 'gen. Sie macht sichtbar, welcher Antrag entschieden werden muss und wie die Abwesenheit in die Personal- und Kalenderplanung einflie' + $chr_sz + 't.' + $fieldHint
        }
        'PartnerCategoryNameEdit$' {
            return 'Diese Ansicht pflegt sprachabh' + $chr_lower_ae + 'ngige Bezeichnungen von Partnerkategorien. Damit werden Kunden, Lieferanten und weitere Gesch' + $chr_lower_ae + 'ftspartner in mehrsprachigen Listen und Auswertungen klar gruppiert.' + $fieldHint
        }
        'PartnerContactPersonEmailCategoryNameEdit$' {
            return 'Diese Ansicht pflegt sprachabh' + $chr_lower_ae + 'ngige Bezeichnungen f' + $chr_lower_ue + 'r E-Mail-Kategorien von Ansprechpartnern. Dadurch lassen sich Kontaktadressen und Kommunikationsarten auch in mehrsprachigen Partnerdaten sauber einordnen.' + $fieldHint
        }
        'PartnerDiscountGroupNameEdit$' {
            return 'Diese Ansicht pflegt sprachabh' + $chr_lower_ae + 'ngige Bezeichnungen von Rabattgruppen. So bleiben Preis- und Konditionsgruppen f' + $chr_lower_ue + 'r Vertrieb, Einkauf und Auswertung eindeutig verst' + $chr_lower_ae + 'ndlich.' + $fieldHint
        }
        'OutputControlFormatEdit$' {
            return 'Diese Ansicht definiert Ausgabeformate f' + $chr_lower_ue + 'r Belege, Druck, Archivierung und elektronische Ausgaben. Damit wird festgelegt, wie X-ERP Dokumente erzeugt und an den passenden Ausgabekanal weitergibt.' + $fieldHint
        }
        'OutputControlPrinterConfigurationEdit$' {
            return 'Diese Ansicht verwaltet Drucker- und Ausgabekonfigurationen. Sie steuert, welche Drucker, Mengen und Ausgabewege f' + $chr_lower_ue + 'r Belege oder interne Dokumente verwendet werden.' + $fieldHint
        }
        'DocEdit$' {
            return 'Diese Ansicht verwaltet Belege und deren zentrale Kopfdaten. Rechnungsadresse, Matchcode und Dokumenttyp verbinden den Beleg mit Partner, Prozess und nachgelagerten ERP-Funktionen.' + $fieldHint
        }
        'ProcurementPlanningAskEdit$' {
            return 'Diese Ansicht bearbeitet Beschaffungsanfragen aus der Planung. Sie zeigt, welcher Bedarf offen ist, wie er best' + $chr_lower_ae + 'tigt wurde und ob daraus Einkauf oder Fertigung entstehen soll.' + $fieldHint
        }
        'ProcurementPlanningBuyEdit$' {
            return 'Diese Ansicht bearbeitet Einkaufsbedarfe aus der Beschaffungsplanung. Sie hilft, Artikel, Kundenbezug und Auftragsbest' + $chr_lower_ae + 'tigung in konkrete Einkaufsprozesse zu ' + $chr_lower_ue + 'berf' + $chr_lower_ue + 'hren.' + $fieldHint
        }
        'ProcurementPlanningMakeEdit$' {
            return 'Diese Ansicht bearbeitet Fertigungsbedarfe aus der Beschaffungsplanung. Sie verbindet Kundenauftrag, Artikel und offene Menge mit der Entscheidung, was intern produziert werden muss.' + $fieldHint
        }
        'BulletinCategoryEdit$' {
            return 'Diese Ansicht verwaltet Kategorien f' + $chr_lower_ue + 'r Mitteilungen. Die Hierarchie hilft, interne Informationen thematisch zu ordnen und f' + $chr_lower_ue + 'r Anwender schneller auffindbar zu machen.' + $fieldHint
        }
        'BulletinEdit$' {
            return 'Diese Ansicht pflegt interne Mitteilungen mit Kategorie, Betreff und Inhalt. Sie dient dazu, Anwender im ERP gezielt mit wichtigen Informationen zu versorgen.' + $fieldHint
        }
        'XXTableEdit$' {
            return 'Diese Ansicht wird f' + $chr_lower_ue + 'r technische Test- und Beispieldaten verwendet. Sie ist vor allem f' + $chr_lower_ue + 'r Entwicklung, Pr' + $chr_lower_ue + 'fung und Dokumentation relevant, nicht f' + $chr_lower_ue + 'r den normalen Tagesprozess.'
        }
        'FormattedSearchEdit$' {
            return 'Diese Ansicht verwaltet formatierte Suchen, mit denen X-ERP Werte per SQL-Auswahl ermittelt und in Zielseiten oder Tabellen weiterverwendet. Dadurch lassen sich Such- und Vorbelegungslogiken zentral pflegen.' + $fieldHint
        }
        'StickyNoteEdit$' {
            return 'Diese Ansicht verwaltet kurze Notizen und Erledigungsvermerke. Sie hilft Anwendern, Hinweise direkt im Arbeitskontext festzuhalten und abgeschlossene Punkte zu kennzeichnen.' + $fieldHint
        }
        'IntercomRecipientStatusEdit$' {
            return 'Diese Ansicht verwaltet Empf' + $chr_lower_ae + 'ngerstatus f' + $chr_lower_ue + 'r interne Nachrichten. Dadurch ist erkennbar, ob eine Nachricht angekommen, gelesen, offen oder abgeschlossen ist.' + $fieldHint
        }
        'IntercomSenderStatusEdit$' {
            return 'Diese Ansicht verwaltet Absenderstatus f' + $chr_lower_ue + 'r interne Nachrichten. Sie hilft, versendete Meldungen aus Sicht des Absenders nachvollziehbar zu steuern.' + $fieldHint
        }
        'IntercomGroupEdit$' {
            return 'Diese Ansicht verwaltet Intercom-Gruppen. Gruppen strukturieren Empf' + $chr_lower_ae + 'nger und Themenbereiche, damit interne Nachrichten zielgerichtet verteilt werden k' + $chr_lower_oe + 'nnen.' + $fieldHint
        }
        'IntercomPrioEdit$' {
            return 'Diese Ansicht verwaltet Priorit' + $chr_lower_ae + 'ten f' + $chr_lower_ue + 'r interne Nachrichten. Sie sorgt daf' + $chr_lower_ue + 'r, dass dringende Meldungen im Arbeitsablauf besser erkannt und behandelt werden.' + $fieldHint
        }
        'IntercomEdit$' {
            return 'Diese Ansicht verwaltet interne Nachrichten zwischen Anwendern. Empf' + $chr_lower_ae + 'nger, Status, Stichtag und Privatkennzeichen helfen, Aufgaben und Informationen im ERP nachvollziehbar zu verteilen.' + $fieldHint
        }
        'AiTextInfoEdit$' {
            return 'Diese Ansicht verwaltet KI-Textinformationen mit Anweisung und Kontext. Sie dient dazu, wiederverwendbare Vorgaben f' + $chr_lower_ue + 'r automatische Textvorschl' + $chr_lower_ae + 'ge und redaktionelle Inhalte bereitzustellen.' + $fieldHint
        }
        'PackingListBoxPositionEdit$' {
            return 'Diese Ansicht verwaltet Positionen innerhalb von Packst' + $chr_lower_ue + 'cken. Sie verbindet Kommissionierliste, Boxtyp und Datum mit den verpackten Waren und macht den Packprozess nachvollziehbar.' + $fieldHint
        }
        'PackingListBoxTypeEdit$' {
            return 'Diese Ansicht verwaltet Packst' + $chr_lower_ue + 'cktypen mit Abmessungen. Dadurch kann X-ERP Kartons, Boxen oder andere Verpackungen f' + $chr_lower_ue + 'r Kommissionierung und Versand einheitlich verwenden.' + $fieldHint
        }
        'PackingListBoxEdit$' {
            return 'Diese Ansicht verwaltet Packst' + $chr_lower_ue + 'cke zu Kommissionierlisten. Sie ordnet Verpackung, Kunde und Matchcode dem konkreten Versand- oder Packvorgang zu.' + $fieldHint
        }
        'ProjectEdit$' {
            return 'Diese Ansicht verwaltet Projekte mit Kunde, Projektdatum und Projektname. Sie b' + $chr_lower_ue + 'ndelt projektbezogene Informationen, damit Aufgaben, Belege und Auswertungen einem Vorhaben zugeordnet werden k' + $chr_lower_oe + 'nnen.' + $fieldHint
        }
        'TextBlockGroupEdit$' {
            return 'Diese Ansicht verwaltet Gruppen f' + $chr_lower_ue + 'r Textbausteine. Die Hierarchie erleichtert das Ordnen, Finden und Wiederverwenden standardisierter Texte in Belegen und Kommunikation.' + $fieldHint
        }
        'TextBlockEdit$' {
            return 'Diese Ansicht verwaltet Textbausteine f' + $chr_lower_ue + 'r Belege, E-Mails und wiederkehrende Formulierungen. Standardtexte werden zentral gepflegt und k' + $chr_lower_oe + 'nnen dadurch konsistent im ERP verwendet werden.' + $fieldHint
        }
        'TextVariableEdit$' {
            return 'Diese Ansicht verwaltet Textvariablen, die dynamische Werte in Vorlagen und Textbausteinen ersetzen. Dadurch k' + $chr_lower_oe + 'nnen Belege und Nachrichten automatisch mit passenden Daten gef' + $chr_lower_ue + 'llt werden.' + $fieldHint
        }
        'WikiCategoryEdit$' {
            return 'Diese Ansicht verwaltet Wiki-Kategorien. Kategorien strukturieren interne Wissensartikel und helfen Anwendern, Prozesswissen und Hinweise schneller zu finden.' + $fieldHint
        }
        'WikiStatusEdit$' {
            return 'Diese Ansicht verwaltet Statuswerte f' + $chr_lower_ue + 'r Wiki-Beitr' + $chr_lower_ae + 'ge. Damit l' + $chr_lower_ae + 'sst sich erkennen, ob ein Beitrag Entwurf, freigegeben, veraltet oder in Bearbeitung ist.' + $fieldHint
        }
        'WikiEdit$' {
            return 'Diese Ansicht verwaltet Wiki-Beitr' + $chr_lower_ae + 'ge mit Kategorie, Betreff, Status und Autor. Sie macht internes Wissen direkt im ERP dokumentierbar und auffindbar.' + $fieldHint
        }
        'TimeTrackingEdit$' {
            return 'Diese Ansicht erfasst Arbeitszeiten mit Mitarbeiter-, Beleg- und Artikelbezug. Sie verbindet geleistete Zeit mit Projekten, Auftr' + $chr_lower_ae + 'gen oder Leistungen und unterst' + $chr_lower_ue + 'tzt Auswertung sowie Abrechnung.' + $fieldHint
        }
        'CompanySetting|ProgramSetting|SupportSetting|TimeTrackingSetting' {
            return 'Diese Ansicht b' + $chr_lower_ue + 'ndelt Einstellungen, die das Verhalten von X-ERP im Bereich ' + $module + ' steuern. Sie ist wichtig, weil hier Vorgaben gepflegt werden, die sich auf viele nachgelagerte Prozesse auswirken.' + $fieldHint
        }
        'Route|SystemView|SystemTableField|SystemMapping|DatabaseChangeLog' {
            return 'Diese Ansicht dokumentiert und pflegt technische Systeminformationen, die f' + $chr_lower_ue + 'r Navigation, Datenmodell und Wartung von X-ERP relevant sind.' + $fieldHint
        }
        'Article' {
            return 'Diese Ansicht pflegt ' + $label + ' und sorgt daf' + $chr_lower_ue + 'r, dass Verkauf, Einkauf, Lager, Produktion und Auswertung auf denselben Artikeldaten arbeiten.' + $fieldHint
        }
        'Partner' {
            return 'Diese Ansicht pflegt ' + $label + ', damit Adressen, Kontakte, Konditionen, Belege und Auswertungen in X-ERP korrekt miteinander verkn' + $chr_lower_ue + 'pft sind.' + $fieldHint
        }
        'Finance|Financial' {
            return 'Diese Ansicht pflegt ' + $label + ' im Finanzwesen. Sie unterst' + $chr_lower_ue + 'tzt korrekte Buchungen, Abstimmungen, Auswertungen und nachvollziehbare ERP-Finanzprozesse.' + $fieldHint
        }
        'Calendar' {
            return 'Diese Ansicht pflegt ' + $label + ' f' + $chr_lower_ue + 'r die Termin- und Arbeitszeitplanung. Dadurch bleiben Kalender, Verf' + $chr_lower_ue + 'gbarkeiten und betriebliche Zeitr' + $chr_lower_ae + 'ume einheitlich nutzbar.' + $fieldHint
        }
        'Email' {
            return 'Diese Ansicht pflegt ' + $label + ' f' + $chr_lower_ue + 'r die E-Mail-Kommunikation in X-ERP. Sie hilft, Versand, Vorlagen, Ablage und automatische Regeln kontrolliert zu steuern.' + $fieldHint
        }
        'Employee' {
            return 'Diese Ansicht pflegt ' + $label + ' und unterst' + $chr_lower_ue + 'tzt Personalstammdaten, Rollen, Abwesenheiten, Provisionen und interne Zuordnungen.' + $fieldHint
        }
        'Shipping' {
            return 'Diese Ansicht pflegt ' + $label + ' f' + $chr_lower_ue + 'r den Versand. Sie unterst' + $chr_lower_ue + 'tzt nachvollziehbare Versandarten, Zonen, Dienstleister und Kostenlogik.' + $fieldHint
        }
        'Production|Resource' {
            return 'Diese Ansicht pflegt ' + $label + ' f' + $chr_lower_ue + 'r die Produktion. Sie verbindet Ressourcen, Arbeitsschritte und Stammdaten zu planbaren Fertigungsabl' + $chr_lower_ae + 'ufen.' + $fieldHint
        }
        'Crm' {
            return 'Diese Ansicht pflegt ' + $label + ' im CRM. Sie unterst' + $chr_lower_ue + 'tzt strukturierte Kundenkontakte, Aktivit' + $chr_lower_ae + 'ten, Statuswerte und vertriebliche Nachverfolgung.' + $fieldHint
        }
        'CountryPackage' {
            return 'Diese Ansicht pflegt landesspezifische Vorgaben f' + $chr_lower_ue + 'r ' + $label + '. Sie ist wichtig f' + $chr_lower_ue + 'r gesetzliche Meldungen, Formate, Datenschutz und regionale ERP-Prozesse.' + $fieldHint
        }
        'Workspace' {
            return 'Diese Ansicht pflegt ' + $label + '. Damit wird gesteuert, welche Funktionen Benutzerrollen im X-ERP-Arbeitsplatz sehen und verwenden k' + $chr_lower_oe + 'nnen.' + $fieldHint
        }
        'Extra' {
            return 'Diese Ansicht pflegt ' + $label + ', mit denen X-ERP ohne Standardanpassung erweitert werden kann. Dadurch lassen sich zus' + $chr_lower_ae + 'tzliche Daten, Tabellen und Funktionen kontrolliert in Prozesse einbinden.' + $fieldHint
        }
        'Wizard$' {
            return 'Dieser Assistent f' + $chr_lower_ue + 'hrt durch den Prozess ' + (Split-TechnicalName $Topic) + '. Er b' + $chr_lower_ue + 'ndelt die notwendigen Eingaben und Aktionen, damit Anwender den Vorgang Schritt f' + $chr_lower_ue + 'r Schritt korrekt ausf' + $chr_lower_ue + 'hren.' + $fieldHint
        }
        default {
            return 'Diese Ansicht verwaltet ' + $label + ' im Bereich ' + $module + '. Sie stellt die fachlichen Stammdaten, Zuordnungen und Regeln bereit, die nachgelagerte ERP-Prozesse konsistent verwenden.' + $fieldHint
        }
    }
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
New-Item -ItemType Directory -Force -Path $BackupRoot | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $BackupRoot "X-ERP-HELP-before-purpose-texts-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) { throw "Worksheet not found or empty: $WorksheetName" }

    $changed = 0
    $examples = New-Object System.Collections.Generic.List[object]
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $contentType = [string]$ws.Cells.Item($r, 16).Text
        $topic = [string]$ws.Cells.Item($r, 1).Text
        if ([string]::IsNullOrWhiteSpace($topic) -or $topic -notmatch '(Edit|Wizard)$') { continue }

        $old = [string]$ws.Cells.Item($r, 8).Text
        if ($contentType -notin @('View', 'EditView') -and $old -notmatch 'Bearbeitungsansicht|Die Hilfe beschreibt|unterst.*Anwender') { continue }

        $fields = New-Object System.Collections.Generic.List[string]
        for ($i = $r + 1; $i -le $ws.Dimension.End.Row; $i++) {
            if ([string]$ws.Cells.Item($i, 16).Text -eq 'View') { break }
            $fieldTitle = [string]$ws.Cells.Item($i, 1).Text
            $fieldName = [string]$ws.Cells.Item($i, 6).Text
            if (-not [string]::IsNullOrWhiteSpace($fieldName) -and -not [string]::IsNullOrWhiteSpace($fieldTitle)) {
                $fields.Add($fieldTitle)
            }
        }

        $folder = [string]$ws.Cells.Item($r, 3).Text
        $new = Get-PurposeText -Topic $topic -Folder $folder -Fields $fields.ToArray()
        if ($new -ne $old) {
            $ws.Cells.Item($r, 8).Value = $new
            $changed++
            if ($examples.Count -lt 20) {
                $examples.Add([pscustomobject]@{ Row=$r; Topic=$topic; Before=$old; After=$new })
            }
        }
    }

    Close-ExcelPackage $pkg
    $pkg = $null

    [pscustomobject]@{
        Workbook = $WorkbookPath
        Backup = $backup
        ChangedViewDescriptions = $changed
        Examples = $examples
    } | ConvertTo-Json -Depth 5
}
finally {
    if ($null -ne $pkg) { $pkg.Dispose() }
}

