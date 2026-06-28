param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-spaced-technical-name-fix-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $changed = 0
    $currentModule = ''
    $currentViewTopic = ''
    $currentTab = ''

    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $level = [int]$ws.Row($r).OutlineLevel
        $topic = [string]$ws.Cells.Item($r, 1).Text
        $field = [string]$ws.Cells.Item($r, 6).Text
        $cell = $ws.Cells.Item($r, 8)
        $text = [string]$cell.Text
        if ($level -eq 1 -and -not [string]::IsNullOrWhiteSpace($topic)) { $currentModule = $topic }
        if ($level -eq 2 -and -not [string]::IsNullOrWhiteSpace($topic)) { $currentViewTopic = $topic; $currentTab = '' }
        if ($level -eq 3 -and [string]::IsNullOrWhiteSpace($field) -and -not [string]::IsNullOrWhiteSpace($topic)) { $currentTab = $topic }
        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        $new = $text
        if ($text -match '\b[A-ZÄÖÜ][a-zäöüß]{0,2}(?:\s+[a-zA-ZÄÖÜäöüß]{1,3}){2,}') {
            if ($level -eq 2) {
                $new = "Diese Bearbeitungsansicht unterstützt Anwender beim Anlegen, Prüfen und Bearbeiten der zugehörigen Datensätze im Bereich $currentModule. Die Hilfe beschreibt die Register, Felder und Aktionen aus fachlicher Sicht."
            }
            elseif (-not [string]::IsNullOrWhiteSpace($field)) {
                $area = if ([string]::IsNullOrWhiteSpace($currentTab)) { "dieser Ansicht" } else { "dem Register $currentTab" }
                $new = "$topic ist eine fachliche Angabe in $area. Sie unterstützt Anwender dabei, den Datensatz korrekt zu pflegen, zu finden und im ERP-Prozess weiterzuverarbeiten."
            }
            elseif ($level -eq 3) {
                $new = "Dieser Bereich bündelt die fachlich passenden Einstellungen und Informationen. Anwender finden hier die Felder, die gemeinsam bearbeitet oder geprüft werden."
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
