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
$backup = Join-Path $backupDir "X-ERP-HELP.before-translation-block-fix-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $changed = 0
    for ($r = 1468; $r -le 1487; $r++) {
        $topic = [string]$ws.Cells.Item($r, 1).Text
        $field = [string]$ws.Cells.Item($r, 6).Text
        $new = $null
        if ($topic -match '^TranslationArticle') {
            $new = "Dieser Übersetzungsbereich verwaltet fremdsprachige Artikeltexte und Bezeichnungen. So bleiben Artikelinformationen in mehrsprachigen Belegen, Ausgaben und Oberflächen konsistent."
        }
        elseif ($topic -eq 'Fremdsprache') {
            $new = "Die Fremdsprache legt fest, für welche Sprache die Übersetzung gilt. Dadurch kann X-ERP Artikeltexte, Namen oder Mengeneinheiten sprachabhängig ausgeben."
        }
        elseif ($topic -eq 'ME-Fremdsprache') {
            $new = "ME-Fremdsprache enthält die übersetzte Bezeichnung der Mengeneinheit. Sie wird verwendet, wenn Artikelinformationen in der gewählten Fremdsprache ausgegeben werden."
        }
        if ($new) {
            $ws.Cells.Item($r, 8).Value = $new
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
