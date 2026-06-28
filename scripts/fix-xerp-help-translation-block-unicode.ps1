param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$Ue = [char]0x00DC
$ue = [char]0x00FC
$ae = [char]0x00E4
$oe = [char]0x00F6

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-translation-unicode-fix-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$textArea = "Dieser $($Ue)bersetzungsbereich verwaltet fremdsprachige Artikeltexte und Bezeichnungen. So bleiben Artikelinformationen in mehrsprachigen Belegen, Ausgaben und Oberfl$($ae)chen konsistent."
$textLanguage = "Die Fremdsprache legt fest, f$($ue)r welche Sprache die $($Ue)bersetzung gilt. Dadurch kann X-ERP Artikeltexte, Namen oder Mengeneinheiten sprachabh$($ae)ngig ausgeben."
$textUom = "ME-Fremdsprache enth$($ae)lt die $($ue)bersetzte Bezeichnung der Mengeneinheit. Sie wird verwendet, wenn Artikelinformationen in der gew$($ae)hlten Fremdsprache ausgegeben werden."

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $changed = 0
    for ($r = 1468; $r -le 1487; $r++) {
        $topic = [string]$ws.Cells.Item($r, 1).Text
        $new = $null
        if ($topic -match '^TranslationArticle') { $new = $textArea }
        elseif ($topic -eq 'Fremdsprache') { $new = $textLanguage }
        elseif ($topic -eq 'ME-Fremdsprache') { $new = $textUom }
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
