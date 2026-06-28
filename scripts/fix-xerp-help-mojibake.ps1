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

$mojiAe = [string]([char]195) + [string]([char]164)
$mojiOe = [string]([char]195) + [string]([char]182)
$mojiUe = [string]([char]195) + [string]([char]188)
$mojiUpperAe = [string]([char]195) + [string]([char]132)
$mojiUpperOe = [string]([char]195) + [string]([char]150)
$mojiUpperUe = [string]([char]195) + [string]([char]156)
$mojiSs = [string]([char]195) + [string]([char]376)

function Fix-Mojibake([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return $text }
    return $text.
        Replace($mojiAe, [string]$lowerAe).
        Replace($mojiOe, [string]$lowerOe).
        Replace($mojiUe, [string]$lowerUe).
        Replace($mojiUpperAe, [string]$upperAe).
        Replace($mojiUpperOe, [string]$upperOe).
        Replace($mojiUpperUe, [string]$upperUe).
        Replace($mojiSs, [string]$ss)
}

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-mojibake-fix-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $changed = 0
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
            $cell = $ws.Cells.Item($r, $c)
            $text = [string]$cell.Text
            $new = Fix-Mojibake $text
            if ($new -cne $text) {
                $cell.Value = $new
                $changed++
            }
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
