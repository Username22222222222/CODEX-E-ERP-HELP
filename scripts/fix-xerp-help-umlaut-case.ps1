param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$lowerOe = [char]0x00F6
$upperOe = [char]0x00D6
$lowerUe = [char]0x00FC
$upperUe = [char]0x00DC
$lowerAe = [char]0x00E4
$upperAe = [char]0x00C4

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-umlaut-case-fix-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $changed = 0
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
            $cell = $ws.Cells.Item($r, $c)
            $text = [string]$cell.Text
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            $new = $text.
                Replace("geh$($upperOe)rt", "geh$($lowerOe)rt").
                Replace("zugeh$($upperOe)r", "zugeh$($lowerOe)r").
                Replace(" $($upperUe)ber", " $($lowerUe)ber")
            if ($new -ne $text) {
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
