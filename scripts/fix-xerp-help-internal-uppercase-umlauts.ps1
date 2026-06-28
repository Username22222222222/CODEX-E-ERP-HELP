param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-internal-uppercase-umlauts-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

function Fix-Text([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return $text }
    $chars = $text.ToCharArray()
    for ($i = 1; $i -lt $chars.Length; $i++) {
        $prev = [int][char]$chars[$i - 1]
        $cur = [int][char]$chars[$i]
        $prevIsLowerAscii = ($prev -ge 97 -and $prev -le 122)
        if ($prevIsLowerAscii -and $cur -eq 214) { $chars[$i] = [char]246 }
        if ($prevIsLowerAscii -and $cur -eq 220) { $chars[$i] = [char]252 }
        if ($prevIsLowerAscii -and $cur -eq 196) { $chars[$i] = [char]228 }
    }
    return -join $chars
}

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $changed = 0
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
            $cell = $ws.Cells.Item($r, $c)
            $text = [string]$cell.Text
            $new = Fix-Text $text
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
