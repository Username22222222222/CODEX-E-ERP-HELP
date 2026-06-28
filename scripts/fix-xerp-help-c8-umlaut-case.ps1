param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

function Fix-Text([string]$text) {
    $chars = $text.ToCharArray()
    for ($i = 1; $i -lt $chars.Length; $i++) {
        $prevCode = [int][char]$chars[$i - 1]
        $curCode = [int][char]$chars[$i]
        if (($prevCode -ge 97 -and $prevCode -le 122) -and $curCode -eq 214) { $chars[$i] = [char]246 }
        if (($prevCode -ge 97 -and $prevCode -le 122) -and $curCode -eq 220) { $chars[$i] = [char]252 }
        if (($prevCode -ge 97 -and $prevCode -le 122) -and $curCode -eq 196) { $chars[$i] = [char]228 }
    }
    return -join $chars
}

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-c8-umlaut-case-fix-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    $changed = 0
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $cell = $ws.Cells.Item($r, 8)
        $text = [string]$cell.Text
        if ([string]::IsNullOrEmpty($text)) { continue }
        $new = Fix-Text $text
        if ($new -cne $text) {
            $cell.Value = $new
            $changed++
        }
    }
    Close-ExcelPackage $pkg
    [pscustomobject]@{
        Workbook = $WorkbookPath
        Backup = $backup
        ChangedC8Cells = $changed
    } | ConvertTo-Json
}
catch {
    $pkg.Dispose()
    throw
}
