param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "",
    [string]$OldText = "Zusatzfelder",
    [string]$NewText = "Extra-Felder",
    [switch]$ReplaceInsideText
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-term-replace-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $changed = 0
    $sheets = if ([string]::IsNullOrWhiteSpace($WorksheetName)) {
        @($pkg.Workbook.Worksheets)
    }
    else {
        @($pkg.Workbook.Worksheets[$WorksheetName])
    }

    foreach ($ws in $sheets) {
        if ($null -eq $ws -or $null -eq $ws.Dimension) { continue }
        for ($r = 1; $r -le $ws.Dimension.End.Row; $r++) {
            for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
                $cell = $ws.Cells.Item($r, $c)
                $text = [string]$cell.Text
                if ($ReplaceInsideText) {
                    if ($text.Contains($OldText)) {
                        $cell.Value = $text.Replace($OldText, $NewText)
                        $changed++
                    }
                }
                elseif ($text -eq $OldText) {
                    $cell.Value = $NewText
                    $changed++
                }
            }
        }
    }

    Close-ExcelPackage $pkg

    [pscustomobject]@{
        Workbook = $WorkbookPath
        Backup = $backup
        OldText = $OldText
        NewText = $NewText
        ReplaceInsideText = [bool]$ReplaceInsideText
        ChangedCells = $changed
    } | ConvertTo-Json
}
catch {
    $pkg.Dispose()
    throw
}
