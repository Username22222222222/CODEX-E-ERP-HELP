param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-level-formatting-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) {
        throw "Worksheet not found or empty: $WorksheetName"
    }

    $colors = @{
        0 = [System.Drawing.Color]::FromArgb(31, 78, 121)
        1 = [System.Drawing.Color]::FromArgb(46, 108, 166)
        2 = [System.Drawing.Color]::FromArgb(74, 137, 199)
        3 = [System.Drawing.Color]::FromArgb(125, 169, 216)
        4 = [System.Drawing.Color]::FromArgb(169, 199, 232)
        5 = [System.Drawing.Color]::FromArgb(214, 228, 244)
    }
    $fontSizes = @{
        0 = 18
        1 = 16
        2 = 14
        3 = 12
        4 = 10
        5 = 9
    }

    $lastRow = $ws.Dimension.End.Row
    $lastCol = $ws.Dimension.End.Column

    $header = $ws.Cells.Item(1, 1, 1, $lastCol)
    $header.Style.Font.Bold = $true
    $header.Style.Font.Size = 11
    $header.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $header.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $header.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31, 78, 121))
    $header.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

    for ($r = 2; $r -le $lastRow; $r++) {
        $level = [int]$ws.Row($r).OutlineLevel
        if (-not $colors.ContainsKey($level)) { $level = 5 }

        $rowRange = $ws.Cells.Item($r, 1, $r, $lastCol)
        $rowRange.Style.Font.Size = $fontSizes[$level]
        $rowRange.Style.Font.Bold = ($level -le 2)
        $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(31, 31, 31))
        $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $rowRange.Style.Fill.BackgroundColor.SetColor($colors[$level])
        $rowRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
        $rowRange.Style.Border.Bottom.Color.SetColor([System.Drawing.Color]::FromArgb(217, 225, 242))

        $ws.Cells.Item($r, 1).Style.Indent = [Math]::Min($level * 2, 12)
        if ($lastCol -ge 2) { $ws.Cells.Item($r, 2).Style.Indent = [Math]::Min($level * 2, 12) }
    }

    $ws.OutLineSummaryBelow = $false
    $ws.OutLineSummaryRight = $false
    $ws.View.FreezePanes(2, 1)

    Close-ExcelPackage $pkg
}
catch {
    $pkg.Dispose()
    throw
}

[pscustomobject]@{
    Workbook = $WorkbookPath
    Backup = $backup
    Worksheet = $WorksheetName
    RowsFormatted = $lastRow - 1
} | ConvertTo-Json
