param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$ScreenshotRoot = "D:\DATEN\HOMEPAGES\x-erp.de\de\help\Screenshots",
    [string]$RelativeBase = "Screenshots",
    [string]$WebBase = "/de/help/Screenshots"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

function Get-SafeFileName([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return '' }
    $v = $value.Trim().Replace('\', '/')
    if ($v -match '^https?://') { return [IO.Path]::GetFileName(([Uri]$v).AbsolutePath) }
    if ($v -match '^[A-Za-z]:/') { return [IO.Path]::GetFileName($v) }
    if ($v.StartsWith('/de/help/Screenshots/', [StringComparison]::OrdinalIgnoreCase)) {
        return [IO.Path]::GetFileName($v)
    }
    if ($v.StartsWith('Screenshots/', [StringComparison]::OrdinalIgnoreCase)) {
        return [IO.Path]::GetFileName($v)
    }
    return [IO.Path]::GetFileName($v)
}

function Get-AltText([string]$topic, [string]$view, [string]$module) {
    $main = if (-not [string]::IsNullOrWhiteSpace($view)) { $view } else { $topic }
    if ([string]::IsNullOrWhiteSpace($main)) { $main = 'X-ERP Hilfe' }
    if (-not [string]::IsNullOrWhiteSpace($module)) {
        return "$main Ansicht im ERP-System X-ERP im Bereich $module"
    }
    return "$main Ansicht im ERP-System X-ERP"
}

function Get-Caption([string]$topic, [string]$view) {
    if (-not [string]::IsNullOrWhiteSpace($view) -and $view -ne $topic) {
        return "$topic in der Ansicht $view."
    }
    return "$topic in X-ERP."
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'ARCHIV\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-screenshot-metadata-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) { throw "Worksheet not found or empty: $WorksheetName" }

    $headers = @{}
    $lastHeaderCol = 0
    for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
        $h = [string]$ws.Cells.Item(1, $c).Text
        if (-not [string]::IsNullOrWhiteSpace($h)) {
            $headers[$h] = $c
            $lastHeaderCol = $c
        }
    }

    if (-not $headers.ContainsKey('Screenshot')) { throw "Column 'Screenshot' not found." }

    foreach ($newHeader in @('SCREENSHOT_REL_PATH', 'SCREENSHOT_WEB_PATH', 'IMAGE_ALT', 'IMAGE_CAPTION', 'IMAGE_STATUS')) {
        if (-not $headers.ContainsKey($newHeader)) {
            $lastHeaderCol++
            $ws.Cells.Item(1, $lastHeaderCol).Value = $newHeader
            $headers[$newHeader] = $lastHeaderCol
        }
    }

    $colScreenshot = $headers['Screenshot']
    $colRel = $headers['SCREENSHOT_REL_PATH']
    $colWeb = $headers['SCREENSHOT_WEB_PATH']
    $colAlt = $headers['IMAGE_ALT']
    $colCaption = $headers['IMAGE_CAPTION']
    $colStatus = $headers['IMAGE_STATUS']

    $updated = 0
    $found = 0
    $missing = 0
    $empty = 0
    $currentModule = ''
    $currentView = ''

    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $level = [int]$ws.Row($r).OutlineLevel
        $topic = [string]$ws.Cells.Item($r, 1).Text
        if ($level -eq 1 -and -not [string]::IsNullOrWhiteSpace($topic)) { $currentModule = $topic }
        if ($level -eq 2 -and -not [string]::IsNullOrWhiteSpace($topic)) { $currentView = $topic }

        $rawScreenshot = [string]$ws.Cells.Item($r, $colScreenshot).Text
        $fileName = Get-SafeFileName $rawScreenshot

        if ([string]::IsNullOrWhiteSpace($fileName)) {
            $ws.Cells.Item($r, $colStatus).Value = 'kein Screenshot'
            $empty++
            continue
        }

        $relPath = "$RelativeBase/$fileName"
        $webPath = "$WebBase/$fileName"
        $diskPath = Join-Path $ScreenshotRoot $fileName
        $status = if (Test-Path -LiteralPath $diskPath) { $found++; 'vorhanden' } else { $missing++; 'fehlt' }

        # Keep the original Screenshot column portable as recommended.
        $ws.Cells.Item($r, $colScreenshot).Value = $relPath
        $ws.Cells.Item($r, $colRel).Value = $relPath
        $ws.Cells.Item($r, $colWeb).Value = $webPath
        $ws.Cells.Item($r, $colAlt).Value = Get-AltText $topic $currentView $currentModule
        $ws.Cells.Item($r, $colCaption).Value = Get-Caption $topic $currentView
        $ws.Cells.Item($r, $colStatus).Value = $status
        $updated++
    }

    $headerRange = $ws.Cells.Item(1, 1, 1, $lastHeaderCol)
    $headerRange.Style.Font.Bold = $true
    $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31, 78, 121))
    for ($c = 1; $c -le $lastHeaderCol; $c++) {
        if ($ws.Column($c).Width -lt 10) { $ws.Column($c).Width = 10 }
        if ($ws.Column($c).Width -gt 60) { $ws.Column($c).Width = 60 }
    }

    Close-ExcelPackage $pkg

    [pscustomobject]@{
        Workbook = $WorkbookPath
        Backup = $backup
        ScreenshotRoot = $ScreenshotRoot
        RowsWithScreenshot = $updated
        ExistingFiles = $found
        MissingFiles = $missing
        RowsWithoutScreenshot = $empty
        AddedOrEnsuredColumns = 'SCREENSHOT_REL_PATH, SCREENSHOT_WEB_PATH, IMAGE_ALT, IMAGE_CAPTION, IMAGE_STATUS'
    } | ConvertTo-Json
}
catch {
    $pkg.Dispose()
    throw
}
