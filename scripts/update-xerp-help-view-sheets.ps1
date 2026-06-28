param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$DataDir = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo",
    [switch]$AllowMainWorkbook
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$mainWorkbook = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx"
if (([System.IO.Path]::GetFullPath($WorkbookPath) -ieq [System.IO.Path]::GetFullPath($mainWorkbook)) -and -not $AllowMainWorkbook) {
    throw "Schutz aktiv: Dieses Skript schreibt nicht mehr automatisch in die Hauptdatei. Nutze -AllowMainWorkbook nur nach manueller Excel-Pruefung einer Kopie."
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}
if (-not (Test-Path -LiteralPath $DataDir)) {
    throw "Data directory not found: $DataDir"
}

$backupDir = Join-Path (Split-Path -Parent $WorkbookPath) 'backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $backupDir "X-ERP-HELP.before-view-sheets-$stamp.xlsx"
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

function Read-Table($name) {
    $path = Join-Path $DataDir "$name.csv"
    if (-not (Test-Path -LiteralPath $path)) { throw "Missing data file: $path" }
    return @(Import-Csv -LiteralPath $path -Delimiter ';' -Encoding UTF8)
}

function Remove-SheetIfExists($package, [string]$name) {
    $ws = $package.Workbook.Worksheets[$name]
    if ($null -ne $ws) {
        $package.Workbook.Worksheets.Delete($ws)
    }
}

function Write-Sheet($package, [string]$sheetName, [object[]]$rows, [hashtable]$options = @{}) {
    Remove-SheetIfExists $package $sheetName
    $ws = $package.Workbook.Worksheets.Add($sheetName)
    if ($rows.Count -eq 0) { return $ws }

    $headers = @($rows[0].PSObject.Properties.Name)
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $cell = $ws.Cells.Item(1, $c + 1)
        $cell.Value = $headers[$c]
        $cell.Style.Font.Bold = $true
        $cell.Style.Font.Color.SetColor([System.Drawing.Color]::White)
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31, 78, 121))
    }

    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $ws.Cells.Item($r + 2, $c + 1).Value = [string]$rows[$r].($headers[$c])
        }
    }

    $lastRow = $rows.Count + 1
    $lastCol = $headers.Count
    $ws.Cells.Item(1, 1, $lastRow, $lastCol).AutoFilter = $true
    $ws.View.FreezePanes(2, 1)

    for ($c = 1; $c -le $lastCol; $c++) {
        $ws.Column($c).AutoFit()
        if ($ws.Column($c).Width -gt 60) { $ws.Column($c).Width = 60 }
        if ($ws.Column($c).Width -lt 10) { $ws.Column($c).Width = 10 }
    }

    if ($options.ContainsKey('OutlineByLevel') -and $options.OutlineByLevel) {
        $levelCol = [Array]::IndexOf($headers, 'LEVEL') + 1
        $nameCol = [Array]::IndexOf($headers, 'DE_NAME') + 1
        $typeCol = [Array]::IndexOf($headers, 'TYPE') + 1
        $ws.OutLineSummaryBelow = $false
        for ($r = 2; $r -le $lastRow; $r++) {
            $level = 0
            if ($levelCol -gt 0) { [void][int]::TryParse([string]($ws.Cells.Item($r, $levelCol).Value), [ref]$level) }
            if ($level -gt 0) { $ws.Row($r).OutlineLevel = [Math]::Min($level, 7) }
            if ($level -gt 1) { $ws.Row($r).Hidden = $true }
            if ($nameCol -gt 0) { $ws.Cells.Item($r, $nameCol).Style.Indent = [Math]::Min($level * 2, 12) }

            $type = if ($typeCol -gt 0) { [string]$ws.Cells.Item($r, $typeCol).Value } else { '' }
            if ($type -in @('EditView', 'Wizard')) {
                $ws.Row($r).Style.Font.Bold = $true
                $ws.Row($r).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Row($r).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(221, 235, 247))
            }
            elseif ($type -eq 'REGISTER_TAB') {
                $ws.Row($r).Style.Font.Bold = $true
                $ws.Row($r).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Row($r).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(234, 241, 221))
            }
        }
    }

    return $ws
}

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $inventory = Read-Table 'views-inventory'
    $structure = Read-Table 'views-structure'
    $drafts = Read-Table 'views-text-drafts'
    $audit = Read-Table 'views-audit'

    Write-Sheet $pkg 'Views-Inventar' $inventory | Out-Null
    Write-Sheet $pkg 'Views-Struktur' $structure @{ OutlineByLevel = $true } | Out-Null
    Write-Sheet $pkg 'Views-Textvorlagen' $drafts | Out-Null
    Write-Sheet $pkg 'Views-Audit' $audit | Out-Null

    Close-ExcelPackage $pkg
}
catch {
    $pkg.Dispose()
    throw
}

$check = Open-ExcelPackage -Path $WorkbookPath
try {
    $de = $check.Workbook.Worksheets['de-DE']
    $outlineCounts = @{}
    if ($null -ne $de -and $null -ne $de.Dimension) {
        for ($r = 1; $r -le $de.Dimension.End.Row; $r++) {
            $level = [int]$de.Row($r).OutlineLevel
            if (-not $outlineCounts.ContainsKey($level)) { $outlineCounts[$level] = 0 }
            $outlineCounts[$level]++
        }
    }
    $result = [pscustomobject]@{
        Workbook = $WorkbookPath
        Backup = $backup
        InventoryRows = (Read-Table 'views-inventory').Count
        StructureRows = (Read-Table 'views-structure').Count
        TextDraftRows = (Read-Table 'views-text-drafts').Count
        DeDEOutline = (($outlineCounts.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name):$($_.Value)" }) -join ', ')
    }
    $result | ConvertTo-Json -Depth 4
}
finally {
    $check.Dispose()
}
