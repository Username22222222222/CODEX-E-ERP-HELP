param(
  [string]$Path = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel -ErrorAction Stop

if (-not (Test-Path -LiteralPath $Path)) {
  throw "Workbook not found: $Path"
}

$backupDir = Join-Path (Split-Path -Parent $Path) "backups"
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$backup = Join-Path $backupDir ("X-ERP-HELP.before-outline-repair-{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Copy-Item -LiteralPath $Path -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $Path
try {
  $ws = $pkg.Workbook.Worksheets['de-DE']
  if (-not $ws) { throw "Worksheet 'de-DE' not found." }

  $endRow = $ws.Dimension.End.Row
  $endCol = $ws.Dimension.End.Column

  # The current workbook still preserves the original visual level styling.
  # Rebuild real Excel outline levels from those stable style markers.
  $levelBySizeAndFill = @{
    '18|FF1F4E78' = 0
    '16|FF2E6CA6' = 1
    '14|FF4A89C7' = 2
    '12|FF7DA9D8' = 3
    '10|FFA9C7E8' = 4
    '8|FFD6E4F4'  = 5
  }

  $indentByLevel = @{
    0 = 0
    1 = 2
    2 = 4
    3 = 6
    4 = 8
    5 = 10
  }

  # Clear stale outline state first.
  for ($r = 1; $r -le $endRow; $r++) {
    $row = $ws.Row($r)
    $row.OutlineLevel = 0
    $row.Hidden = $false
    $row.Collapsed = $false
  }

  for ($r = 2; $r -le $endRow; $r++) {
    $cell = $ws.Cells[$r, 1]
    $size = [int][Math]::Round([double]$cell.Style.Font.Size)
    $fill = "$($cell.Style.Fill.BackgroundColor.Rgb)"
    $key = "$size|$fill"
    $level = if ($levelBySizeAndFill.ContainsKey($key)) { [int]$levelBySizeAndFill[$key] } else { 3 }

    $row = $ws.Row($r)
    $row.OutlineLevel = $level
    $row.Hidden = $level -gt 0
    $row.Collapsed = $false

    $cell.Style.Indent = [int]$indentByLevel[$level]
  }

  # Mark summary rows collapsed when the following row is a child.
  for ($r = 2; $r -lt $endRow; $r++) {
    $level = [int]$ws.Row($r).OutlineLevel
    $nextLevel = [int]$ws.Row($r + 1).OutlineLevel
    if ($nextLevel -gt $level) {
      $ws.Row($r).Collapsed = $true
    }
  }

  $ws.OutLineSummaryBelow = $false
  $ws.OutLineSummaryRight = $false
  $ws.View.FreezePanes(2, 1)

  Close-ExcelPackage $pkg
  [pscustomobject]@{
    Workbook = $Path
    Backup = $backup
    RowsProcessed = $endRow
    ColumnsPreserved = $endCol
  } | Format-List
}
catch {
  Close-ExcelPackage $pkg -NoSave
  throw
}
