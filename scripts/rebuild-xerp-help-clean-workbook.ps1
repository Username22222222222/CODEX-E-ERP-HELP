param(
    [string]$SourcePath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$OutputPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.clean.xlsx"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

if (-not (Test-Path -LiteralPath $SourcePath)) {
    throw "Source workbook not found: $SourcePath"
}

$source = Open-ExcelPackage -Path $SourcePath
$target = New-Object OfficeOpenXml.ExcelPackage
try {
    foreach ($srcWs in $source.Workbook.Worksheets) {
        if (-not $srcWs.Dimension) { continue }
        $dstWs = $target.Workbook.Worksheets.Add($srcWs.Name)
        $rows = $srcWs.Dimension.End.Row
        $cols = $srcWs.Dimension.End.Column

        for ($r = 1; $r -le $rows; $r++) {
            $srcRow = $srcWs.Row($r)
            $dstRow = $dstWs.Row($r)
            $dstRow.OutlineLevel = [Math]::Min([int]$srcRow.OutlineLevel, 7)
            $dstRow.Hidden = $srcRow.Hidden
            $dstRow.Collapsed = $srcRow.Collapsed
            if ($srcRow.Height -gt 0) { $dstRow.Height = $srcRow.Height }

            for ($c = 1; $c -le $cols; $c++) {
                $srcCell = $srcWs.Cells.Item($r, $c)
                $dstCell = $dstWs.Cells.Item($r, $c)
                if ($srcCell.Value -ne $null) {
                    $dstCell.Value = $srcCell.Value
                }
            }
        }

        for ($c = 1; $c -le $cols; $c++) {
            $srcCol = $srcWs.Column($c)
            $dstCol = $dstWs.Column($c)
            if ($srcCol.Width -gt 0) { $dstCol.Width = [Math]::Min([Math]::Max($srcCol.Width, 8), 60) }
            $dstCol.Hidden = $srcCol.Hidden
        }

        $dstWs.OutLineSummaryBelow = $srcWs.OutLineSummaryBelow
        $dstWs.OutLineSummaryRight = $srcWs.OutLineSummaryRight
        $dstWs.View.FreezePanes(2, 1)

        $header = $dstWs.Cells.Item(1, 1, 1, $cols)
        $header.Style.Font.Bold = $true
        $header.Style.Font.Color.SetColor([System.Drawing.Color]::White)
        $header.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $header.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31, 78, 121))
        $header.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

        if ($rows -gt 1 -and $cols -gt 0) {
            $dstWs.Cells.Item(1, 1, $rows, $cols).AutoFilter = $true
        }
    }

    if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }
    $file = New-Object System.IO.FileInfo($OutputPath)
    $target.SaveAs($file)
}
finally {
    $source.Dispose()
    $target.Dispose()
}

[pscustomobject]@{
    Source = $SourcePath
    Output = $OutputPath
    Length = (Get-Item -LiteralPath $OutputPath).Length
} | ConvertTo-Json
