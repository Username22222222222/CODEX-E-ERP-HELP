param(
    [string]$SourcePath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$OutputPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.sorted.xlsx",
    [string]$RegisterDataPath = "C:\X-ERP\X-ERP\X-ERP.Shared\Constants\SystemRegisterData.cs",
    [string]$ResxPath = "C:\X-ERP\X-ERP\X-ERP.Client\ResourceFiles\LzrResource.de-DE.resx"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

function Convert-ToKey([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    return ($text.Trim().ToLowerInvariant() -replace '\s+', ' ')
}

function Load-Translations([string]$path) {
    [xml]$xml = Get-Content -LiteralPath $path -Raw
    $map = @{}
    foreach ($d in $xml.root.data) {
        if ($d.name) { $map[$d.name] = [string]$d.value }
    }
    return $map
}

function Load-Registers([string]$path, [hashtable]$translations) {
    $text = Get-Content -LiteralPath $path -Raw
    $rx = [regex]'new\("([^"]+)",\s*(\d+),\s*"([^"]+)",\s*"([^"]+)"\)'
    $byType = @{}
    foreach ($m in $rx.Matches($text)) {
        $english = $m.Groups[3].Value
        $de = if ($translations.ContainsKey($english)) { $translations[$english] } else { $english }
        $entry = [pscustomobject]@{
            Id = $m.Groups[1].Value
            Order = [int]$m.Groups[2].Value
            English = $english
            German = $de
            Type = $m.Groups[4].Value
        }
        $key = Convert-ToKey $entry.Type
        if (-not $byType.ContainsKey($key)) { $byType[$key] = New-Object System.Collections.Generic.List[object] }
        $byType[$key].Add($entry)
    }
    foreach ($key in @($byType.Keys)) {
        $byType[$key] = @($byType[$key] | Sort-Object Order)
    }
    return $byType
}

function Get-RegisterTypeForView([string]$viewName) {
    switch ($viewName) {
        'ArticleEdit' { return 'Article' }
        'PartnerEdit' { return 'Partner' }
        'CrmEdit' { return 'Crm' }
        'ResourceEdit' { return 'Resource' }
        default {
            if ($viewName.EndsWith('Edit')) { return $viewName.Substring(0, $viewName.Length - 4) }
            if ($viewName.EndsWith('Wizard')) { return $viewName }
            return $viewName
        }
    }
}

function New-RowObject($ws, [int]$r, [int]$lastCol) {
    $values = New-Object string[] ($lastCol + 1)
    for ($c = 1; $c -le $lastCol; $c++) {
        $v = $ws.Cells.Item($r, $c).Value
        if ($null -ne $v) { $values[$c] = [string]$v }
    }
    return [pscustomobject]@{
        Values = $values
        OutlineLevel = [int]$ws.Row($r).OutlineLevel
        Hidden = [bool]$ws.Row($r).Hidden
        Collapsed = [bool]$ws.Row($r).Collapsed
        Height = [double]$ws.Row($r).Height
        SourceRow = $r
    }
}

function Get-Cell([object]$row, [int]$col) {
    if ($null -eq $row -or $null -eq $row.Values -or $row.Values.Length -le $col) { return '' }
    return [string]$row.Values[$col]
}

function Build-RegisterLookup([array]$registers) {
    $lookup = @{}
    for ($i = 0; $i -lt $registers.Count; $i++) {
        $reg = $registers[$i]
        foreach ($label in @($reg.English, $reg.German)) {
            $key = Convert-ToKey $label
            if ($key -and -not $lookup.ContainsKey($key)) {
                $lookup[$key] = [pscustomobject]@{ Index = $i; Register = $reg }
            }
        }
    }
    return $lookup
}

function Try-MatchRegister([object]$row, [hashtable]$lookup) {
    foreach ($col in @(2, 1)) {
        $key = Convert-ToKey (Get-Cell $row $col)
        if ($lookup.ContainsKey($key)) { return $lookup[$key] }
    }
    return $null
}

function Sort-ViewChildren([array]$children, [array]$registers) {
    if ($registers.Count -eq 0 -or $children.Count -eq 0) { return $children }
    $lookup = Build-RegisterLookup $registers
    $prefix = New-Object System.Collections.Generic.List[object]
    $blocks = @()
    $suffix = New-Object System.Collections.Generic.List[object]

    $i = 0
    $foundFirstRegister = $false
    while ($i -lt $children.Count) {
        $row = $children[$i]
        $match = if ($row.OutlineLevel -eq 3) { Try-MatchRegister $row $lookup } else { $null }
        if ($null -ne $match) {
            $foundFirstRegister = $true
            $blockRows = New-Object System.Collections.Generic.List[object]
            $blockRows.Add($row)
            $i++
            while ($i -lt $children.Count -and $children[$i].OutlineLevel -gt 3) {
                $blockRows.Add($children[$i])
                $i++
            }
            $blocks += [pscustomobject]@{
                SortIndex = [int]$match.Index
                OriginalFirstRow = [int]$row.SourceRow
                Rows = $blockRows.ToArray()
            }
            continue
        }

        if (-not $foundFirstRegister) { $prefix.Add($row) }
        else { $suffix.Add($row) }
        $i++
    }

    if ($blocks.Count -eq 0) { return $children }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($row in $prefix) { $result.Add($row) }
    foreach ($block in ($blocks | Sort-Object SortIndex, OriginalFirstRow)) {
        foreach ($row in $block.Rows) { $result.Add($row) }
    }
    foreach ($row in $suffix) { $result.Add($row) }
    return $result.ToArray()
}

function Sort-DeRows([array]$rows, [hashtable]$registersByType) {
    $result = New-Object System.Collections.Generic.List[object]
    $i = 0
    while ($i -lt $rows.Count) {
        $row = $rows[$i]
        $topic = Get-Cell $row 1
        if ($row.OutlineLevel -eq 2 -and ($topic.EndsWith('Edit') -or $topic.EndsWith('Wizard'))) {
            $result.Add($row)
            $children = New-Object System.Collections.Generic.List[object]
            $i++
            while ($i -lt $rows.Count -and $rows[$i].OutlineLevel -gt 2) {
                $children.Add($rows[$i])
                $i++
            }
            $type = Get-RegisterTypeForView $topic
            $key = Convert-ToKey $type
            $registers = if ($registersByType.ContainsKey($key)) { @($registersByType[$key]) } else { @() }
            foreach ($child in (Sort-ViewChildren $children.ToArray() $registers)) { $result.Add($child) }
            continue
        }
        $result.Add($row)
        $i++
    }
    return $result.ToArray()
}

function Apply-LevelFormatting($ws, [int]$lastRow, [int]$lastCol) {
    $colors = @{
        0 = [System.Drawing.Color]::FromArgb(31, 78, 121)
        1 = [System.Drawing.Color]::FromArgb(46, 108, 166)
        2 = [System.Drawing.Color]::FromArgb(74, 137, 199)
        3 = [System.Drawing.Color]::FromArgb(125, 169, 216)
        4 = [System.Drawing.Color]::FromArgb(169, 199, 232)
        5 = [System.Drawing.Color]::FromArgb(214, 228, 244)
    }
    $fontSizes = @{ 0 = 18; 1 = 16; 2 = 14; 3 = 12; 4 = 10; 5 = 9 }

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
        $range = $ws.Cells.Item($r, 1, $r, $lastCol)
        $range.Style.Font.Size = $fontSizes[$level]
        $range.Style.Font.Bold = ($level -le 2)
        $range.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(31, 31, 31))
        $range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $range.Style.Fill.BackgroundColor.SetColor($colors[$level])
        $range.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
        $range.Style.Border.Bottom.Color.SetColor([System.Drawing.Color]::FromArgb(217, 225, 242))
        $ws.Cells.Item($r, 1).Style.Indent = [Math]::Min($level * 2, 12)
        if ($lastCol -ge 2) { $ws.Cells.Item($r, 2).Style.Indent = [Math]::Min($level * 2, 12) }
    }
}

$translations = Load-Translations $ResxPath
$registersByType = Load-Registers $RegisterDataPath $translations

$source = Open-ExcelPackage -Path $SourcePath
$target = New-Object OfficeOpenXml.ExcelPackage
try {
    foreach ($srcWs in $source.Workbook.Worksheets) {
        if (-not $srcWs.Dimension) { continue }
        $lastRow = $srcWs.Dimension.End.Row
        $lastCol = $srcWs.Dimension.End.Column
        $rows = New-Object System.Collections.Generic.List[object]
        for ($r = 1; $r -le $lastRow; $r++) {
            $rows.Add((New-RowObject $srcWs $r $lastCol))
        }
        $outRows = if ($srcWs.Name -eq 'de-DE') { Sort-DeRows $rows.ToArray() $registersByType } else { $rows.ToArray() }

        $dstWs = $target.Workbook.Worksheets.Add($srcWs.Name)
        for ($r = 1; $r -le $outRows.Count; $r++) {
            $row = $outRows[$r - 1]
            $dstRow = $dstWs.Row($r)
            $dstRow.OutlineLevel = [Math]::Min([int]$row.OutlineLevel, 7)
            $dstRow.Hidden = [bool]$row.Hidden
            $dstRow.Collapsed = [bool]$row.Collapsed
            if ($row.Height -gt 0) { $dstRow.Height = $row.Height }
            for ($c = 1; $c -le $lastCol; $c++) {
                if (-not [string]::IsNullOrEmpty($row.Values[$c])) {
                    $dstWs.Cells.Item($r, $c).Value = $row.Values[$c]
                }
            }
        }
        for ($c = 1; $c -le $lastCol; $c++) {
            $srcCol = $srcWs.Column($c)
            $dstCol = $dstWs.Column($c)
            if ($srcCol.Width -gt 0) { $dstCol.Width = [Math]::Min([Math]::Max($srcCol.Width, 8), 60) }
            $dstCol.Hidden = $srcCol.Hidden
        }
        $dstWs.OutLineSummaryBelow = $false
        $dstWs.OutLineSummaryRight = $false
        $dstWs.View.FreezePanes(2, 1)
        if ($outRows.Count -gt 1 -and $lastCol -gt 0) {
            $dstWs.Cells.Item(1, 1, $outRows.Count, $lastCol).AutoFilter = $true
        }
        if ($srcWs.Name -eq 'de-DE') {
            Apply-LevelFormatting $dstWs $outRows.Count $lastCol
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
