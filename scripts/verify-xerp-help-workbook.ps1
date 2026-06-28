param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $sheets = @()
    foreach ($ws in $pkg.Workbook.Worksheets) {
        $sheets += [pscustomobject]@{
            Name = $ws.Name
            Rows = if ($ws.Dimension) { $ws.Dimension.End.Row } else { 0 }
            Cols = if ($ws.Dimension) { $ws.Dimension.End.Column } else { 0 }
        }
    }

    $outline = @{}
    $de = $pkg.Workbook.Worksheets['de-DE']
    if ($de -and $de.Dimension) {
        for ($r = 1; $r -le $de.Dimension.End.Row; $r++) {
            $level = [int]$de.Row($r).OutlineLevel
            if (-not $outline.ContainsKey($level)) { $outline[$level] = 0 }
            $outline[$level]++
        }
    }

    $structureOutline = @{}
    $structure = $pkg.Workbook.Worksheets['Views-Struktur']
    if ($structure -and $structure.Dimension) {
        for ($r = 1; $r -le $structure.Dimension.End.Row; $r++) {
            $level = [int]$structure.Row($r).OutlineLevel
            if (-not $structureOutline.ContainsKey($level)) { $structureOutline[$level] = 0 }
            $structureOutline[$level]++
        }
    }

    $errors = New-Object System.Collections.Generic.List[string]
    foreach ($ws in $pkg.Workbook.Worksheets) {
        if (-not $ws.Dimension) { continue }
        for ($r = 1; $r -le $ws.Dimension.End.Row; $r++) {
            for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
                $v = [string]$ws.Cells.Item($r, $c).Text
                if ($v -match '#REF!|#DIV/0!|#VALUE!|#NAME\?|#N/A') {
                    $errors.Add("$($ws.Name)!$r,$c=$v")
                    if ($errors.Count -ge 20) { break }
                }
            }
            if ($errors.Count -ge 20) { break }
        }
        if ($errors.Count -ge 20) { break }
    }

    [pscustomobject]@{
        Workbook = $WorkbookPath
        Sheets = $sheets
        DeDEOutline = (($outline.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name):$($_.Value)" }) -join ', ')
        ViewsStructureOutline = (($structureOutline.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name):$($_.Value)" }) -join ', ')
        FormulaLikeErrorsFound = $errors.Count
        FirstFormulaLikeErrors = @($errors)
    } | ConvertTo-Json -Depth 6
}
finally {
    $pkg.Dispose()
}
