param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx"
)
$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $cell = $pkg.Workbook.Worksheets['de-DE'].Cells.Item(1161, 8)
    $text = [string]$cell.Text
    $chars = $text.ToCharArray()
    $hits = @()
    for ($i = 1; $i -lt $chars.Length; $i++) {
        $prev = [int][char]$chars[$i - 1]
        $cur = [int][char]$chars[$i]
        if ($cur -eq 214 -or $cur -eq 220 -or $cur -eq 196) {
            $hits += [pscustomobject]@{
                Index = $i
                Prev = $prev
                Cur = $cur
                PrevChar = [string]$chars[$i - 1]
                CurChar = [string]$chars[$i]
                PrevLowerAscii = ($prev -ge 97 -and $prev -le 122)
            }
        }
    }
    [pscustomobject]@{
        Text = $text
        HitCount = $hits.Count
        Hits = $hits
    } | ConvertTo-Json -Depth 5
}
finally { $pkg.Dispose() }
