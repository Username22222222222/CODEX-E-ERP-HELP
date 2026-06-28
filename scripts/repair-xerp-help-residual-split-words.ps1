param(
    [string]$WorkbookPath = "C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx",
    [string]$WorksheetName = "de-DE",
    [string]$BackupRoot = "C:\Users\micha\Documents\X-ERP-HELP\ARCHIV"
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Import-Module ImportExcel

function Get-SpacedRegex {
    param([string]$Text)

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($ch in [char[]]$Text) {
        if ([char]::IsLetterOrDigit($ch)) {
            $parts.Add([regex]::Escape([string]$ch))
        }
        else {
            $parts.Add('\s*' + [regex]::Escape([string]$ch) + '\s*')
        }
    }

    return '(?<![\p{L}\p{N}])' + ($parts -join '\s*') + '(?![\p{L}\p{N}])'
}

function Get-CleanTechnicalBases {
    param([string]$Title)

    $bases = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Title)) { return $bases }

    $bases.Add($Title)
    if ($Title.EndsWith('Wizard')) { $bases.Add($Title.Substring(0, $Title.Length - 6)) }
    if ($Title.EndsWith('Edit')) { $bases.Add($Title.Substring(0, $Title.Length - 4)) }
    if ($Title.EndsWith('Designer')) { $bases.Add($Title.Substring(0, $Title.Length - 8)) }
    return @($bases | Where-Object { $_.Length -ge 4 } | Select-Object -Unique)
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

New-Item -ItemType Directory -Force -Path $BackupRoot | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $BackupRoot ("X-ERP-HELP-before-residual-split-repair-$stamp.xlsx")
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
try {
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $ws -or $null -eq $ws.Dimension) {
        throw "Worksheet not found or empty: $WorksheetName"
    }

    $changed = 0
    $examples = New-Object System.Collections.Generic.List[object]
    $stack = @{}

    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $level = [int]$ws.Row($r).OutlineLevel
        $topic = [string]$ws.Cells.Item($r, 1).Text
        if (-not [string]::IsNullOrWhiteSpace($topic)) {
            $stack[$level] = $topic
            foreach ($k in @($stack.Keys)) {
                if ([int]$k -gt $level) { $stack.Remove($k) }
            }
        }

        $old = [string]$ws.Cells.Item($r, 8).Text
        if ([string]::IsNullOrWhiteSpace($old)) { continue }

        $new = $old
        $new = $new -replace 'G\s+ut\s+sc\s+hr\s+if\s+t', 'Gutschrift'
        $new = $new -replace 'G\s+ru\s+pp\s+en', 'Gruppen'
        $new = $new -replace ('f' + [char]0x00FC + 'r ' + [char]0x00FC + 'bersicht'), ('f' + [char]0x00FC + 'r ' + [char]0x00DC + 'bersicht')
        $new = $new -replace ('Ein Zusatzfeld n' + [char]0x00FC + 'tzt'), ('Ein Extra-Feld n' + [char]0x00FC + 'tzt')

        $names = New-Object System.Collections.Generic.List[string]
        foreach ($k in ($stack.Keys | Sort-Object {[int]$_})) {
            foreach ($base in (Get-CleanTechnicalBases -Title ([string]$stack[$k]))) {
                $names.Add($base)
            }
        }
        foreach ($base in (Get-CleanTechnicalBases -Title $topic)) {
            $names.Add($base)
        }

        foreach ($name in ($names | Select-Object -Unique)) {
            $pattern = Get-SpacedRegex -Text $name
            $new = [regex]::Replace($new, $pattern, {
                param($m)
                if ($m.Value -match '\s') { return $name }
                return $m.Value
            })
        }

        $new = $new -replace 'in der ([A-Za-z0-9]+)-Assistent\.', 'im $1-Assistenten.'

        if ($new -ne $old) {
            $ws.Cells.Item($r, 8).Value = $new
            $changed++
            if ($examples.Count -lt 20) {
                $examples.Add([pscustomobject]@{
                    Row = $r
                    Topic = $topic
                    Before = $old
                    After = $new
                })
            }
        }
    }

    Close-ExcelPackage $pkg
    $pkg = $null

    [pscustomobject]@{
        Workbook = $WorkbookPath
        Backup = $backup
        ChangedRows = $changed
        Examples = $examples
    } | ConvertTo-Json -Depth 5
}
finally {
    if ($null -ne $pkg) { $pkg.Dispose() }
}
