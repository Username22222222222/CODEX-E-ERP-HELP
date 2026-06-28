param(
    [string]$WorkbookPath = 'C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx',
    [string]$WorksheetName = 'de-DE',
    [string]$BackupRoot = 'C:\Users\micha\Documents\X-ERP-HELP\ARCHIV'
)

$ErrorActionPreference = 'Stop'
Import-Module ImportExcel

function Replace-KnownSplitTerms {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }

    $replacements = [ordered]@{
        'D es kt op' = 'Desktop'
        'L is tV ie w' = 'ListView'
        'E di tV ie w' = 'EditView'
        'A ss is te nt en' = 'Assistenten'
        'F el dt yp en' = 'Feldtypen'
        'A rt ik el' = 'Artikel'
        'P ar tn er' = 'Partner'
        'L ag er un gs hi st or ie' = 'Lagerungshistorie'
        'P ro du kt io ns sc hr it t' = 'Produktionsschritt'
        'P ro du kt io n' = 'Produktion'
        'R es so ur ce n' = 'Ressourcen'
        'V er we nd un g' = 'Verwendung'
        'P os it io ns li st e' = 'Positionsliste'
        'U ms at z' = 'Umsatz'
        'O pt io ne n' = 'Optionen'
        'L ie fe ra dr es se n' = 'Lieferadressen'
        'O ff en e' = 'Offene'
        'P os te n' = 'Posten'
        'M ak ro' = 'Makro'
        'S et s' = 'Sets'
        'S et' = 'Set'
        'b ei' = 'bei'
        'E-M ai l' = 'E-Mail'
        'Gül ti gk ei t' = 'Gültigkeit'
        'Lager pl at z' = 'Lagerplatz'
        'P ac ks tüc k' = 'Packstück'
        'F ra ch tfüh re r' = 'Frachtführer'
        'G erät ea kt e' = 'Geräteakte'
        'F el düb er na hm e' = 'Feldübernahme'
        'M enü u nd N av ig at io n' = 'Menü und Navigation'
        'S ch al tf läc he n' = 'Schaltflächen'
        'Gül ti gk ei ts be re ic h' = 'Gültigkeitsbereich'
        'Partner an bi nd un g' = 'Partneranbindung'
        'Umsatzs te ue r-V or an me ld un g' = 'Umsatzsteuer-Voranmeldung'
        'L ei tw eg-I D' = 'Leitweg-ID'
        'Lös ch un g u nd S pe rr un g' = 'Löschung und Sperrung'
        'Datenbank-Üb er se tz un ge n' = 'Datenbank-Übersetzungen'
        'W ar tu ng sa rbei te n' = 'Wartungsarbeiten'
        'Einkaufspreis li st en' = 'Einkaufspreislisten'
        ' ß ' = ' – '
    }

    $result = $Text
    foreach ($entry in $replacements.GetEnumerator()) {
        $result = $result.Replace($entry.Key, $entry.Value)
    }

    return $result
}

function Get-SpacedVariant {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $Text = $Text -creplace '([a-z0-9])([A-Z])', '$1 $2'
    $Text = $Text -replace '_+', ' '

    $sb = [System.Text.StringBuilder]::new()
    $parts = [regex]::Matches($Text, '\p{L}+|\p{N}+|[^\p{L}\p{N}]+')
    foreach ($part in $parts) {
        $v = $part.Value
        if ($v -match '^\p{L}+$' -and $v.Length -gt 2) {
            [void]$sb.Append($v.Substring(0, 1))
            $i = 1
            while ($i -lt $v.Length) {
                $take = [Math]::Min(2, $v.Length - $i)
                [void]$sb.Append(' ')
                [void]$sb.Append($v.Substring($i, $take))
                $i += $take
            }
        }
        else {
            [void]$sb.Append($v)
        }
    }

    return $sb.ToString()
}

function Get-SpacedVariantRaw {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $Text = $Text -replace '_+', ' '
    $sb = [System.Text.StringBuilder]::new()
    $parts = [regex]::Matches($Text, '\p{L}+|\p{N}+|[^\p{L}\p{N}]+')
    foreach ($part in $parts) {
        $v = $part.Value
        if ($v -match '^\p{L}+$' -and $v.Length -gt 2) {
            [void]$sb.Append($v.Substring(0, 1))
            $i = 1
            while ($i -lt $v.Length) {
                $take = [Math]::Min(2, $v.Length - $i)
                [void]$sb.Append(' ')
                [void]$sb.Append($v.Substring($i, $take))
                $i += $take
            }
        }
        else {
            [void]$sb.Append($v)
        }
    }
    return $sb.ToString()
}

function Replace-GeneratedSplitNames {
    param(
        [string]$Text,
        [string[]]$Names
    )

    $result = $Text
    foreach ($name in $Names) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $cleanNames = New-Object System.Collections.Generic.List[string]
        $cleanNames.Add($name)
        $cleanNames.Add((Display-CleanName $name))
        if ($name.EndsWith('Edit')) {
            $cleanNames.Add($name.Substring(0, $name.Length - 4))
            $cleanNames.Add((Display-CleanName $name.Substring(0, $name.Length - 4)))
        }
        if ($name.EndsWith('Setting')) {
            $cleanNames.Add($name.Substring(0, $name.Length - 7))
            $cleanNames.Add((Display-CleanName $name.Substring(0, $name.Length - 7)))
        }

        foreach ($clean in ($cleanNames | Select-Object -Unique)) {
            foreach ($variant in @((Get-SpacedVariant $clean), (Get-SpacedVariantRaw $clean))) {
                if (-not [string]::IsNullOrWhiteSpace($variant) -and $variant -ne $clean) {
                    $result = $result.Replace($variant, $clean)
                }
            }
            if ($clean.Length -ge 4) {
                $chars = [char[]]$clean
                $parts = New-Object System.Collections.Generic.List[string]
                foreach ($ch in $chars) {
                    $parts.Add([regex]::Escape([string]$ch))
                }
                $pattern = '(?<![\p{L}\p{N}])' + ($parts -join '\s*') + '(?![\p{L}\p{N}])'
                $result = [regex]::Replace($result, $pattern, {
                    param($m)
                    if ($m.Value -match '\s') { return $clean }
                    return $m.Value
                })
            }
        }
    }

    return $result
}

function Display-CleanName {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $x = $Text -replace '_+', ' '
    $x = $x -creplace '([a-z0-9])([A-Z])', '$1 $2'
    $x = $x -replace '\s+', ' '
    return $x.Trim()
}

function Get-CleanAncestorTitle {
    param(
        [hashtable]$Stack,
        [int]$Level
    )

    for ($l = $Level - 1; $l -ge 0; $l--) {
        if ($Stack.ContainsKey($l) -and -not [string]::IsNullOrWhiteSpace($Stack[$l])) {
            return $Stack[$l]
        }
    }

    return ''
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

New-Item -ItemType Directory -Force -Path $BackupRoot | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $BackupRoot ("X-ERP-HELP-before-controlled-text-repair-$stamp.xlsx")
Copy-Item -LiteralPath $WorkbookPath -Destination $backup -Force

$pkg = Open-ExcelPackage -Path $WorkbookPath
$ws = $pkg.Workbook.Worksheets[$WorksheetName]
if (-not $ws) {
    Close-ExcelPackage $pkg -NoSave
    throw "Worksheet not found: $WorksheetName"
}

$changed = 0
$examples = New-Object System.Collections.Generic.List[object]
$stack = @{}

for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
    $level = [int]$ws.Row($r).OutlineLevel
    $title = $ws.Cells.Item($r, 1).Text
    if (-not [string]::IsNullOrWhiteSpace($title)) {
        $stack[$level] = $title
        foreach ($k in @($stack.Keys)) {
            if ([int]$k -gt $level) {
                $stack.Remove($k)
            }
        }
    }

    $old = $ws.Cells.Item($r, 8).Text
    if ([string]::IsNullOrWhiteSpace($old)) {
        continue
    }

    $names = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($title)) {
        $names.Add($title)
    }
    foreach ($k in ($stack.Keys | Sort-Object {[int]$_})) {
        if (-not [string]::IsNullOrWhiteSpace($stack[$k])) {
            $names.Add($stack[$k])
        }
    }

    $new = Replace-KnownSplitTerms $old
    $new = Replace-GeneratedSplitNames -Text $new -Names $names.ToArray()

    if ($r -eq 758) { $new = $new -replace 'P\s+ac\s+ks\s+\S+\s+k', $title }
    if ($r -eq 762) { $new = $new -replace 'F\s+ra\s+ch\s+tf\S+h\s+re\s+r', $title }
    if ($r -eq 779) { $new = $new -replace 'G\s+er\S+t\s+ea\s+kt\s+e', $title }
    if ($r -eq 807) { $new = $new -replace 'G\S+l\s+ti\s+gk\s+ei\s+t', $title }
    if ($r -eq 848) { $new = $new -replace 'F\s+el\s+d\S+b\s+er\s+na\s+hm\s+e', $title }
    if ($r -eq 854) { $new = $new -replace 'M\s+en\S+\s+u\s+nd\s+N\s+av\s+ig\s+at\s+io\s+n', $title }
    if ($r -eq 855) { $new = $new -replace 'S\s+ch\s+al\s+tf\s+l\S+c\s+he\s+n', $title }
    if ($r -eq 875) { $new = $new -replace 'G\S+l\s+ti\s+gk\s+ei\s+ts\s+be\s+re\s+ic\s+h', $title }
    if ($r -eq 907) { $new = $new -replace 'L\S+s\s+ch\s+un\s+g\s+u\s+nd\s+S\s+pe\s+rr\s+un\s+g', $title }
    if ($r -eq 926) { $new = $new -replace 'Datenbank-\S+b\s+er\s+se\s+tz\s+un\s+ge\s+n', $title }
    if ($r -eq 883) { $new = $new -replace 'Austausch\s+\S\s+Formate', ('Austausch ' + [char]0x2013 + ' Formate') }

    if ($new -match '^Der Bereich .+ die Einstellungen und Informationen, die in der Ansicht .+ fachlichen Abschnitt relevant sind\.$') {
        $ancestor = Get-CleanAncestorTitle -Stack $stack -Level $level
        if (-not [string]::IsNullOrWhiteSpace($ancestor)) {
            $area = $ws.Cells.Item($r, 1).Text
            $u = [string][char]0x00FC
            $new = "Der Bereich $area b${u}ndelt die Einstellungen und Informationen, die in der Ansicht $ancestor f${u}r diesen fachlichen Abschnitt relevant sind."
        }
    }

    if ($new -ne $old) {
        $ws.Cells.Item($r, 8).Value = $new
        $changed++
        if ($examples.Count -lt 20) {
            $examples.Add([pscustomobject]@{
                Row = $r
                Thema = $title
                Before = $old
                After = $new
            })
        }
    }
}

Close-ExcelPackage $pkg

[pscustomobject]@{
    Workbook = $WorkbookPath
    Backup = $backup
    ChangedRows = $changed
    Examples = $examples
} | ConvertTo-Json -Depth 5
