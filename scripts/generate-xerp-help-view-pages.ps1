param(
    [string]$WorkbookPath = 'C:\Users\micha\Documents\X-ERP-HELP\X-ERP-HELP.xlsx',
    [string]$WorksheetName = 'de-DE',
    [string]$OutputRoot = 'D:\DATEN\HOMEPAGES\x-erp.de\de\help\views',
    [string]$ScreenshotWebRoot = '/de/help/Screenshots',
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
Import-Module ImportExcel

$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)

function Write-Utf8File {
    param([string]$Path, [string]$Content)
    if ($WhatIf) { return }
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, $Content, $Utf8NoBom)
}

function HtmlEncode {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Attr {
    param([string]$Text)
    return (HtmlEncode $Text).Replace('"', '&quot;')
}

function Slugify {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return 'seite' }
    $normalized = $Text.Trim()
    $normalized = $normalized -creplace '([a-z0-9])([A-Z])', '$1-$2'
    $normalized = $normalized -replace '_+', '-'
    $normalized = $normalized.ToLowerInvariant()
    $normalized = $normalized.Replace([string][char]0x00E4, 'ae')
    $normalized = $normalized.Replace([string][char]0x00F6, 'oe')
    $normalized = $normalized.Replace([string][char]0x00FC, 'ue')
    $normalized = $normalized.Replace([string][char]0x00DF, 'ss')
    $normalized = $normalized.Replace([string][char]0x00C4, 'ae')
    $normalized = $normalized.Replace([string][char]0x00D6, 'oe')
    $normalized = $normalized.Replace([string][char]0x00DC, 'ue')
    $normalized = $normalized -replace '[^a-z0-9]+', '-'
    $normalized = $normalized.Trim('-')
    if ([string]::IsNullOrWhiteSpace($normalized)) { return 'seite' }
    return $normalized
}

function DisplayName {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $x = $Text -replace '_', ' '
    $x = $x -creplace '([a-z0-9])([A-Z])', '$1 $2'
    $x = $x -replace '\s+', ' '
    return $x.Trim()
}

function Get-SpacedVariantForName {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
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

function Get-NameCandidates {
    param([string[]]$Names)
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($name in $Names) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $items.Add($name)
        $items.Add((DisplayName $name))
        if ($name.EndsWith('Edit')) {
            $base = $name.Substring(0, $name.Length - 4)
            $items.Add($base)
            $items.Add((DisplayName $base))
        }
        if ($name.EndsWith('Setting')) {
            $base = $name.Substring(0, $name.Length - 7)
            $items.Add($base)
            $items.Add((DisplayName $base))
        }
    }
    return $items | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Sort-Object Length -Descending
}

function CleanGeneratedText {
    param(
        [string]$Text,
        [string[]]$Names
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }
    $result = $Text
    foreach ($name in (Get-NameCandidates -Names $Names)) {
        $variant = Get-SpacedVariantForName $name
        if (-not [string]::IsNullOrWhiteSpace($variant) -and $variant -ne $name) {
            $result = $result.Replace($variant, $name)
        }
    }
    $explicit = [ordered]@{
        'A dm in Co mp an ySet ti ng' = 'AdminCompanySetting'
        'A dm in Pr og ra mSet ti ng' = 'AdminProgramSetting'
        'A rt ic le Sa le sP ri ce' = 'ArticleSalesPrice'
        'F in an ce Du nn in g' = 'FinanceDunning'
        'F in an ce Fi xe dA ss et De pr ec ia ti on' = 'FinanceFixedAssetDepreciation'
        'F in an ce Gr ou pPartner' = 'FinanceGroupPartner'
        'F in an ce Op en It em On Pa rt ne rP ag e' = 'FinanceOpenItemOnPartnerPage'
    }
    foreach ($entry in $explicit.GetEnumerator()) {
        $result = $result.Replace($entry.Key, $entry.Value)
    }
    return $result
}

function TrimMeta {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $x = ($Text -replace '\s+', ' ').Trim()
    if ($x.Length -gt 155) { $x = $x.Substring(0, 152).TrimEnd() + '...' }
    return $x
}

function Read-Rows {
    $pkg = Open-ExcelPackage -Path $WorkbookPath
    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
    if (-not $ws) {
        Close-ExcelPackage $pkg -NoSave
        throw "Worksheet not found: $WorksheetName"
    }

    $rows = New-Object System.Collections.Generic.List[object]
    for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
        $topic = $ws.Cells.Item($r, 1).Text
        if ([string]::IsNullOrWhiteSpace($topic)) { continue }
        $rows.Add([pscustomobject]@{
            Row = $r
            Level = [int]$ws.Row($r).OutlineLevel
            Topic = $topic
            Original = $ws.Cells.Item($r, 2).Text
            Folder = $ws.Cells.Item($r, 3).Text
            Screenshot = $ws.Cells.Item($r, 5).Text
            Field = $ws.Cells.Item($r, 6).Text
            Important = $ws.Cells.Item($r, 7).Text
            Description = $ws.Cells.Item($r, 8).Text
            PageId = $ws.Cells.Item($r, 10).Text
            Slug = $ws.Cells.Item($r, 11).Text
            Title = $ws.Cells.Item($r, 12).Text
            Meta = $ws.Cells.Item($r, 13).Text
            H1 = $ws.Cells.Item($r, 14).Text
            PrimaryKeyword = $ws.Cells.Item($r, 15).Text
            ContentType = $ws.Cells.Item($r, 16).Text
            DirectoryPath = $ws.Cells.Item($r, 22).Text
            FileName = $ws.Cells.Item($r, 23).Text
            UrlPath = $ws.Cells.Item($r, 24).Text
            NavTitle = $ws.Cells.Item($r, 26).Text
            Breadcrumb = $ws.Cells.Item($r, 27).Text
            ScreenshotRelPath = $ws.Cells.Item($r, 32).Text
            ScreenshotWebPath = $ws.Cells.Item($r, 33).Text
            ImageAlt = $ws.Cells.Item($r, 34).Text
            ImageCaption = $ws.Cells.Item($r, 35).Text
            ImageStatus = $ws.Cells.Item($r, 36).Text
        })
    }
    Close-ExcelPackage $pkg -NoSave
    return $rows
}

function Get-ChildrenUntilNextView {
    param($Rows, [int]$Index)
    $page = $Rows[$Index]
    $children = New-Object System.Collections.Generic.List[object]
    for ($i = $Index + 1; $i -lt $Rows.Count; $i++) {
        $x = $Rows[$i]
        if ($x.PageId -like 'views/*' -and $x.ContentType -eq 'View') { break }
        if ($x.Level -le $page.Level -and $x.PageId -like 'views/*') { break }
        if ($x.Row -gt $page.Row) { $children.Add($x) }
    }
    return $children
}

function Get-Tabs {
    param($Page, $Children)
    $tabs = New-Object System.Collections.Generic.List[object]
    foreach ($child in $Children) {
        if ($child.Level -eq ($Page.Level + 1) -and [string]::IsNullOrWhiteSpace($child.Field)) {
            $tabs.Add($child)
        }
    }
    return $tabs
}

function Get-FieldsForScope {
    param($Scope, $Children)
    $fields = New-Object System.Collections.Generic.List[object]
    $scopeIndex = $Children.IndexOf($Scope)
    for ($i = $scopeIndex + 1; $i -lt $Children.Count; $i++) {
        $x = $Children[$i]
        if ($x.Level -le $Scope.Level) { break }
        if (-not [string]::IsNullOrWhiteSpace($x.Field) -or -not [string]::IsNullOrWhiteSpace($x.Description)) {
            $fields.Add($x)
        }
    }
    return $fields
}

function Get-PageFieldsWithoutTabs {
    param($Page, $Children)
    $fields = New-Object System.Collections.Generic.List[object]
    foreach ($x in $Children) {
        if ($x.Level -eq ($Page.Level + 1) -and (-not [string]::IsNullOrWhiteSpace($x.Field))) {
            $fields.Add($x)
        }
    }
    return $fields
}

function Render-FieldTable {
    param($Fields, [string[]]$Names)
    if ($Fields.Count -eq 0) { return '' }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<div class="table-wrap"><table>')
    [void]$sb.AppendLine('<thead><tr><th>Feld</th><th>Technischer Name</th><th>Beschreibung</th></tr></thead><tbody>')
    foreach ($f in $Fields) {
        $cleanDesc = CleanGeneratedText -Text $f.Description -Names $Names
        [void]$sb.AppendLine(('<tr><td>{0}</td><td><code>{1}</code></td><td>{2}</td></tr>' -f (HtmlEncode $f.Topic), (HtmlEncode $f.Field), (HtmlEncode $cleanDesc)))
    }
    [void]$sb.AppendLine('</tbody></table></div>')
    return $sb.ToString()
}

function Render-Layout {
    param(
        [string]$Title,
        [string]$Meta,
        [string]$Canonical,
        [string]$Body,
        [string]$JsonLd = ''
    )
    $css = @'
:root{--bg:#161616;--panel:#1f1f1f;--panel2:#1a1a1a;--ink:#f2f2f2;--muted:#a9a9a9;--line:#343434;--accent:#67b7ff;--accent2:#79dfbf;--soft:#202f3b;--soft2:#242424;--shadow:0 18px 45px rgba(0,0,0,.32)}
*{box-sizing:border-box}body{margin:0;background:#161616;color:var(--ink);font:16px/1.58 system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif}a{color:var(--accent);text-decoration:none}a:hover{text-decoration:underline}.shell{max-width:1180px;margin:0 auto;padding:30px 22px 56px}.topbar{display:flex;align-items:center;justify-content:space-between;gap:18px;margin-bottom:26px}.brand{display:flex;align-items:baseline;gap:12px}.logo{font-weight:800;font-size:23px;letter-spacing:0;color:#fff}.tag{color:var(--muted);font-size:14px}.crumb{font-size:13px;color:var(--muted);margin:4px 0 14px}.hero{display:grid;grid-template-columns:minmax(0,1fr) 260px;gap:28px;align-items:start;border-bottom:1px solid var(--line);padding-bottom:24px;margin-bottom:26px}.hero h1{margin:0 0 12px;font-size:clamp(30px,4vw,46px);line-height:1.08;letter-spacing:0;color:#fff}.lead{font-size:18px;color:#d5d5d5;max-width:820px}.badges{display:flex;flex-wrap:wrap;gap:8px;margin-top:18px}.badge{background:var(--soft);border:1px solid #31556f;border-radius:6px;color:#c9e7ff;padding:5px 9px;font-size:13px}.toc{background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:16px;box-shadow:var(--shadow)}.toc h2{font-size:15px;margin:0 0 10px;color:#fff}.toc ul{list-style:none;margin:0;padding:0}.toc li{margin:7px 0}.grid{display:grid;grid-template-columns:260px minmax(0,1fr);gap:28px}.side{position:sticky;top:14px;align-self:start}.content{background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:26px;box-shadow:var(--shadow)}.content h2{margin:26px 0 12px;font-size:25px;color:#fff}.content h3{margin:22px 0 8px;font-size:19px;color:#fff}.content p{color:#dddddd}.screenshot{margin:20px 0 26px}.screenshot img{display:block;width:100%;height:auto;border:1px solid #3a3a3a;border-radius:8px;background:#101820;box-shadow:0 16px 36px rgba(0,0,0,.42)}.caption{font-size:13px;color:var(--muted);margin-top:8px}.summary{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin:18px 0 22px}.metric{border:1px solid var(--line);border-radius:8px;padding:12px;background:var(--soft2)}.metric b{display:block;font-size:22px;color:#fff}.metric span{color:var(--muted);font-size:13px}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:8px;margin:12px 0 22px}table{width:100%;border-collapse:collapse;background:#1b1b1b}th,td{padding:10px 12px;border-bottom:1px solid var(--line);vertical-align:top;text-align:left}th{background:#282828;font-size:13px;color:#ededed}td{color:#dedede}code{font-family:Consolas,Monaco,monospace;font-size:.92em;color:#9dd2ff}.tabs{display:grid;gap:14px}.tabcard{border:1px solid var(--line);border-radius:8px;padding:16px;background:var(--soft2)}.tabcard h3{margin-top:0}.links{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:10px}.linkcard{display:block;border:1px solid var(--line);border-radius:8px;background:var(--soft2);padding:12px;transition:border-color .15s ease,transform .15s ease,background .15s ease}.linkcard:hover{border-color:#4a7fa5;background:#2b2b2b;transform:translateY(-1px);text-decoration:none}.linkcard strong{color:#e9f5ff}.linkcard small{display:block;color:var(--muted);margin-top:3px}.footer{margin-top:30px;color:var(--muted);font-size:13px}@media (max-width:850px){.hero,.grid{grid-template-columns:1fr}.side{position:static}.summary{grid-template-columns:1fr}.shell{padding:18px 14px 38px}}
'@
    $json = ''
    if (-not [string]::IsNullOrWhiteSpace($JsonLd)) {
        $json = "<script type=`"application/ld+json`">$JsonLd</script>"
    }
    return @"
<!doctype html>
<html lang="de">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>$(HtmlEncode $Title)</title>
  <meta name="description" content="$(Attr $Meta)">
  <link rel="canonical" href="$(Attr $Canonical)">
  <style>$css</style>
  $json
</head>
<body>$Body</body>
</html>
"@
}

function Render-ViewPage {
    param($Page, $Children, $AllPages)
    $slug = if ($Page.Slug) { $Page.Slug } else { Slugify $Page.Topic }
    $h1 = if ($Page.H1) { $Page.H1 } elseif ($Page.NavTitle) { $Page.NavTitle } else { DisplayName $Page.Topic }
    $title = if ($Page.Title) { $Page.Title } else { "$h1 | X-ERP ERP Hilfe" }
    $meta = if ($Page.Meta) { TrimMeta $Page.Meta } else { TrimMeta "ERP Hilfe zu $h1 in X-ERP: Zweck, Aufbau, Felder, Register und Bedienung der View." }
    $canonical = "https://x-erp.de/de/help/views/$slug.html"
    $tabs = Get-Tabs -Page $Page -Children $Children
    $pageFields = Get-PageFieldsWithoutTabs -Page $Page -Children $Children
    $fieldCount = @($Children | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Field) }).Count
    $screenshot = if ($Page.ScreenshotRelPath) { $Page.ScreenshotRelPath } elseif ($Page.Screenshot) { $Page.Screenshot } else { '' }
    $screenshotName = Split-Path -Leaf $screenshot
    $screenshotSrc = if ($screenshotName) { "$ScreenshotWebRoot/$screenshotName" } else { '' }
    $pageNames = @($Page.Topic, $Page.Original, $Page.NavTitle, $h1)
    $desc = if ($Page.Description) { CleanGeneratedText -Text $Page.Description -Names $pageNames } else { "Diese ERP Hilfe beschreibt Zweck, Aufbau und Bedienung der X-ERP View $h1." }

    $tocItems = New-Object System.Collections.Generic.List[string]
    $tocItems.Add('<li><a href="#aufgabe">Aufgabe</a></li>')
    if ($screenshotSrc) { $tocItems.Add('<li><a href="#screenshot">Screenshot</a></li>') }
    if ($pageFields.Count -gt 0) { $tocItems.Add('<li><a href="#felder">Felder</a></li>') }
    if ($tabs.Count -gt 0) { $tocItems.Add('<li><a href="#register">Register</a></li>') }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<div class="shell">')
    [void]$sb.AppendLine('<div class="topbar"><div class="brand"><div class="logo">X-ERP Hilfe</div><div class="tag">ERP Views und Wizards</div></div><a href="./index.html">Alle Views</a></div>')
    [void]$sb.AppendLine('<div class="hero"><div>')
    [void]$sb.AppendLine(('<div class="crumb">{0}</div>' -f (HtmlEncode $Page.Breadcrumb)))
    [void]$sb.AppendLine(('<h1>{0}</h1>' -f (HtmlEncode $h1)))
    [void]$sb.AppendLine(('<p class="lead">{0}</p>' -f (HtmlEncode $desc)))
    [void]$sb.AppendLine('<div class="badges"><span class="badge">ERP</span><span class="badge">X-ERP View</span><span class="badge">Deutsche Hilfe</span></div>')
    [void]$sb.AppendLine('</div><aside class="toc"><h2>Inhalt</h2><ul>')
    foreach ($item in $tocItems) { [void]$sb.AppendLine($item) }
    [void]$sb.AppendLine('</ul></aside></div>')
    [void]$sb.AppendLine('<div class="grid"><aside class="side toc"><h2>Kennzahlen</h2>')
    [void]$sb.AppendLine(('<div class="metric"><b>{0}</b><span>Register/Abschnitte</span></div>' -f $tabs.Count))
    [void]$sb.AppendLine(('<div class="metric"><b>{0}</b><span>Felder</span></div>' -f $fieldCount))
    [void]$sb.AppendLine(('<div class="metric"><b>{0}</b><span>Quelle: Excel-Zeile</span></div>' -f $Page.Row))
    [void]$sb.AppendLine('</aside><main class="content">')
    [void]$sb.AppendLine('<section id="aufgabe"><h2>Aufgabe der View</h2>')
    [void]$sb.AppendLine(('<p>{0}</p></section>' -f (HtmlEncode $desc)))
    if ($screenshotSrc) {
        $alt = if ($Page.ImageAlt) { $Page.ImageAlt } else { "Screenshot der X-ERP View $h1" }
        $cap = if ($Page.ImageCaption) { $Page.ImageCaption } else { "Screenshot aus X-ERP: $h1." }
        [void]$sb.AppendLine('<section id="screenshot" class="screenshot"><h2>Screenshot</h2>')
        [void]$sb.AppendLine(('<img src="{0}" alt="{1}" loading="lazy">' -f (Attr $screenshotSrc), (Attr $alt)))
        [void]$sb.AppendLine(('<div class="caption">{0}</div></section>' -f (HtmlEncode $cap)))
    }
    if ($pageFields.Count -gt 0) {
        [void]$sb.AppendLine('<section id="felder"><h2>Felder im Kopfbereich</h2>')
        [void]$sb.AppendLine((Render-FieldTable -Fields $pageFields -Names $pageNames))
        [void]$sb.AppendLine('</section>')
    }
    if ($tabs.Count -gt 0) {
        [void]$sb.AppendLine('<section id="register"><h2>Register und Abschnitte</h2><div class="tabs">')
        foreach ($tab in $tabs) {
            $tabSlug = Slugify $tab.Topic
            $fields = Get-FieldsForScope -Scope $tab -Children $Children
            $tabNames = @($Page.Topic, $Page.Original, $Page.NavTitle, $h1, $tab.Topic, $tab.Original)
            [void]$sb.AppendLine(('<article class="tabcard" id="{0}"><h3>{1}</h3>' -f (Attr $tabSlug), (HtmlEncode $tab.Topic)))
            if ($tab.Description) { [void]$sb.AppendLine(('<p>{0}</p>' -f (HtmlEncode (CleanGeneratedText -Text $tab.Description -Names $tabNames)))) }
            if ($fields.Count -gt 0) { [void]$sb.AppendLine((Render-FieldTable -Fields $fields -Names $tabNames)) }
            [void]$sb.AppendLine('</article>')
        }
        [void]$sb.AppendLine('</div></section>')
    }
    [void]$sb.AppendLine('<div class="footer">Quelle: X-ERP-HELP.xlsx. Diese Seite wurde automatisch aus der deutschen Hilfe-Struktur erzeugt.</div>')
    [void]$sb.AppendLine('</main></div></div>')

    $json = @{
        '@context' = 'https://schema.org'
        '@type' = 'TechArticle'
        headline = $h1
        description = $meta
        inLanguage = 'de-DE'
        keywords = 'ERP, X-ERP, Hilfe, View'
        image = if ($screenshotSrc) { "https://x-erp.de$screenshotSrc" } else { $null }
        mainEntityOfPage = $canonical
    } | ConvertTo-Json -Depth 5 -Compress

    return [pscustomobject]@{
        Slug = $slug
        Title = $h1
        FilePath = Join-Path $OutputRoot "$slug.html"
        Html = Render-Layout -Title $title -Meta $meta -Canonical $canonical -Body $sb.ToString() -JsonLd $json
        FieldCount = [int]$fieldCount
        TabCount = [int]$tabs.Count
    }
}

function Render-IndexPage {
    param($Pages)
    $ue = [string][char]0x00FC
    $indexMeta = "Deutsche ERP Hilfe f${ue}r X-ERP Views und Wizards mit Aufbau, Feldern, Registern und Screenshots."
    $body = [System.Text.StringBuilder]::new()
    [void]$body.AppendLine('<div class="shell">')
    [void]$body.AppendLine('<div class="topbar"><div class="brand"><div class="logo">X-ERP Hilfe</div><div class="tag">ERP Views und Wizards</div></div></div>')
    [void]$body.AppendLine(('<div class="hero"><div><div class="crumb">ERP Hilfe &gt; Views</div><h1>ERP Views in X-ERP</h1><p class="lead">Diese Hilfe f{0}hrt durch die wichtigsten EditViews und Wizards in X-ERP. Jede Seite beschreibt Zweck, Aufbau, Felder, Register und die zugeh{1}rigen Screenshots.</p><div class="badges"><span class="badge">ERP</span><span class="badge">X-ERP</span><span class="badge">Views</span></div></div><aside class="toc"><h2>Umfang</h2><ul>' -f ([char]0x00FC), ([char]0x00F6)))
    [void]$body.AppendLine(('<li>{0} View-Seiten</li>' -f $Pages.Count))
    [void]$body.AppendLine(('<li>{0} dokumentierte Felder</li>' -f (($Pages | Measure-Object FieldCount -Sum).Sum)))
    [void]$body.AppendLine('</ul></aside></div>')
    [void]$body.AppendLine('<main class="content"><h2>Alle Views</h2><div class="links">')
    foreach ($p in ($Pages | Sort-Object Title)) {
        [void]$body.AppendLine(('<a class="linkcard" href="./{0}.html"><strong>{1}</strong><small>{2} Felder, {3} Register</small></a>' -f (Attr $p.Slug), (HtmlEncode $p.Title), $p.FieldCount, $p.TabCount))
    }
    [void]$body.AppendLine('</div></main><div class="footer">Quelle: X-ERP-HELP.xlsx. Automatisch erzeugte deutsche X-ERP ERP-Hilfe.</div></div>')
    $json = @{
        '@context' = 'https://schema.org'
        '@type' = 'CollectionPage'
        name = 'ERP Views in X-ERP'
        description = $indexMeta
        inLanguage = 'de-DE'
        keywords = 'ERP, X-ERP, Hilfe, Views, Wizards'
        mainEntityOfPage = 'https://x-erp.de/de/help/views/index.html'
    } | ConvertTo-Json -Depth 5 -Compress
    return Render-Layout -Title 'ERP Views in X-ERP | X-ERP Hilfe' -Meta $indexMeta -Canonical 'https://x-erp.de/de/help/views/index.html' -Body $body.ToString() -JsonLd $json
}

$rows = Read-Rows
$viewIndexes = New-Object System.Collections.Generic.List[int]
for ($i = 0; $i -lt $rows.Count; $i++) {
    if ($rows[$i].PageId -like 'views/*' -and $rows[$i].ContentType -eq 'View') {
        $viewIndexes.Add($i)
    }
}

$pages = New-Object System.Collections.Generic.List[object]
foreach ($idx in $viewIndexes) {
    $page = $rows[$idx]
    $children = Get-ChildrenUntilNextView -Rows $rows -Index $idx
    $rendered = Render-ViewPage -Page $page -Children $children -AllPages $pages
    Write-Utf8File -Path $rendered.FilePath -Content $rendered.Html
    $pages.Add($rendered)
}

$indexHtml = Render-IndexPage -Pages $pages
Write-Utf8File -Path (Join-Path $OutputRoot 'index.html') -Content $indexHtml

[pscustomobject]@{
    OutputRoot = $OutputRoot
    Pages = $pages.Count
    Fields = (($pages | Measure-Object FieldCount -Sum).Sum)
    Tabs = (($pages | Measure-Object TabCount -Sum).Sum)
    WhatIf = [bool]$WhatIf
} | ConvertTo-Json -Depth 4
