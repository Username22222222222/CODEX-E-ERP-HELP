param(
    [string]$SiteRoot = 'D:\DATEN\HOMEPAGES\x-erp.de',
    [string]$ViewsRoot = 'D:\DATEN\HOMEPAGES\x-erp.de\de\help\views'
)

$ErrorActionPreference = 'Stop'

$htmlFiles = Get-ChildItem -LiteralPath $ViewsRoot -File -Filter *.html
$generated = $htmlFiles | Where-Object {
    (Get-Content -LiteralPath $_.FullName -Raw) -like '*Quelle: X-ERP-HELP.xlsx*'
}

$missingScreenshots = New-Object System.Collections.Generic.List[object]
$missingMeta = New-Object System.Collections.Generic.List[object]
$missingStructuredData = New-Object System.Collections.Generic.List[object]

foreach ($file in $generated) {
    $text = Get-Content -LiteralPath $file.FullName -Raw
    if ($text -notmatch '<meta name="description"') {
        $missingMeta.Add($file.FullName)
    }
    if ($text -notmatch 'application/ld\+json') {
        $missingStructuredData.Add($file.FullName)
    }
    foreach ($match in [regex]::Matches($text, '<img[^>]+src="([^"]+)"')) {
        $src = $match.Groups[1].Value
        if ($src.StartsWith('/de/help/Screenshots/')) {
            $relative = $src.TrimStart('/').Replace('/', '\')
            $path = Join-Path $SiteRoot $relative
            if (-not (Test-Path -LiteralPath $path)) {
                $missingScreenshots.Add([pscustomobject]@{
                    File = $file.FullName
                    Src = $src
                    Path = $path
                })
            }
        }
    }
}

[pscustomobject]@{
    HtmlFilesAtRoot = $htmlFiles.Count
    GeneratedFiles = $generated.Count
    MissingMetaDescription = $missingMeta.Count
    MissingStructuredData = $missingStructuredData.Count
    MissingScreenshotLinks = $missingScreenshots.Count
    FirstMissingScreenshots = $missingScreenshots | Select-Object -First 10
} | ConvertTo-Json -Depth 5
