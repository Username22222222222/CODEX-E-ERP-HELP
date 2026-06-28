param(
    [string]$BaseUrl = "http://127.0.0.1:18080/de/help/views/"
)

$ErrorActionPreference = 'Stop'

$pages = @('index.html', 'article-edit.html', 'partner-edit.html', 'finance-payment-edit.html')
$results = foreach ($page in $pages) {
    $url = [Uri]::new([Uri]$BaseUrl, $page).AbsoluteUri
    $response = Invoke-WebRequest -UseBasicParsing -Uri $url -TimeoutSec 10
    $html = [string]$response.Content

    [pscustomobject]@{
        Page = $page
        StatusCode = [int]$response.StatusCode
        HasTitle = $html -match '<title>[^<]+</title>'
        HasH1 = $html -match '<h1>[^<]+</h1>'
        HasMetaDescription = $html -match '<meta name="description" content="[^"]{40,155}"'
        HasStructuredData = $html.Contains('application/ld+json')
        HasCanonical = $html -match '<link rel="canonical" href="[^"]+"'
        HasScreenshot = if ($page -eq 'index.html') { $true } else { $html -match '/de/help/Screenshots/[^"]+\.webp' }
        HasOldZusatzTerm = $html -match 'Zusatzfeld|Zusatzfelder|Zusatz-Feld|Zusatz-Felder'
        HasBrokenFieldCount = $html -match '>Felder,\s*\d+\s*Register'
    }
}

$failed = @($results | Where-Object {
    $_.StatusCode -ne 200 -or
    -not $_.HasTitle -or
    -not $_.HasH1 -or
    -not $_.HasMetaDescription -or
    -not $_.HasStructuredData -or
    -not $_.HasCanonical -or
    -not $_.HasScreenshot -or
    $_.HasOldZusatzTerm -or
    $_.HasBrokenFieldCount
})

[pscustomobject]@{
    Checked = $results.Count
    Failed = $failed.Count
    Results = $results
} | ConvertTo-Json -Depth 4
