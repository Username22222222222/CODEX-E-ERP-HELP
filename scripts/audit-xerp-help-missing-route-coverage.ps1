param(
    [string]$ScreenshotAuditCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\screenshot-audit.csv",
    [string]$BuildRoot = "C:\X-ERP\DOCUMENTATION\X-ERP-Help\_source\de\_build",
    [string]$OutputCsv = "C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\missing-screenshot-route-coverage.csv"
)

$ErrorActionPreference = 'Stop'

function Read-RouteJson([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
}

$routeSets = @(
    @{ Name = 'list'; Items = Read-RouteJson (Join-Path $BuildRoot 'list_routes.json') },
    @{ Name = 'direct'; Items = Read-RouteJson (Join-Path $BuildRoot 'direct_edit_routes.json') },
    @{ Name = 'popup'; Items = Read-RouteJson (Join-Path $BuildRoot 'popup_routes.json') },
    @{ Name = 'tab'; Items = Read-RouteJson (Join-Path $BuildRoot 'tab_routes.json') }
)

$rows = Import-Csv -LiteralPath $ScreenshotAuditCsv -Delimiter ';' |
    Where-Object { $_.Status -eq 'fehlt' } |
    ForEach-Object {
        $topic = $_.Topic
        $hits = foreach ($set in $routeSets) {
            foreach ($item in $set.Items) {
                $file = [string]$item.File
                $route = ([string]$item.Route).TrimStart('/')
                $editFile = [regex]::Replace($file, 'List$', 'Edit', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                $editRoute = [regex]::Replace($route, 'List$', 'Edit', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($file.Equals($topic, [StringComparison]::OrdinalIgnoreCase) -or
                    $route.Equals($topic, [StringComparison]::OrdinalIgnoreCase) -or
                    $editFile.Equals($topic, [StringComparison]::OrdinalIgnoreCase) -or
                    $editRoute.Equals($topic, [StringComparison]::OrdinalIgnoreCase)) {
                    [pscustomobject]@{
                        Kind = $set.Name
                        File = $file
                        Route = $route
                    }
                }
            }
        }
        $first = @($hits) | Select-Object -First 1
        [pscustomobject]@{
            Row = $_.Row
            Topic = $topic
            Screenshot = $_.Screenshot
            RouteKinds = (@($hits) | ForEach-Object { $_.Kind } | Select-Object -Unique) -join ','
            FirstRouteKind = if ($first) { $first.Kind } else { '' }
            FirstRouteFile = if ($first) { $first.File } else { '' }
            FirstRoute = if ($first) { $first.Route } else { '' }
            RouteHitCount = @($hits).Count
        }
    }

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputCsv) | Out-Null
$rows | Export-Csv -LiteralPath $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8

[pscustomobject]@{
    OutputCsv = $OutputCsv
    Missing = @($rows).Count
    WithRoute = @($rows | Where-Object { [int]$_.RouteHitCount -gt 0 }).Count
    WithoutRoute = @($rows | Where-Object { [int]$_.RouteHitCount -eq 0 }).Count
    ByKind = @($rows | Where-Object { $_.FirstRouteKind } | Group-Object FirstRouteKind | ForEach-Object { [pscustomobject]@{ Kind = $_.Name; Count = $_.Count } })
} | ConvertTo-Json -Depth 4
