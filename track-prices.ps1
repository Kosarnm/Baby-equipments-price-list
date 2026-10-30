<#
  Baby price tracker for Prisjakt.no

  Reads baby-list.csv, fetches the current lowest price for every row that has a
  PrisjaktUrl, appends it to price-history.csv and builds report.html.

  Run:  powershell -ExecutionPolicy Bypass -File .\track-prices.ps1
  Options:
    -ListFile / -HistoryFile / -ReportFile   use other file names
    -JsonFile                                also write the list + history as JSON (for the phone app)
    -NoOpen                                  don't open the report afterwards
#>
param(
    [string]$ListFile    = (Join-Path $PSScriptRoot 'baby-list.csv'),
    [string]$HistoryFile = (Join-Path $PSScriptRoot 'price-history.csv'),
    [string]$ReportFile  = (Join-Path $PSScriptRoot 'report.html'),
    [string]$JsonFile,
    [switch]$NoOpen
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$headers = @{
    'User-Agent'      = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
    'Accept-Language' = 'nb-NO,nb;q=0.9,en;q=0.8'
}
$nb = [Globalization.CultureInfo]::GetCultureInfo('nb-NO')

function Format-Kr($value) {
    if ($null -eq $value -or $value -eq '') { return '&ndash;' }
    return ([double]$value).ToString('N0', $nb) + ' kr'
}

function HtmlEnc($s) { return [Net.WebUtility]::HtmlEncode([string]$s) }

function Get-PrisjaktPrice([string]$url) {
    $resp  = Invoke-WebRequest -UseBasicParsing -Uri $url -Headers $headers -TimeoutSec 30
    # Decode as UTF-8 ourselves so Norwegian letters survive.
    $resp.RawContentStream.Position = 0
    $html = (New-Object IO.StreamReader($resp.RawContentStream, [Text.Encoding]::UTF8)).ReadToEnd()

    $name = ([regex]::Match($html, '<meta property="og:title" content="([^"]+)"')).Groups[1].Value
    $name = [Net.WebUtility]::HtmlDecode(($name -replace '\s*-\s*Sammenlign priser hos Prisjakt$', ''))

    # A link to a specific variant (e.g. ...&a-116257=Black) gets that variant's price.
    $variant = $null
    if ($url -match '[?&](a-\d+=[^&#]+)') { $variant = [Uri]::UnescapeDataString($Matches[1]).Replace('+', ' ') }

    $low = $null; $offers = $null
    foreach ($m in [regex]::Matches($html, '<script[^>]*application/ld\+json[^>]*>(.*?)</script>', 'Singleline')) {
        try { $ld = $m.Groups[1].Value | ConvertFrom-Json } catch { continue }
        $candidates = @()
        if ($ld.offers) { $candidates += $ld }
        if ($ld.hasVariant) {
            if ($variant) {
                $candidates = @($ld.hasVariant | Where-Object {
                    [Uri]::UnescapeDataString([string]$_.url).Replace('+', ' ') -like "*$variant*" })
            } else {
                $candidates += $ld.hasVariant
            }
        }
        foreach ($c in $candidates) {
            $o = $c.offers
            $p = if ($o.lowPrice) { [double]$o.lowPrice } elseif ($o.price) { [double]$o.price } else { $null }
            if ($p -and ($null -eq $low -or $p -lt $low)) { $low = $p; $offers = $o.offerCount }
        }
    }

    # Each offer row has screen-reader text like "Shop, product title, 7 688 kr".
    $shopOffers = foreach ($m in [regex]::Matches($html, '<span class="sr-only">([^<]+?), (\d[\d\s  ]*)[\s  ]kr</span>')) {
        [pscustomobject]@{
            Shop  = [Net.WebUtility]::HtmlDecode(($m.Groups[1].Value -split ',')[0].Trim())
            Price = [double](($m.Groups[2].Value) -replace '[\s  ]', '')
        }
    }

    # The structured data only has per-colour prices, so it misses offers listed as "any colour".
    # Unless a specific variant was asked for, use the page's own "cheapest right now" price.
    if (-not $variant) {
        $m = [regex]::Match($html, 'Den billigste prisen for .{1,300}? akkurat nå er (\d[\d\s  ]*)[\s  ]kr')
        if ($m.Success) {
            $low = [double]($m.Groups[1].Value -replace '[\s  ]', '')
        } elseif ($shopOffers) {
            $low = ($shopOffers | Measure-Object Price -Minimum).Minimum
        }
    }
    if ($null -eq $low -and $shopOffers) { $low = ($shopOffers | Measure-Object Price -Minimum).Minimum }

    $shop = ''
    $cheapest = $shopOffers | Where-Object { $_.Price -eq $low } | Select-Object -First 1
    if ($cheapest) { $shop = $cheapest.Shop }

    if ($null -eq $low) { throw 'no price found on page' }
    return [pscustomobject]@{ Name = $name; Price = $low; Shop = $shop; Offers = $offers }
}

# ---------------------------------------------------------------- fetch prices
# Norwegian Excel saves CSV with ';' - detect the delimiter from the header line.
$delim = if ((Get-Content -Path $ListFile -TotalCount 1 -Encoding UTF8) -match ';') { ';' } else { ',' }
$list = @(Import-Csv -Path $ListFile -Delimiter $delim -Encoding UTF8)
$now  = Get-Date -Format 'yyyy-MM-dd HH:mm'
$new  = @()
$tracked = @($list | Where-Object { $_.PrisjaktUrl -match '^https?://(www\.)?prisjakt\.no/' })

Write-Host "Checking $($tracked.Count) product(s) on Prisjakt..."
foreach ($row in $tracked) {
    try {
        $r = Get-PrisjaktPrice $row.PrisjaktUrl.Trim()
        Write-Host ("  {0,-40} {1,12}  {2}" -f $row.Item, (Format-Kr $r.Price), $r.Shop)
        $new += [pscustomobject]@{
            Date = $now; Item = $row.Item; Url = $row.PrisjaktUrl.Trim(); Product = $r.Name
            Price = $r.Price; Shop = $r.Shop; Offers = $r.Offers
        }
    } catch {
        Write-Warning "  $($row.Item): $($_.Exception.Message)"
    }
    Start-Sleep -Seconds (Get-Random -Minimum 2 -Maximum 5)   # be polite to Prisjakt
}
if ($new.Count) {
    $new | Export-Csv -Path $HistoryFile -Append -NoTypeInformation -Encoding UTF8
}

# ---------------------------------------------------------------- build report
$history = if (Test-Path $HistoryFile) { @(Import-Csv -Path $HistoryFile -Encoding UTF8) } else { @() }

if ($JsonFile) {
    $items = foreach ($row in $list) {
        $url = $row.PrisjaktUrl.Trim()
        [ordered]@{
            category = $row.Category; item = $row.Item; search = $row.PrisjaktSearch; priority = $row.Priority
            blackWeek = ($row.BlackWeek -eq 'Yes'); qty = [Math]::Max(1, [int]("0$($row.Qty)" -replace '\D', ''))
            url = $url; target = $(if ($row.TargetPrice -match '\d') { [double]($row.TargetPrice -replace '[^\d.]', '') } else { $null })
            bought = ($row.Bought -match '^(y|yes|ja|x|1|true)$'); notes = $row.Notes
            history = @($history | Where-Object { $url -and $_.Url -eq $url } | Sort-Object Date | ForEach-Object {
                [ordered]@{ date = $_.Date; price = [double]$_.Price; shop = $_.Shop; product = $_.Product } })
        }
    }
    $json = [ordered]@{ updated = $now; checked = $tracked.Count; ok = $new.Count; items = @($items) } |
        ConvertTo-Json -Depth 6
    [IO.File]::WriteAllText($JsonFile, $json, (New-Object Text.UTF8Encoding($false)))
    Write-Host "Data written to $JsonFile"
}

function Get-Sparkline($points) {
    if ($points.Count -lt 2) { return '' }
    $vals = $points | ForEach-Object { [double]$_.Price }
    $min = ($vals | Measure-Object -Minimum).Minimum; $max = ($vals | Measure-Object -Maximum).Maximum
    $w = 120; $h = 28; $span = [Math]::Max($max - $min, 1)
    $coords = for ($i = 0; $i -lt $vals.Count; $i++) {
        $x = [Math]::Round($i * $w / ($vals.Count - 1), 1)
        $y = [Math]::Round($h - 2 - ($vals[$i] - $min) * ($h - 4) / $span, 1)
        "$x,$y"
    }
    $pts = $coords -join ' '
    return "<svg width='$w' height='$h' viewBox='0 0 $w $h'><polyline fill='none' stroke='currentColor' stroke-width='1.5' points='$pts'/></svg>"
}

$rowsHtml = New-Object Text.StringBuilder
$totalNow = 0; $totalLowest = 0; $totalFirst = 0; $openCount = 0; $hits = 0
foreach ($cat in ($list | Group-Object Category)) {
    [void]$rowsHtml.Append("<tr class='cat'><th colspan='8'>$(HtmlEnc $cat.Name)</th></tr>")
    foreach ($row in $cat.Group) {
        $qty    = [Math]::Max(1, [int]("0$($row.Qty)" -replace '\D', ''))
        $bought = $row.Bought -match '^(y|yes|ja|x|1|true)$'
        if (-not $bought) { $openCount++ }
        $url    = $row.PrisjaktUrl.Trim()
        $pts    = @($history | Where-Object { $_.Url -eq $url } | Sort-Object Date)

        $nowP = $lowP = $firstP = $null; $prod = ''; $shop = ''; $lowDate = ''
        if ($pts.Count) {
            $last = $pts[-1]; $nowP = [double]$last.Price; $prod = $last.Product; $shop = $last.Shop
            $firstP = [double]$pts[0].Price
            $lowRow = $pts | Sort-Object { [double]$_.Price }, Date | Select-Object -First 1
            $lowP = [double]$lowRow.Price; $lowDate = $lowRow.Date.Substring(0, 10)
            if (-not $bought) {
                $totalNow += $nowP * $qty; $totalLowest += $lowP * $qty; $totalFirst += $firstP * $qty
            }
        }

        $target = if ($row.TargetPrice -match '\d') { [double]($row.TargetPrice -replace '[^\d.]', '') } else { $null }
        $badges = @()
        if ($bought) { $badges += "<span class='b done'>Bought</span>" }
        if ($row.BlackWeek -eq 'Yes') { $badges += "<span class='b bw'>Black Week</span>" }
        if ($row.Priority -eq 'Nice') { $badges += "<span class='b nice'>Nice to have</span>" }
        if ($nowP -and $target -and $nowP -le $target -and -not $bought) { $badges += "<span class='b hit'>Target hit!</span>"; $hits++ }
        if ($nowP -and $lowP -and $pts.Count -gt 1 -and $nowP -le $lowP) { $badges += "<span class='b low'>Lowest so far</span>" }

        $change = ''
        if ($nowP -and $firstP -and $firstP -ne $nowP) {
            $pct = [Math]::Round(($nowP - $firstP) / $firstP * 100)
            $cls = if ($pct -lt 0) { 'down' } else { 'up' }
            $change = "<span class='$cls'>$(if ($pct -gt 0) {'+'})$pct %</span>"
        }

        $itemCell = HtmlEnc $row.Item
        if ($url) {
            $label = if ($prod) { $prod } else { 'Prisjakt' }
            $itemCell += "<div class='sub'><a href='$(HtmlEnc $url)' target='_blank'>$(HtmlEnc $label)</a></div>"
        } else {
            $q = [Uri]::EscapeDataString($(if ($row.PrisjaktSearch) { $row.PrisjaktSearch } else { $row.Item -replace '\s*\(.*\)', '' }))
            $itemCell += "<div class='sub'><a href='https://www.prisjakt.no/search?search=$q' target='_blank'>Find on Prisjakt &rarr;</a></div>"
        }
        if ($row.Notes) { $itemCell += "<div class='note'>$(HtmlEnc $row.Notes)</div>" }

        $trClass = if ($bought) { " class='bought'" } else { '' }
        [void]$rowsHtml.Append(@"
<tr$trClass>
  <td>$itemCell<div>$($badges -join ' ')</div></td>
  <td class='n'>$qty</td>
  <td class='n big'>$(Format-Kr $nowP)<div class='sub'>$(HtmlEnc $shop)</div></td>
  <td class='n'>$(Format-Kr $lowP)<div class='sub'>$lowDate</div></td>
  <td class='n'>$(Format-Kr $firstP)</td>
  <td class='n'>$change</td>
  <td class='n'>$(Format-Kr $target)</td>
  <td class='spark'>$(Get-Sparkline $pts)</td>
</tr>
"@)
    }
}

$blackFriday = [datetime]'2026-11-27'
$daysLeft = [Math]::Max(0, ($blackFriday - (Get-Date).Date).Days)
$saved = $totalFirst - $totalNow
$trackedCount = @($list | Where-Object { $_.PrisjaktUrl }).Count

$report = @"
<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Baby Price Tracker</title>
<style>
:root{--bg:#faf8f5;--card:#fff;--fg:#2b2a28;--muted:#77736c;--line:#e8e3dc;--accent:#c2603d;--good:#2f7d4f;--bad:#b3403a;--chip:#f1ece5}
@media (prefers-color-scheme:dark){:root{--bg:#1c1b19;--card:#262522;--fg:#ece9e4;--muted:#a19c94;--line:#3a3834;--accent:#e08a66;--good:#6cc290;--bad:#e2847e;--chip:#33312d}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif}
main{max-width:1150px;margin:0 auto;padding:24px 16px 60px}
h1{margin:0 0 4px;font-size:26px}.lead{color:var(--muted);margin:0 0 20px}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:22px}
.stat{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.stat b{display:block;font-size:22px}.stat span{color:var(--muted);font-size:13px}
.wrap{overflow-x:auto;background:var(--card);border:1px solid var(--line);border-radius:12px}
table{border-collapse:collapse;width:100%;min-width:860px}
th,td{padding:9px 12px;border-bottom:1px solid var(--line);vertical-align:top;text-align:left}
thead th{font-size:12px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);position:sticky;top:0;background:var(--card)}
tr.cat th{background:var(--chip);font-size:14px}
td.n,th.n{text-align:right;white-space:nowrap}.big{font-weight:600}
.sub{font-size:12px;color:var(--muted)}.note{font-size:12px;color:var(--muted);font-style:italic;max-width:420px}
a{color:var(--accent)}.down{color:var(--good);font-weight:600}.up{color:var(--bad);font-weight:600}
.spark{color:var(--accent)}
.b{display:inline-block;font-size:11px;padding:1px 7px;border-radius:99px;background:var(--chip);margin-top:4px}
.b.bw{background:#2b2a28;color:#fff}.b.hit{background:var(--good);color:#fff}.b.low{background:var(--accent);color:#fff}.b.done{background:var(--good);color:#fff}
tr.bought td{opacity:.5}
</style></head><body><main>
<h1>Baby price tracker</h1>
<p class="lead">Updated $now &middot; prices from Prisjakt.no &middot; Black Friday is 27 Nov 2026 ($daysLeft days to go)</p>
<div class="stats">
  <div class="stat"><b>$openCount / $($list.Count)</b><span>items still to buy</span></div>
  <div class="stat"><b>$trackedCount</b><span>items with a Prisjakt link</span></div>
  <div class="stat"><b>$(Format-Kr $totalNow)</b><span>today's price for the tracked items you still need</span></div>
  <div class="stat"><b>$(Format-Kr $totalLowest)</b><span>if every item were at its lowest price seen</span></div>
  <div class="stat"><b class="$(if ($saved -ge 0) {'down'} else {'up'})">$(Format-Kr $saved)</b><span>change since first check (positive = cheaper)</span></div>
  <div class="stat"><b>$hits</b><span>items at or below your target price</span></div>
</div>
<div class="wrap"><table>
<thead><tr><th>Item</th><th class="n">Qty</th><th class="n">Now</th><th class="n">Lowest seen</th><th class="n">First seen</th><th class="n">Change</th><th class="n">Target</th><th>Trend</th></tr></thead>
<tbody>
$($rowsHtml.ToString())
</tbody></table></div>
<p class="lead" style="margin-top:16px">To track an item, paste its Prisjakt link into the PrisjaktUrl column of baby-list.csv. You can also set a TargetPrice and mark items as bought (Bought = yes).</p>
</main></body></html>
"@
[IO.File]::WriteAllText($ReportFile, $report, (New-Object Text.UTF8Encoding($false)))
Write-Host "Report written to $ReportFile"
if (-not $NoOpen) { Start-Process $ReportFile }
# Fail loudly (e.g. in GitHub Actions) if Prisjakt refused every request.
if ($tracked.Count -and -not $new.Count) { Write-Error 'Could not read any prices from Prisjakt.' }
