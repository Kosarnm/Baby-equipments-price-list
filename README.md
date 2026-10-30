# Baby Prices

A phone app for the baby shopping list. GitHub checks the Prisjakt prices every
morning at 07:00 and the app shows today's price, the lowest price seen and a chart.

- `baby-list.csv`: the list. Edit it on github.com: paste Prisjakt links into
  `PrisjaktUrl`, set a `TargetPrice`, write `yes` under `Bought`.
- `track-prices.ps1`: fetches the prices (run by GitHub Actions).
- `price-history.csv`: every price ever seen (written automatically).
- `docs/`: the phone app (served by GitHub Pages).
- `.github/workflows/check-prices.yml`: the daily schedule.

To check prices right now: **Actions** tab → **Check Prisjakt prices** → **Run workflow**.
