# GOFO Seattle Region Ops Console · 西雅图片区站点数据看板

Static site, same logic as the NC console: ZIP-level network map (Stations / DSP / Routes / Price / Volume views) plus a DSP pricing subpage, driven by a repo-committed price sheet.

```
index.html                      Network map (main page) — data embedded, works offline too
analytics/dsp-pricing.html      DSP pricing — avg price by station (linked from map sidebar)
data/price_adjustments.csv      Price sheet — edit to adjust ZIP prices without rebuilding
data/stations.csv               Station addresses, pin coordinates, station specialist
data/source/sea_zip_price.xlsx  Source workbook (SEA邮编价格.xlsx)
tools/build.py                  Regenerates all embedded data from the workbook
tools/zcta_pnw.min.geojson      ZIP boundaries for all WA/OR/ID ZIPs (Census 2020 ZCTA, simplified)
```

## 1. First push (VS Code terminal)

Create an **empty** repo on github.com named `sea-ops-console` (no README, no .gitignore). Then, inside this folder:

```bash
git init
git add .
git commit -m "Seattle region ops console"
git branch -M main
git remote add origin https://github.com/vincezez/sea-ops-console.git
git push -u origin main
```

## 2. Turn on GitHub Pages

Repo → **Settings → Pages → Source: Deploy from a branch → Branch: main, folder: / (root) → Save**.
After about a minute the link is live: **https://vincezez.github.io/sea-ops-console/**

Public repos work on the free plan; private repos need GitHub Pro for Pages.

## 3. Daily update — new volumes or prices from Excel

1. Replace `data/source/sea_zip_price.xlsx` with the new workbook (same 6 columns, one sheet per station).
2. Run the build (needs Python 3 and `pip install openpyxl`):
   ```bash
   python tools/build.py
   ```
   It prints station totals and flags anything odd: ZIPs with no map boundary, missing prices, a ZIP listed under two stations, and any manual CSV price edits it replaced.
3. Commit and push:
   ```bash
   git add . && git commit -m "Data update $(date +%F)" && git push
   ```

New ZIPs anywhere in WA/OR/ID get boundaries automatically. New stations get a color automatically.

## Quick price tweak (no rebuild)

Edit `data/price_adjustments.csv` (columns `zip,price,station,route`; only the first two are read), commit, push. Map colors, route/station/DSP averages and rankings recompute on page load. The sidebar's "Price sheet" line shows how many ZIPs differ from base.

Note: running `tools/build.py` rewrites this CSV from the workbook, so put lasting price changes in the Excel.

## Station pins

`data/stations.csv` holds each station's address, `lat`, `lon`, specialist name (`head`) and pinyin. While `lat`/`lon` are blank, the pin is placed at the volume-weighted centre of the station's ZIPs and marked with an amber dot. Fill them in and rerun `python tools/build.py`.

## Local preview

Opening `index.html` by double-click works, but browsers block it from reading the CSV. To preview exactly what GitHub Pages will show:

```bash
python -m http.server 8000
```
then open http://localhost:8000 (or use the VS Code "Live Server" extension).

## Data notes

- DSP = the part of Sorting Center after the hyphen (`SEA-MIL` → `MIL`), so a DSP serving several stations is summarized network-wide.
- ZIP price shown = volume-weighted average if a ZIP appears on several rows; route, station and DSP prices are volume-weighted.
- PO-box ZIPs have no Census boundary: they count in totals and are searchable, but aren't drawn.
