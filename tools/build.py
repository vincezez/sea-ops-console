#!/usr/bin/env python3
"""
Rebuild the Seattle ops console data from the ZIP price workbook.

    python tools/build.py                          # uses data/source/sea_zip_price.xlsx
    python tools/build.py path/to/new_sheet.xlsx   # use a different workbook

Workbook format: one sheet per station (SEA, PDX, ...), row 1 = headers, columns in order:
    Station | Sorting Center | Route | Price | Zipcode | Daily Volume
The DSP code is taken from the Sorting Center after the hyphen (SEA-MIL -> MIL).
A ZIP may appear on several rows (split across routes/DSPs); volumes are summed and the
ZIP price is the volume-weighted average of its rows.

What it writes:
    index.html                  <script id="mapdata"> and <script id="zipgeo"> lines
    analytics/dsp-pricing.html  <script id="pricedata"> line
    data/price_adjustments.csv  full price baseline (zip,price,station,route)

An optional sheet named "Address" (Station | Address | Station Head) is not read as ZIP data;
its address and head columns are copied into data/stations.csv (station codes like SEA01 match
sheet SEA). Station pins: exact lat/lon from data/stations.csv when filled in; otherwise the
centre of the ZIP in the station's address; otherwise the volume-weighted centre of the
station's own ZIPs. Both fallbacks get an amber "approximate" flag on the map.

Only dependency: openpyxl  (pip install openpyxl)
"""
import csv
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is missing. Run:  pip install openpyxl")

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_XLSX = ROOT / "data" / "source" / "sea_zip_price.xlsx"
GEO_LIB = ROOT / "tools" / "zcta_pnw.min.geojson"
STATIONS_CSV = ROOT / "data" / "stations.csv"
PRICE_CSV = ROOT / "data" / "price_adjustments.csv"
INDEX = ROOT / "index.html"
DSP_PAGE = ROOT / "analytics" / "dsp-pricing.html"


def num(v, default=None):
    try:
        if v is None or (isinstance(v, str) and not v.strip()):
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def read_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows, problems = [], []
    for ws in wb.worksheets:
        if ws.title.strip().lower() == "address":
            continue
        for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            r = list(r[:6]) + [None] * (6 - len(r[:6]))
            station, sc, route, price, zipc, vol = r
            if all(x is None or str(x).strip() == "" for x in r):
                continue
            if zipc is None or station is None:
                problems.append(f"{ws.title}!row {i}: missing station or ZIP -> skipped")
                continue
            z = str(int(float(zipc))).zfill(5) if not isinstance(zipc, str) else zipc.strip().zfill(5)
            if not re.fullmatch(r"\d{5}", z):
                problems.append(f"{ws.title}!row {i}: bad ZIP {zipc!r} -> skipped")
                continue
            sc = str(sc or "").strip()
            dsp = sc.split("-", 1)[1].strip() if "-" in sc else (sc or "UNKNOWN")
            p = num(price, 0.0)
            v = num(vol, 0.0)
            if p <= 0:
                problems.append(f"{ws.title}!row {i}: ZIP {z} has no price")
            rows.append(dict(station=str(station).strip(), dsp=dsp, sc=sc,
                             route=str(route or "").strip() or "UNROUTED",
                             price=p, zip=z, volume=int(round(v))))
    return rows, problems


def read_stations_csv():
    meta = {}
    if not STATIONS_CSV.exists():
        return meta
    with open(STATIONS_CSV, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            code = (r.get("station") or "").strip()
            if code:
                meta[code] = {k: (v or "").strip() for k, v in r.items()}
    return meta


def read_address_sheet(path):
    """Address sheet rows -> {code: (address, head)}; codes as written (e.g. SEA01)."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out = {}
    for ws in wb.worksheets:
        if ws.title.strip().lower() != "address":
            continue
        for r in ws.iter_rows(min_row=2, values_only=True):
            r = list(r[:3]) + [None] * (3 - len(r[:3]))
            code = str(r[0] or "").strip()
            if code:
                out[code] = (str(r[1] or "").strip(), str(r[2] or "").strip())
    return out


def sync_stations_csv(smeta, addr, station_codes):
    """Copy Address-sheet address/head into stations.csv (keeps lat/lon/pinyin)."""
    notes = []
    for code, (address, head) in addr.items():
        st = code if code in station_codes else next(
            (s for s in station_codes if code.upper().startswith(s.upper())), None)
        if st is None:
            notes.append(f"Address sheet: {code} matches no station sheet -> ignored")
            continue
        m = smeta.setdefault(st, {"station": st})
        if address:
            m["address"] = address
        if head:
            m["head"] = head
    cols = ["station", "address", "lat", "lon", "head", "head_pinyin"]
    order = sorted(smeta, key=lambda s: (s not in station_codes, s))
    with open(STATIONS_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(cols)
        for s in order:
            w.writerow([smeta[s].get(c, "") for c in cols])
    return notes


def wavg(pairs):
    """pairs = [(value, weight)] -> weighted avg, simple mean if all weights are 0."""
    pairs = [(v, w) for v, w in pairs if v and v > 0]
    if not pairs:
        return 0.0
    tw = sum(w for _, w in pairs)
    return sum(v * w for v, w in pairs) / tw if tw > 0 else sum(v for v, _ in pairs) / len(pairs)


def main():
    xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_XLSX
    if not xlsx.exists():
        sys.exit(f"Workbook not found: {xlsx}")
    rows, problems = read_workbook(xlsx)
    if not rows:
        sys.exit("No data rows found in workbook.")

    geo = {f["properties"]["zip"]: f for f in json.load(open(GEO_LIB, encoding="utf-8"))["features"]}
    smeta = read_stations_csv()
    addr = read_address_sheet(xlsx)
    if addr:
        problems += sync_stations_csv(smeta, addr, {r["station"] for r in rows})

    # ---- aggregate rows -> ZIP records (a ZIP belongs to one station) ----
    zips = {}
    for r in rows:
        z = zips.setdefault(r["zip"], dict(zip=r["zip"], station=r["station"], volume=0,
                                           _pv=[], dsps=defaultdict(int), routes=defaultdict(int)))
        if z["station"] != r["station"]:
            problems.append(f"ZIP {r['zip']} listed under {z['station']} and {r['station']} -> kept {z['station']}")
            continue
        z["volume"] += r["volume"]
        z["_pv"].append((r["price"], r["volume"]))
        z["dsps"][r["dsp"]] += r["volume"]
        z["routes"][r["route"]] += r["volume"]

    for z in zips.values():
        z["price"] = round(wavg(z.pop("_pv")), 3)
        g = geo.get(z["zip"])
        z["has_geom"] = g is not None
        z["lat"] = float(g["properties"]["lat"]) if g else None
        z["lon"] = float(g["properties"]["lon"]) if g else None
        z["dsps"] = dict(z["dsps"])
        z["routes"] = dict(z["routes"])

    # ---- stations ----
    by_station = defaultdict(list)
    for z in zips.values():
        by_station[z["station"]].append(z)

    stations = {}
    for st, zl in sorted(by_station.items()):
        zl.sort(key=lambda z: -z["volume"])
        tot = sum(z["volume"] for z in zl)
        # routes
        rmem = defaultdict(list)
        rdsp = defaultdict(lambda: defaultdict(float))
        for z in zl:
            for rt, v in z["routes"].items():
                rmem[rt].append(z)
                share = v / z["volume"] if z["volume"] else 1.0 / len(z["routes"])
                for d, dv in z["dsps"].items():
                    # split the ZIP's DSP volume across its routes pro-rata (+tiny weight so 0-vol ZIPs count)
                    rdsp[rt][d] += dv * share + 1e-6
        routes = []
        for rt, members in rmem.items():
            rv = sum(m["routes"][rt] for m in members)
            dd = rdsp[rt]
            routes.append(dict(route=rt, volume=rv, zip_count=len(members),
                               price=round(wavg([(m["price"], m["routes"][rt]) for m in members]), 3),
                               dominant_dsp=max(dd, key=dd.get) if dd else None))
        routes.sort(key=lambda r: r["route"])

        m = smeta.get(st, {})
        lat, lon = num(m.get("lat")), num(m.get("lon"))
        azip = re.findall(r"\b(\d{5})(?:-\d{4})?\b", m.get("address") or "")
        ag = geo.get(azip[-1]) if azip else None
        if lat is not None and lon is not None:
            src = "address"
        elif ag is not None:
            lat, lon = float(ag["properties"]["lat"]), float(ag["properties"]["lon"])
            src = "address_zip"
        else:
            pts = [(z["lat"], z["lon"], z["volume"]) for z in zl if z["lat"] is not None]
            if pts and sum(p[2] for p in pts) == 0:
                pts = [(a, b, 1) for a, b, _ in pts]
            w = sum(p[2] for p in pts)
            lat = sum(p[0] * p[2] for p in pts) / w if pts else None
            lon = sum(p[1] * p[2] for p in pts) / w if pts else None
            src = "service_area_centroid"
        dsps = sorted({d for z in zl for d in z["dsps"]})
        stations[st] = dict(
            name=st, head=m.get("head") or None, head_pinyin=m.get("head_pinyin") or None,
            address=m.get("address") or None, dsps=dsps, total_volume=tot,
            avg_price=round(wavg([(z["price"], z["volume"]) for z in zl]), 3),
            zip_count=len(zl), coord_source=src,
            lat=round(lat, 6) if lat is not None else None,
            lon=round(lon, 6) if lon is not None else None,
            zips=zl, routes=routes)

    # ---- DSP summary ----
    dsp_summary = {}
    for d in sorted({d for z in zips.values() for d in z["dsps"]}):
        members = [z for z in zips.values() if d in z["dsps"]]
        dsp_summary[d] = dict(
            volume=sum(z["dsps"][d] for z in members), zip_count=len(members),
            stations=sorted({z["station"] for z in members}),
            avg_price=round(wavg([(z["price"], z["dsps"][d]) for z in members]), 3))

    mapdata = dict(stations=stations, dsp_summary=dsp_summary)
    zipgeo = dict(type="FeatureCollection",
                  features=[geo[z] for z in sorted(zips) if z in geo])
    pricedata = [dict(z=z["zip"], s=z["station"], p=z["price"], v=z["volume"], d=z["dsps"])
                 for z in sorted(zips.values(), key=lambda z: -z["volume"])]

    # ---- inject into HTML ----
    def inject(path, tag_id, obj):
        html = path.read_text(encoding="utf-8")
        payload = json.dumps(obj, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
        pat = re.compile(r'(<script id="%s" type="application/json">).*?(</script>)' % tag_id, re.S)
        if not pat.search(html):
            sys.exit(f'Could not find <script id="{tag_id}"> in {path}')
        html = pat.sub(lambda m: m.group(1) + payload + m.group(2), html, count=1)
        path.write_text(html, encoding="utf-8")

    inject(INDEX, "mapdata", mapdata)
    inject(INDEX, "zipgeo", zipgeo)
    inject(DSP_PAGE, "pricedata", pricedata)

    # ---- price sheet baseline (warn if it overwrites manual edits) ----
    old = {}
    if PRICE_CSV.exists():
        with open(PRICE_CSV, newline="", encoding="utf-8-sig") as f:
            for r in csv.reader(f):
                if len(r) >= 2 and re.fullmatch(r"\d{5}", r[0].strip()):
                    old[r[0].strip()] = num(r[1])
    drift = [z for z in zips if z in old and old[z] is not None and abs(old[z] - zips[z]["price"]) > 1e-9]
    with open(PRICE_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["zip", "price", "station", "route"])
        for z in sorted(zips.values(), key=lambda z: (z["station"], z["zip"])):
            w.writerow([z["zip"], f'{z["price"]:.2f}', z["station"],
                        max(z["routes"], key=z["routes"].get)])

    # ---- report ----
    nogeom = sorted(z for z in zips if not zips[z]["has_geom"])
    est = [s for s in stations if stations[s]["coord_source"] != "address"]
    print(f"Built from {xlsx.name}: {len(stations)} stations, {len(dsp_summary)} DSPs, "
          f"{len(zips)} ZIPs, {sum(s['total_volume'] for s in stations.values()):,} pkgs/day")
    for s, info in stations.items():
        print(f"  {s:6} {info['zip_count']:4} ZIPs  {info['total_volume']:7,}/day  "
              f"${info['avg_price']:.3f}  {len(info['routes'])} routes  DSPs {','.join(info['dsps'])}")
    if nogeom:
        print(f"! {len(nogeom)} ZIPs have no boundary (PO box / not in 2020 ZCTA): {', '.join(nogeom)}")
    if est:
        print(f"! Estimated station pins (fill lat/lon in data/stations.csv): {', '.join(est)}")
    if drift:
        print(f"! data/price_adjustments.csv had {len(drift)} manual price edits that differ from the "
              f"workbook; they were replaced by workbook prices: {', '.join(drift[:15])}{' ...' if len(drift) > 15 else ''}")
    for p in problems:
        print("! " + p)
    print("Done. Commit and push to publish.")


if __name__ == "__main__":
    main()
