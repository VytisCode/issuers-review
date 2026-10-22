"""Rinkos aplinka (be AI, nemokamai): ECB palūkanos, Euribor, euro zonos obligacijų pajamingumas ir infliacija -> data/market.csv.

Šaltinis - ECB Data Portal API (vieši duomenys, be rakto). Kiekvienam rodikliui saugomos mėnesio reikšmės
(paskutinė mėnesio stebėsena) už ~3 metus; failas perrašomas iš to paties šaltinio (istorija ECB pusėje nekinta).
Ataskaitose rodoma skiltyje „Rinka ir aplinka“: dabar, prieš 3 ir 12 mėn.

Naudojimas: python tools/fetch_market.py
"""
import csv
import io
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

API = "https://data-api.ecb.europa.eu/service/data/{flow}/{key}?format=csvdata&startPeriod={start}"
SERIES = {  # id: (flow, key)
    "ecb_dfr": ("FM", "B.U2.EUR.4F.KR.DFR.LEV"),
    "euribor_3m": ("FM", "M.U2.EUR.RT.MM.EURIBOR3MD_.HSTA"),
    "euribor_6m": ("FM", "M.U2.EUR.RT.MM.EURIBOR6MD_.HSTA"),
    "euribor_12m": ("FM", "M.U2.EUR.RT.MM.EURIBOR1YD_.HSTA"),
    "ea_10y": ("YC", "B.U2.EUR.4F.G_N_A.SV_C_YM.SR_10Y"),
    "ea_hicp": ("ICP", "M.U2.N.000000.4.ANR"),
}
OUT = C.DATA / "market.csv"
FIELDS = ["series", "period", "value", "source"]


def get(flow, key, start):
    url = API.format(flow=flow, key=key, start=start)
    req = urllib.request.Request(url, headers={"User-Agent": "issuers-review (public data)", "Accept": "text/csv"})
    with urllib.request.urlopen(req, timeout=60) as r:
        rows = list(csv.DictReader(io.StringIO(r.read().decode("utf-8"))))
    page = f"https://data.ecb.europa.eu/data/datasets/{flow}/{flow}.{key}"
    return [(x["TIME_PERIOD"], x["OBS_VALUE"]) for x in rows if x.get("OBS_VALUE")], page


def main():
    import datetime as dt
    start = (dt.date.today().replace(day=1) - dt.timedelta(days=3 * 366)).isoformat()
    old = {(r["series"], r["period"]): r for r in C.read_csv(OUT)}
    errors = []
    for sid, (flow, key) in SERIES.items():
        try:
            obs, page = get(flow, key, start)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{sid}: {e}")
            continue
        monthly = {}
        for period, v in obs:  # dienos duomenys -> paskutinė mėnesio reikšmė
            monthly[period[:7]] = (period, v)
        for m, (period, v) in monthly.items():
            old[(sid, m)] = {"series": sid, "period": m, "value": v, "source": page}
    rows = sorted(old.values(), key=lambda r: (r["series"], r["period"]))
    C.write_csv(OUT, rows, FIELDS)
    status = C.load_json(C.DATA / "source_status.json")
    status["market"] = {"date": C.TODAY, "errors": errors}
    C.save_json(C.DATA / "source_status.json", status)
    print(f"Rinkos duomenys: {len(rows)} eilučių; klaidų: {len(errors)}")
    for e in errors:
        print("  ", e)


if __name__ == "__main__":
    main()
