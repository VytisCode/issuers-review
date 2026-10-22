"""Turimų obligacijų kotiruotės (be AI, nemokamai): Nasdaq Baltic obligacijų sąrašas (XLSX) -> data/prices.csv.

Saugoma kiekvieną paleidimą (kas savaitę): pirkimo / pardavimo kaina (be sukaupto kupono), pajamingumas, paskutinio sandorio kaina.
Tik papildoma: ta pati diena perrašoma, senos dienos lieka (kainų istorija). Obligacijų, kurios neprekiaujamos
Nasdaq Baltic (dauguma XS bankų obligacijų), viešos nemokamos kainos nėra - jos praleidžiamos.

Naudojimas: python tools/fetch_prices.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402
from lookup_isins import LISTS, fetch, rows  # noqa: E402

OUT = C.DATA / "prices.csv"
FIELDS = ["date", "isin", "bid", "ask", "last", "bid_yield", "ask_yield", "trades", "turnover_eur", "source"]
COLS = {"bid": "bid clean", "ask": "ask clean", "last": "last trade price clean", "bid_yield": "bid yield", "ask_yield": "ask yield",
        "trades": "trades", "turnover_eur": "turnover eur"}


def num(v):
    try:
        x = float(str(v).replace(",", ".").replace(" ", "").strip())
        return f"{x:g}" if x else ""
    except ValueError:
        return ""


def main():
    held = {s["isin"] for s in C.securities() if s.get("held") == "yes" and s["type"] == "bond"}
    try:
        data = rows(fetch(LISTS["bond"]))
    except Exception as e:  # noqa: BLE001
        print("Nasdaq Baltic sąrašas nepasiekiamas:", e)
        return
    old = [r for r in C.read_csv(OUT) if r["date"] != C.TODAY]
    new = []
    for x in data:
        isin = str(x.get("isin", "")).strip()
        if isin not in held:
            continue
        r = {"date": C.TODAY, "isin": isin, "source": f"https://nasdaqbaltic.com/statistics/en/instrument/{isin}/trading"}
        r.update({k: num(x.get(c, "")) for k, c in COLS.items()})
        if any(r[k] for k in ("bid", "ask", "last")):
            new.append(r)
    C.write_csv(OUT, sorted(old + new, key=lambda r: (r["isin"], r["date"])), FIELDS)
    print(f"Kotiruotės: {len(new)} iš {len(held)} turimų obligacijų")


if __name__ == "__main__":
    main()
