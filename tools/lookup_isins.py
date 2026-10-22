"""Naujų ISIN duomenų paieška Nasdaq Baltic sąrašuose (nemokama, be AI).

config/securities.csv eilutėms be emitento ar sąlygų (pvz., pridėtoms iš dashboard'o ISIN sąrašo) parsisiunčiami
Nasdaq Baltic obligacijų ir akcijų XLSX sąrašai ir užpildomi tušti laukai: pavadinimas, tipas, kuponas, išpirkimas,
listingas, šaltinis. Užpildyti laukai žymimi verified=check (paimta iš biržos sąrašo, ne iš emisijos dokumentų).
Esami (ne tušti) laukai niekada neperrašomi. Naudojimas: python tools/lookup_isins.py
"""
import csv
import io
import re
import sys
import urllib.request
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

LISTS = {
    "bond": "https://nasdaqbaltic.com/statistics/en/bonds?download=1",
    "share": "https://nasdaqbaltic.com/statistics/en/shares?download=1",
}
PAGE = "https://nasdaqbaltic.com/statistics/en/instrument/{isin}/trading"
MARKETS = {"VLN": "Nasdaq Vilnius", "RIG": "Nasdaq Riga", "TLN": "Nasdaq Tallinn"}


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (issuers-review)"})
    with urllib.request.urlopen(req, timeout=90) as r:
        return r.read()


def rows(raw):
    """XLSX (arba CSV) -> sąrašas dict'ų su mažosiomis raidėmis rašytais stulpelių pavadinimais."""
    if raw[:2] == b"PK":
        ws = load_workbook(io.BytesIO(raw), read_only=True, data_only=True).active
        table = [[("" if c is None else c) for c in r] for r in ws.iter_rows(values_only=True)]
    else:
        txt = raw.decode("utf-8-sig", "ignore")
        table = list(csv.reader(io.StringIO(txt), delimiter=";" if txt.count(";") > txt.count(",") else ","))
    for i, r in enumerate(table[:10]):
        head = [str(c).strip().lower() for c in r]
        if "isin" in head:
            print("Stulpeliai:", head)
            return [dict(zip(head, x)) for x in table[i + 1:] if any(str(c).strip() for c in x)]
    raise ValueError(f"nerastas ISIN stulpelis; pirmos eilutės: {table[:3]}")


def col(row, *names):
    for k, v in row.items():
        if any(n in k for n in names) and str(v).strip():
            return v
    return ""


def date(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    m = re.match(r"(\d{1,2})[./](\d{1,2})[./](\d{4})", str(v).strip())
    if m:
        return f"{m[3]}-{int(m[2]):02d}-{int(m[1]):02d}"
    return str(v).strip()[:10] if re.match(r"\d{4}-\d{2}-\d{2}", str(v).strip()) else ""


def num(v):
    s = str(v).strip().replace(",", ".")
    return s if re.fullmatch(r"\d+(\.\d+)?", s) else ""


def main():
    path = C.CONFIG / "securities.csv"
    sec = list(csv.DictReader(open(path, encoding="utf-8")))
    fields = list(sec[0].keys())
    todo = [r for r in sec if not r["issuer"] or not r["name"]]
    if not todo:
        print("Naujų ISIN nėra")
        return
    found = {}
    for typ, url in LISTS.items():
        try:
            for x in rows(fetch(url)):
                found.setdefault(str(x.get("isin", "")).strip().upper(), (typ, x))
        except Exception as e:  # noqa: BLE001
            print(f"{typ}: klaida {e}")
    names = {k: v["name"].lower() for k, v in C.issuers().items()}
    n = 0
    for r in todo:
        hit = found.get(r["isin"])
        if not hit:
            print(r["isin"], "- nerasta Nasdaq Baltic sąrašuose")
            continue
        typ, x = hit
        name = str(col(x, "name", "company")).strip()
        market = str(col(x, "home market", "market", "list")).strip()
        new = {
            "type": typ, "name": name, "currency": "EUR" if typ == "share" else str(col(x, "currency")).strip() or "EUR",
            "coupon_pct": num(col(x, "coupon")) if typ == "bond" else "",
            "maturity_date": date(col(x, "maturity")) if typ == "bond" else "",
            "issue_date": date(col(x, "issued", "issue date", "listing date")) if typ == "bond" else "",
            "listing": MARKETS.get(market.upper()[:3], market), "source_url": PAGE.format(isin=r["isin"]),
        }
        if not r["issuer"]:  # žinomas emitentas atpažįstamas pagal pavadinimą
            r["issuer"] = next((k for k, v in names.items() if v.split()[0] in name.lower()), "")
        for k, v in new.items():
            if k in r and not str(r[k]).strip() and v:
                r[k] = v
        if r.get("status_note", "").startswith("Added from the ISIN list"):
            r["status_note"] += "; terms from the Nasdaq Baltic list"
        r["verified"] = "check"
        n += 1
        print(r["isin"], "-", name, new["coupon_pct"], new["maturity_date"])
    if n:
        w = csv.DictWriter(open(path, "w", encoding="utf-8", newline=""), fieldnames=fields)
        w.writeheader()
        w.writerows(sec)
    print(f"Užpildyta: {n} iš {len(todo)}")


if __name__ == "__main__":
    main()
