"""ISIN sąrašo sinchronizavimas iš dashboard'o saugyklos (Google Apps Script, config/settings.yaml isin_store_url).

Sąraše esantys ISIN -> config/securities.csv held=yes; sąraše nesantys -> held=no (eilutės netrinamos - istorija).
Nauji ISIN pridedami kaip nauji įrašai (emitentas ir sąlygos - nustatomi analizės metu).
Pakeitimai registruojami data/isin_changes.csv. Naudojimas: python tools/sync_isins.py
"""
import csv
import json
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402


def main():
    url = (C.settings().get("isin_store_url") or "").strip()
    if not url:
        print("isin_store_url nenurodytas - praleidžiama")
        return
    with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "issuers-review"}), timeout=60) as r:
        data = json.loads(r.read().decode("utf-8"))
    want = {x["isin"]: x for x in data.get("isins", [])}
    path = C.CONFIG / "securities.csv"
    rows = list(csv.DictReader(open(path, encoding="utf-8")))
    fields = list(rows[0].keys())
    changes, known = [], set()
    for r in rows:
        known.add(r["isin"])
        held = "yes" if r["isin"] in want else "no"
        if r["held"] != held:
            changes.append([C.TODAY, r["isin"], r.get("issuer", ""), "held" if held == "yes" else "removed"])
            r["held"] = held
    for isin, x in want.items():
        if isin not in known:
            row = {k: "" for k in fields}
            row.update(isin=isin, held="yes", name=x.get("note", ""), status_note=f"Added from the ISIN list {C.TODAY}", verified="review")
            rows.append(row)
            changes.append([C.TODAY, isin, "", "new"])
    if not changes:
        print("Pakeitimų nėra")
        return
    w = csv.DictWriter(open(path, "w", encoding="utf-8", newline=""), fieldnames=fields)
    w.writeheader()
    w.writerows(rows)
    log = C.DATA / "isin_changes.csv"
    new = not log.exists()
    with open(log, "a", encoding="utf-8", newline="") as f:
        cw = csv.writer(f)
        if new:
            cw.writerow(["date", "isin", "issuer", "change"])
        cw.writerows(changes)
    print(f"{len(changes)} pakeitimai:", changes)


if __name__ == "__main__":
    main()
