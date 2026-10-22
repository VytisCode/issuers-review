"""Sankcijų sąrašų patikra (nemokama, be AI): ES, JT, JAV OFAC ir JK konsoliduoti sąrašai.

Kiekvieno emitento pavadinimai (config/issuers.yaml news_match) lyginami su sankcionuotų asmenų vardų / pavadinimų
ir slapyvardžių laukais (ne su pastabomis - ten minimi bankai, kuriuose buvo sąskaitos). Paminėjimai kituose laukuose
įrašomi atskirai (mentions) su ištrauka, bet nelaikomi sutapimu.
Rezultatas - data/sanctions.json: {emitentas: {date, lists: {EU: not_found|found|error, ...}, matches: [...], mentions: [...]}}.
Ataskaitose ESG punktas "not in the list of sanctions" užpildomas pagal šį failą (common.esg_items).
Naudojimas: python tools/check_sanctions.py
"""
import csv
import io
import json
import re
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

LISTS = {
    "EU": "https://webgate.ec.europa.eu/fsd/fsf/public/files/csvFullSanctionsList_1_1/content?token=dG9rZW4tMjAxNw",
    "UN": "https://scsanctions.un.org/resources/xml/en/consolidated.xml",
    "US OFAC": "https://www.treasury.gov/ofac/downloads/sdn.csv",
    "US OFAC alt": "https://www.treasury.gov/ofac/downloads/alt.csv",
    "UK": "https://ofsistorage.blob.core.windows.net/publishlive/2022format/ConList.csv",
}


def norm(s):
    return re.sub(r"\s+", " ", re.sub(r"[^0-9a-ząčęėįšųūž]+", " ", s.lower())).strip()


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (issuers-review sanctions check)"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read().decode("utf-8", "ignore")


def names(name, raw):
    """Grąžina sankcionuotų asmenų vardų / pavadinimų ir slapyvardžių sąrašą."""
    if name == "UN":
        return re.findall(r"<(?:FIRST|SECOND|THIRD|FOURTH)_NAME>([^<]+)<|<ALIAS_NAME>([^<]+)<", raw)
    if name.startswith("US OFAC"):  # sdn.csv: 2 stulpelis - pavadinimas; alt.csv: 4 stulpelis - slapyvardis
        col = 3 if name.endswith("alt") else 1
        return [r[col] for r in csv.reader(io.StringIO(raw)) if len(r) > col]
    lines = raw.splitlines()
    delim = ";" if lines and lines[0].count(";") > lines[0].count(",") else ","
    rows = list(csv.reader(lines, delimiter=delim))
    for i, r in enumerate(rows[:5]):
        cols = [j for j, c in enumerate(r) if re.search(r"wholename|^name ?[1-6]$|lastname|firstname", c.replace("_", "").lower())]
        if cols:
            return [" ".join(x[j] for j in cols if j < len(x)) for x in rows[i + 1:]]
    raise ValueError("nerasti vardų stulpeliai")


def snippet(txt, term):
    i = txt.find(f" {term} ")
    return txt[max(0, i - 80):i + len(term) + 80].strip()


def main():
    data = {}
    for name, url in LISTS.items():
        try:
            raw = fetch(url)
            nm = [norm(" ".join(n) if isinstance(n, tuple) else n) for n in names(name, raw)]
            data[name] = (" " + " | ".join(nm) + " ", " " + norm(raw) + " ")
            print(f"{name}: {len(nm):,} vardų")
        except Exception as e:  # noqa: BLE001
            print(f"{name}: klaida {e}")
            data[name] = None
    if data.get("US OFAC") and data.get("US OFAC alt"):
        a, b = data.pop("US OFAC"), data.pop("US OFAC alt")
        data["US OFAC"] = (a[0] + b[0], a[1] + b[1])
    else:
        data.pop("US OFAC alt", None)
    out = {}
    for key, iss in C.issuers().items():
        terms = [norm(t) for t in iss.get("news_match", [])] or [norm(iss["name"])]
        excl = [norm(x["name"]) for x in iss.get("sanctions_exclude", [])]
        res, matches, mentions, excluded = {}, [], [], set()
        for name, d in data.items():
            if d is None:
                res[name] = "error"
                continue
            nm, full = d
            for ex in excl:  # bendravardžiai - išimami prieš paiešką
                if f" {ex} " in nm:
                    excluded.add(f"{name}: {ex}")
                nm, full = nm.replace(f" {ex} ", " "), full.replace(f" {ex} ", " ")
            hit = [t for t in terms if f" {t} " in nm]
            res[name] = "found" if hit else "not_found"
            matches += [f"{name}: {snippet(nm, t)}" for t in hit]
            mentions += [f"{name}: …{snippet(full, t)}…" for t in terms if t not in hit and f" {t} " in full]
        out[key] = {"date": C.TODAY, "lists": res, "matches": matches, "mentions": mentions, "excluded": sorted(excluded)}
        print(key, res, matches, mentions)
    C.save_json(C.DATA / "sanctions.json", out)


if __name__ == "__main__":
    main()
