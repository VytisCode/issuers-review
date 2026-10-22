"""Sankcijų sąrašų patikra (nemokama, be AI): ES, JT, JAV OFAC ir JK konsoliduoti sąrašai.

Kiekvieno emitento pavadinimai (config/issuers.yaml news_match) ieškomi visame sąrašo tekste.
Rezultatas - data/sanctions.json: {emitentas: {date, lists: {EU: not_found|found|error, ...}, matches: [...]}}.
Ataskaitose ESG punktas "not in the list of sanctions" užpildomas pagal šį failą (common.esg_items).
Naudojimas: python tools/check_sanctions.py
"""
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
    "UK": "https://ofsistorage.blob.core.windows.net/publishlive/2022format/ConList.csv",
}


def norm(s):
    return re.sub(r"\s+", " ", re.sub(r"[^0-9a-ząčęėįšųūž]+", " ", s.lower())).strip()


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (issuers-review sanctions check)"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read().decode("utf-8", "ignore")


def main():
    texts = {}
    for name, url in LISTS.items():
        try:
            texts[name] = " " + norm(fetch(url)) + " "
            print(f"{name}: {len(texts[name]):,} simbolių")
        except Exception as e:  # noqa: BLE001
            print(f"{name}: klaida {e}")
            texts[name] = None
    out = {}
    for key, iss in C.issuers().items():
        terms = [norm(t) for t in iss.get("news_match", [])] or [norm(iss["name"])]
        res, matches = {}, []
        for name, txt in texts.items():
            if txt is None:
                res[name] = "error"
                continue
            hit = [t for t in terms if f" {t} " in txt]
            res[name] = "found" if hit else "not_found"
            matches += [f"{name}: {t}" for t in hit]
        out[key] = {"date": C.TODAY, "lists": res, "matches": matches}
        print(key, res)
    C.save_json(C.DATA / "sanctions.json", out)


if __name__ == "__main__":
    main()
