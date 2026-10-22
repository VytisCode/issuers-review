"""Naujų emitentų ataskaitų paieška (be AI, nemokamai) -> data/filings.csv (tik papildoma).

Kiekvienam emitentui atsidaro config/issuers.yaml "report_pages" puslapius, surenka nuorodas į dokumentus
(PDF, XLSX ar nuorodas su žodžiais report / ataskaita / aruanne / pārskats / results / final terms) ir palygina su jau žinomais.
Nauji dokumentai įrašomi su first_seen data; dashboard'e emitentas pažymimas „Nauja ataskaita – reikia analizės“,
kol reports/<emitentas>/ neatsiras naujesnė analizė.
Pirmas paleidimas emitentui tik užfiksuoja esamą būklę (baseline=yes), kad senos ataskaitos nebūtų rodomos kaip naujos.

Naudojimas: python tools/check_filings.py
"""
import hashlib
import html
import re
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
OUT = C.DATA / "filings.csv"
FIELDS = ["id", "issuer", "kind", "title", "url", "page", "first_seen", "baseline"]
DOC = re.compile(r"\.(pdf|xlsx?|zip)(\?|$)", re.I)
WORDS = re.compile(r"report|results|interim|annual|quarter|presentation|fact.?sheet|final.?terms|prospectus|ataskait|rezultat|aruanne|"
                   r"tulemus|pārskat|rezultāt|\b(q[1-4]|[1-4]q|[369]m|h[12]|1h|2h)\b", re.I)
KIND = [("annual", r"annual|metin|aasta|gada"), ("interim", r"interim|quarter|q[1-4]|[1-4]q|[369]m|six months|three months|nine months|h1|1h|ketv|vahe|starp"),
        ("presentation", r"presentation|webinar|prezent|esitlus"), ("bond", r"final.?terms|prospectus|tender|emtn|bond"),
        ("factsheet", r"fact.?sheet")]


def get(url, timeout=40):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "en,lt;q=0.8"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", "replace")


def links(page_url, text):
    for m in re.finditer(r'<a\b[^>]*href=["\']([^"\'#]+)["\'][^>]*>(.*?)</a>', text, re.I | re.S):
        href = urllib.parse.urljoin(page_url, html.unescape(m.group(1)))
        label = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", m.group(2)))).strip()
        name = urllib.parse.unquote(href.rsplit("/", 1)[-1])
        if DOC.search(href) or WORDS.search(label):
            yield href, (label or name)[:200], name


def kind(s):
    for k, rx in KIND:
        if re.search(rx, s, re.I):
            return k
    return "other"


def main():
    old = C.read_csv(OUT)
    known = {r["id"] for r in old}
    have_issuer = {r["issuer"] for r in old}
    new, errors = [], []
    for key, iss in C.issuers().items():
        baseline = "yes" if key not in have_issuer else "no"
        for page in iss.get("report_pages", []):
            try:
                text = get(page)
            except Exception as e:  # noqa: BLE001
                errors.append(f"{key}: {page}: {e}")
                continue
            for href, label, name in links(page, text):
                rid = hashlib.sha1(href.encode()).hexdigest()[:12]
                if rid in known:
                    continue
                known.add(rid)
                new.append({"id": rid, "issuer": key, "kind": kind(label + " " + name), "title": label, "url": href, "page": page,
                            "first_seen": C.TODAY, "baseline": baseline})
            time.sleep(1)
    C.write_csv(OUT, old + new, FIELDS)
    status = C.load_json(C.DATA / "source_status.json")
    status["filings"] = {"date": C.TODAY, "new": sum(1 for r in new if r["baseline"] == "no"), "baseline": sum(1 for r in new if r["baseline"] == "yes"),
                         "errors": errors[:20]}
    C.save_json(C.DATA / "source_status.json", status)
    print(f"Dokumentų: +{len(new)} (iš jų pradinė būklė: {status['filings']['baseline']}); klaidų: {len(errors)}")
    for e in errors[:5]:
        print("  ", e)


if __name__ == "__main__":
    main()
