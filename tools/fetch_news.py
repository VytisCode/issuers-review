"""Naujienos apie emitentus (be AI, nemokamai) -> data/news.csv (tik papildoma, istorija neperrašoma).

Pritaikyta iš pension-fund-tracker tools/fetch_news.py: Google News paieškos RSS kiekvienam emitentui jo šalies
kalba(-omis) pagal config/issuers.yaml "media" (LT -> lt, LV -> lv, EE -> et, all -> lt, lv, et ir en).
Kiekviena antraštė patikrinama neigiamų žodžių sąrašu (tyrimas, bauda, teismas, sankcijos, korupcija...):
jei randama - pažymima "review" (reikia peržiūrėti), kitaip "info". Ataskaitų / reitingų naujienos žymimos temomis.

Naudojimas: python tools/fetch_news.py            - paskutinių 30 d. naujienos
            python tools/fetch_news.py --days 365 - ilgesnė istorija (vienkartinis užpildymas)
"""
import hashlib
import html
import re
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

UA = "Mozilla/5.0 (issuers-review; public issuer news monitor)"
OUT = C.DATA / "news.csv"
FIELDS = ["id", "issuer", "date", "lang", "source", "title", "url", "flag", "keywords", "topics", "first_seen"]
LANGS = {"LT": ["lt"], "LV": ["lv"], "EE": ["et"], "all": ["lt", "lv", "et", "en"]}
GL = {"lt": ("lt", "LT", "LT:lt"), "lv": ("lv", "LV", "LV:lv"), "et": ("et", "EE", "EE:et"), "en": ("en-GB", "GB", "GB:en")}

# Neigiami žodžiai (žodžių pradžios) visomis kalbomis - radus antraštėje, naujiena žymima "review".
ADVERSE = {
    "lt": ["tyrim", "baud", "skandal", "teism", "korupc", "kyš", "sukčiav", "sankcij", "bankrot", "nuostol", "pažeid", "įtar", "prokur", "stt", "fntt", "ieškin", "kreml", "gazprom", "uokik"],
    "lv": ["izmeklē", "sod", "skandāl", "tiesa", "tiesā", "korupc", "kukuļ", "krāpšan", "sankcij", "bankrot", "zaudējum", "pārkāp", "prokurat", "knab", "prasīb", "kreml", "gazprom", "uokik"],
    "et": ["uurim", "trahv", "skandaal", "kohtuasi", "kohtusse", "kohtuotsus", "kohtuvaid", "kohus", "korrupt", "altkäemaks", "pettus", "sanktsioon", "pankrot", "kahjum", "rikkum", "prokurat", "kahtlust", "hagi", "kreml", "gazprom", "ettekirjut", "uokik"],
    "en": ["investigat", "fine", "fined", "scandal", "lawsuit", "court", "corrupt", "brib", "fraud", "sanction", "bankrupt", "breach", "probe",
           "downgrad", "money laundering", "default", "penalt", "allegation", "kremlin", "gazprom", "precept", "enforcement", "uokik"],
}
TOPICS = {
    "ataskaita": ["report", "results", "ataskait", "rezultat", "aruanne", "tulemus", "pārskat", "rezultāt", "q1", "q2", "q3", "q4", "6m", "9m", "pusmeč", "ketvir", "profit", "pelnas", "pelną", "pelno", "kasum", "peļņ"],
    "reitingas": ["rating", "moody", "s&p", "fitch", "reiting", "reitingu"],
    "obligacijos": ["bond", "obligac", "võlakir", "obligāc", "notes", "tender", "emtn"],
    "dividendai": ["dividend"],
    "vadovai": ["ceo", "cfo", "management board", "supervisory board", "valdyb", "vadov", "juhatus", "nõukogu", "juht", "valde", "padome", "vadītāj"],
    "sandoris": ["acqui", "merger", "sale of", "įsigij", "pardav", "omandam", "iegād", "otp"],
}


def get(url, timeout=40):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def gn_url(query, lang, days):
    hl, gl, ceid = GL[lang]
    q = urllib.parse.quote(f"{query} when:{days}d")
    return f"https://news.google.com/rss/search?q={q}&hl={hl}&gl={gl}&ceid={ceid}"


def clean(s):
    s = html.unescape(re.sub(r"<[^>]+>", " ", s or ""))
    return re.sub(r"\s+", " ", s).strip()


def parse_rss(raw):
    root = ET.fromstring(raw)
    for it in root.iter("item"):
        title = clean(it.findtext("title"))
        src = clean(it.findtext("source")) if it.find("source") is not None else ""
        if src and title.endswith(" - " + src):
            title = title[: -len(src) - 3]
        try:
            d = parsedate_to_datetime(it.findtext("pubDate")).date().isoformat()
        except Exception:
            d = ""
        yield {"title": title, "url": (it.findtext("link") or "").strip(), "date": d, "source": src}


def classify(title):
    low = title.lower()
    hits = sorted({w for words in ADVERSE.values() for w in words if re.search(r"(^|\W)" + re.escape(w), low)})
    topics = [t for t, words in TOPICS.items() if any(w in low for w in words)]
    return ("review" if hits else "info"), ";".join(hits), ";".join(topics)


def main():
    days = 30
    if "--days" in sys.argv:
        days = int(sys.argv[sys.argv.index("--days") + 1])
    old = C.read_csv(OUT)
    seen = {r["id"] for r in old}
    titles = {(r["issuer"], r["title"].lower()) for r in old}
    new, errors = [], []
    for key, iss in C.issuers().items():
        names = [n.lower() for n in iss.get("news_match", [iss["name"]])]
        exclude = [n.lower() for n in iss.get("news_exclude", [])]
        for lang in LANGS.get(iss.get("media", "all"), ["en"]):
            for q in iss.get("news_queries", [f'"{iss["name"]}"']):
                try:
                    items = list(parse_rss(get(gn_url(q, lang, days))))
                except Exception as e:  # noqa: BLE001
                    errors.append(f"{key}/{lang}: {e}")
                    continue
                for it in items:
                    low = it["title"].lower()
                    # antraštėje turi būti emitento vardas (Google kartais grąžina nesusijusius)
                    if not any(n in low for n in names) or any(x in low for x in exclude):
                        continue
                    rid = hashlib.sha1((key + it["title"].lower()).encode()).hexdigest()[:12]
                    if rid in seen or (key, low) in titles:
                        continue
                    flag, kw, topics = classify(it["title"])
                    new.append({"id": rid, "issuer": key, "date": it["date"], "lang": lang, "source": it["source"], "title": it["title"],
                                "url": it["url"], "flag": flag, "keywords": kw, "topics": topics, "first_seen": C.TODAY})
                    seen.add(rid)
                    titles.add((key, low))
                time.sleep(1)
    excl = {k: [n.lower() for n in v.get("news_exclude", [])] for k, v in C.issuers().items()}
    old = [r for r in old if not any(x in r["title"].lower() for x in excl.get(r["issuer"], []))]
    for r in old:  # žodžių sąrašai tikslinami - senos antraštės perklasifikuojamos tomis pačiomis taisyklėmis
        r["flag"], r["keywords"], r["topics"] = classify(r["title"])
    rows = sorted(old + new, key=lambda r: (r["date"], r["id"]), reverse=True)
    C.write_csv(OUT, rows, FIELDS)
    status = C.load_json(C.DATA / "source_status.json")
    status["news"] = {"date": C.TODAY, "new": len(new), "errors": errors[:20]}
    C.save_json(C.DATA / "source_status.json", status)
    print(f"Naujienų: +{len(new)} (iš viso {len(rows)}); klaidų: {len(errors)}")
    for e in errors[:5]:
        print("  ", e)


if __name__ == "__main__":
    main()
