"""Bendros funkcijos: konfigūracijos, istorijos ir ataskaitų skaitymas, skaičių formatavimas."""
import csv
import json
import re
from datetime import date
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
CONFIG = ROOT / "config"
DATA = ROOT / "data"
REPORTS = ROOT / "reports"
DOCS = ROOT / "docs"
TODAY = date.today().isoformat()


class _Loader(yaml.SafeLoader):
    """YAML be „yes/no → True/False“ (ESG būsenos yes/no turi likti tekstu)."""


_Loader.yaml_implicit_resolvers = {k: [(tag, rx) for tag, rx in v if tag != "tag:yaml.org,2002:bool"]
                                   for k, v in yaml.SafeLoader.yaml_implicit_resolvers.items()}


def load_yaml(path):
    with open(path, encoding="utf-8") as f:
        return yaml.load(f, Loader=_Loader) or {}


def settings():
    return load_yaml(CONFIG / "settings.yaml")


def issuers():
    return load_yaml(CONFIG / "issuers.yaml")


def i18n():
    return load_yaml(CONFIG / "i18n.yaml")


def metrics():
    return load_yaml(CONFIG / "metrics.yaml")


def metric_rules():
    return load_yaml(CONFIG / "metric_rules.yaml") if (CONFIG / "metric_rules.yaml").exists() else {}


def judge(metric, value, prev=None):
    """Grąžina (trend, level): trend good/bad/flat/none pagal kryptį; level good/mid/weak/None pagal ribas."""
    r = metric_rules().get(metric, {})
    b = r.get("better", "none")
    trend = "none"
    if prev is not None and b in ("up", "down"):
        d = float(value) - float(prev)
        trend = "flat" if abs(d) < 1e-9 else ("good" if (d > 0) == (b == "up") else "bad")
    level = None
    if "good" in r and "weak" in r:
        v, g, w = float(value), float(r["good"]), float(r["weak"])
        if b == "down":
            level = "good" if v <= g else "weak" if v > w else "mid"
        else:
            level = "good" if v >= g else "weak" if v < w else "mid"
    return trend, level


def _pkey(p):
    return (p[-4:], 1 if p.startswith("FY") else 0, p)


def ratio_change(metric, vals, periods, lang):
    """Paskutinio laikotarpio pokytis ir vertinimas: (last, prev, tekstas, trend, level).
    Srautų rodikliai lyginami su tuo pačiu laikotarpiu prieš metus; būklės (stock) - su naujausiu ankstesniu."""
    have = [p for p in periods if p in vals]
    if not have:
        return None, None, "", "none", None
    last = have[-1]
    same = last[:-4] + str(int(last[-4:]) - 1)
    if same in have:
        prev = same
    elif metric_rules().get(metric, {}).get("stock"):
        earlier = [x for x in sorted(have, key=_pkey) if _pkey(x) < _pkey(last)]
        prev = earlier[-1] if earlier else None
    else:
        prev = None
    if not prev:
        return last, None, "", "none", judge(metric, vals[last]["value"])[1]
    unit, a, b = vals[last]["unit"], float(vals[prev]["value"]), float(vals[last]["value"])
    arrow = "▲" if b > a else "▼" if b < a else "•"
    if unit in ("%", "x"):
        d = abs(b - a)
        txt = f"{arrow} {fmt_num(f'{d:.2f}'.rstrip('0').rstrip('.'), 'x' if unit == 'x' else '', lang)}{' p. p.' if unit == '%' else ''}"
    elif a > 0 and abs(b - a) / a <= 5:
        txt = f"{arrow} {fmt_num(f'{abs((b - a) / abs(a) * 100):.1f}', '', lang)} %"
    elif a or b:  # neigiama ar artima nuliui bazė - procentas neprasmingas, rodomas absoliutus pokytis
        txt = f"{arrow} {fmt_num(f'{abs(b - a):.2f}'.rstrip('0').rstrip('.'), '', lang)} {unit}"
    else:
        txt = ""
    trend, level = judge(metric, b, a)
    return last, prev, txt, trend, level


_LABEL = re.compile(r"\*\*([^*]{1,60}?):?\*\*:?\s*")
_SEMI = re.compile(r";\s+(?![^\[]*\])(?![^(]*\))")
# sakinio pabaiga, bet ne po santrumpų (mln. EUR, 2025 m. Grupė, žr. ...), kad punktas nesuskiltų per vidurį
_SENT = re.compile(r"(?<=[.!?])(?<!mln\.)(?<!mlrd\.)(?<!tūkst\.)(?<!proc\.)(?<!žr\.)(?<!pvz\.)(?<!\sm\.)(?<!\sp\.)(?<!\sd\.)(?<!\bbn\.)"
                   r"\s+(?![^\[]*\])(?=[A-ZĄČĘĖĮŠŲŪŽ0-9„\"(])")


def points(text):
    """Ilgą įrodymo tekstą suskaido į struktūrą: [(antraštė arba None, [punktai])].
    Antraštės - **Žodis:** žymės; punktai - sakiniai ir ; atskirtos dalys (išnašų [..] viduje neskaidoma)."""
    text = (text or "").strip()
    if not text:
        return []
    parts, pos, label = [], 0, None
    for m in _LABEL.finditer(text):
        chunk = text[pos:m.start()].strip()
        if chunk or label:
            parts.append((label, chunk))
        label, pos = m.group(1).strip(), m.end()
    parts.append((label, text[pos:].strip()))
    out = []
    for lab, chunk in parts:
        items = [y.strip() for x in _SEMI.split(chunk) for y in _SENT.split(x) if y.strip()]
        items = [x[0].upper() + x[1:] if x[0].islower() else x for x in items]
        if lab or items:
            out.append((lab, items))
    return out


def esg_items(rep):
    """ESG punktai iš ataskaitos; sankcijų punktas užpildomas pagal automatinę patikrą (data/sanctions.json), jei ji yra."""
    items = {k: dict(v) for k, v in rep.get("esg", {}).get("items", {}).items()}
    s = load_json(DATA / "sanctions.json").get(rep.get("issuer"), {})
    lists = {k: v for k, v in s.get("lists", {}).items() if v != "error"}
    if lists:
        names = ", ".join(lists)
        if any(v == "found" for v in lists.values()):
            hit = "; ".join(s.get("matches", []))
            items["sanctions"] = {"status": "review", "evidence": {
                "en": f"**Possible match – review:** {hit} (automatic check {s['date']}).",
                "lt": f"**Galimas atitikmuo – peržiūrėti:** {hit} (automatinė patikra {s['date']})."}}
        else:
            ex = [x for x in issuers().get(rep.get("issuer"), {}).get("sanctions_exclude", [])
                  if any(e.endswith(": " + re.sub(r"\s+", " ", x["name"].lower())) for e in s.get("excluded", []))]
            nb = {lg: "".join(f" Namesake on the lists: {x.get('en', x['name'])}." if lg == "en" else f" Bendravardis sąrašuose: {x.get('lt', x['name'])}." for x in ex) for lg in ("en", "lt")}
            items["sanctions"] = {"status": "yes", "evidence": {
                "en": f"Not found on the {names} consolidated sanctions lists (automatic check {s['date']}).{nb['en']}",
                "lt": f"Nerasta {names} konsoliduotuose sankcijų sąrašuose (automatinė patikra {s['date']}).{nb['lt']}"}}
    if "sanctions" in items:
        ev = items["sanctions"].get("evidence", "")
        ev = ev if isinstance(ev, dict) else {"en": ev, "lt": ev}
        items["sanctions"]["evidence"] = {k: (ev.get(k) or ev.get("en", "")).rstrip() + SANCTION_NOTE.get(k, SANCTION_NOTE["en"]) for k in set(ev) | {"en", "lt"}}
    # žiniasklaidos punktas papildomas automatine naujienų suvestine (ta pati, kaip skiltyje „Ką rašo spauda“)
    dig, start, end = news_digest(rep.get("issuer"), days=90)
    if "media" in items:
        n, flagged = sum(1 + d["more"] for d in dig), [d for d in dig if d["topic"] == "review"]
        fl = {"en": "", "lt": ""}
        if flagged:
            fl = {"en": " Headlines with an adverse keyword: " + "; ".join(f"„{d['title']}“ ({d['source']}, {d['date']})" for d in flagged[:3]) + ".",
                  "lt": " Antraštės su neigiamu žodžiu: " + "; ".join(f"„{d['title']}“ ({d['source']}, {d['date']})" for d in flagged[:3]) + "."}
        auto = {"en": f" News summary {start} – {end}: {n} articles, {len(flagged)} with an adverse keyword in the headline (details in the section \"In the press\").{fl['en']}",
                "lt": f" Naujienų suvestinė {start} – {end}: {n} pranešimų, iš jų su neigiamu žodžiu antraštėje – {len(flagged)} (plačiau – skiltyje „Ką rašo spauda“).{fl['lt']}"}
        ev = items["media"].get("evidence", "")
        ev = ev if isinstance(ev, dict) else {"en": ev, "lt": ev}
        items["media"]["evidence"] = {k: (ev.get(k) or ev.get("en", "")).rstrip() + auto.get(k, auto["en"]) for k in set(ev) | {"en", "lt"}}
    return items


def seb():
    return load_yaml(CONFIG / "seb_policies.yaml")


def read_csv(path):
    if not Path(path).exists():
        return []
    with open(path, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def write_csv(path, rows, fields):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def securities():
    """Vertybiniai popieriai su apskaičiuota būsena: active / matured (išpirkimo ar ankstyvo išpirkimo - redeemed_date - data praėjo)."""
    rows = read_csv(CONFIG / "securities.csv")
    for r in rows:
        m = r.get("redeemed_date") or r.get("maturity_date") or ""
        r["status"] = "matured" if m and m <= TODAY else "active"
        # žalioji obligacija: stulpelis green (yes/no) – rankinis sprendimas; tuščias – pagal pavadinimą (green / EuGB, kaip skelbia emitentas)
        g = (r.get("green") or "").strip().lower()
        r["green"] = g if g in ("yes", "no") else ("yes" if r.get("type") == "bond" and GREEN_RX.search(r.get("name", "")) else "no")
    return rows


GREEN_RX = re.compile(r"\bgreen\b|\bEuGB\b|žalio|zaļ|roheli", re.I)


def period_end(period, fy_end="12-31"):
    """Laikotarpio pabaigos data: 6M2026 -> 2026-06-30; FY2026 -> 2026-12-31. fy_end – finansinių metų pabaiga (LATRAPS: 06-30)."""
    m = re.match(r"(FY|12M|Q4|9M|Q3|3Q|6M|1H|H1|Q2|2Q|3M|Q1|1Q)(\d{4})$", str(period))
    if not m:
        return ""
    months = {"FY": 12, "12M": 12, "Q4": 12, "9M": 9, "Q3": 9, "3Q": 9, "6M": 6, "1H": 6, "H1": 6, "Q2": 6, "2Q": 6, "3M": 3, "Q1": 3, "1Q": 3}[m.group(1)]
    y, fm = int(m.group(2)), int(fy_end[:2])
    mon = (fm - 12 + months - 1) % 12 + 1  # finansinių metų mėnuo -> kalendorinis
    if mon > fm:
        y -= 1
    last = {2: 28, 4: 30, 6: 30, 9: 30, 11: 30}.get(mon, 31)
    return f"{y}-{mon:02d}-{last}"


def quarter_label(end):
    """2026-09-30 -> 9M 2026 (Q3)."""
    y, mo = end[:4], int(end[5:7])
    return {3: f"Q1 {y}", 6: f"6M {y} (Q2)", 9: f"9M {y} (Q3)", 12: f"FY {y}"}.get(mo, end)


_EXP_MONTHS = {m: i for i, m in enumerate(["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"], 1)}


def expected_date(text, after=""):
    """Numatoma kitos ataskaitos data iš teksto: pirma YYYY-MM-DD (ne ankstesnė už after), kitaip „November 2026“ -> mėnesio pabaiga (orientacinė)."""
    text = plain(str(text or ""))
    for d in re.findall(r"\d{4}-\d{2}-\d{2}", text):
        if d >= after:
            return d, True
    m = re.search(r"(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{4})", text, re.I)
    if m:
        mon = _EXP_MONTHS[m.group(1).lower()]
        last = {2: 28, 4: 30, 6: 30, 9: 30, 11: 30}.get(mon, 31)
        return f"{m.group(2)}-{mon:02d}-{last}", False
    return "", False


# Senesnė formuluotė istorijos pastabose (CSV tik papildomas, todėl keičiama rodant)
CALC_WORDING = [("apskaičiuota įrankio", "mūsų skaičiavimas"), ("Apskaičiuota įrankio", "Mūsų skaičiavimas"),
                ("calculated by the tool", "own calculation"), ("Calculated by the tool", "Own calculation")]


def calc_wording(text):
    for a, b in CALC_WORDING:
        text = text.replace(a, b)
    return text


def history(issuer):
    """{metric: {period: row}} iš data/history/<issuer>.csv."""
    out = {}
    for r in read_csv(DATA / "history" / f"{issuer}.csv"):
        if r.get("note"):
            r["note"] = calc_wording(r["note"])
        out.setdefault(r["metric"], {})[r["period"]] = r
    return out


# rankinės sankcijų patikros nuorodos (savininko prašymas 2026-10-23) - pridedamos prie kiekvienos ataskaitos;
# SEB vidinė eNameChecker nuoroda viešoje ataskaitoje neskelbiama (tik svetainėje, naršyklės atmintyje, ir DD faile)
SANCTION_SRC = {
    "SX1": {"title": "OFAC Sanctions List Search (US Treasury)", "url": "https://sanctionssearch.ofac.treas.gov/"},
    "SX2": {"title": "EU Sanctions Map", "url": "https://www.sanctionsmap.eu/#/main"},
}
SANCTION_NOTE = {"en": " **Manual check:** OFAC sanctions search [SX1]; EU Sanctions Map [SX2]; SEB eNameChecker (internal SEB system).",
                 "lt": " **Rankinė patikra:** OFAC sankcijų paieška [SX1]; ES sankcijų žemėlapis [SX2]; SEB eNameChecker (vidinė SEB sistema)."}


def reports(issuer):
    """Visos emitento ataskaitos (naujausia paskutinė pagal analizės datą)."""
    items = []
    for p in sorted((REPORTS / issuer).glob("*.yaml")):
        if p.name.startswith("_"):
            continue
        d = load_yaml(p)
        d["_file"] = p
        d["sources"] = {**(d.get("sources") or {}), **SANCTION_SRC}
        items.append(d)
    items.sort(key=lambda d: str(d.get("analysis_date", "")))
    return items


def tr(value, lang):
    """Tekstas norima kalba; jei vertimo nėra - lietuviškas, tada angliškas."""
    if isinstance(value, dict):
        return value.get(lang) or value.get("lt") or value.get("en") or ""
    return "" if value is None else str(value)


def fmt_num(value, unit, lang):
    if value in (None, ""):
        return "–"
    try:
        v = float(value)
    except ValueError:
        return str(value)
    decimals = len(str(value).split(".")[1]) if "." in str(value) else 0
    s = f"{v:,.{decimals}f}"
    if lang != "en":
        s = s.replace(",", " ").replace(".", ",")
    if unit == "%":
        return s + " %" if lang != "en" else s + "%"
    if unit == "x":
        return s + "x"
    return s


def unit_label(unit, lang):
    return {"EURm": {"lt": "mln. EUR", "en": "EUR m", "lv": "milj. EUR", "et": "mln eurot"},
            "EUR": {"lt": "EUR", "en": "EUR", "lv": "EUR", "et": "eurot"},
            "GW": {"lt": "GW", "en": "GW", "lv": "GW", "et": "GW"},
            "Mt": {"lt": "mln. t", "en": "m t", "lv": "milj. t", "et": "mln t"},
            "m": {"lt": "mln.", "en": "m", "lv": "milj.", "et": "mln"}}.get(unit, {}).get(lang, unit if unit not in ("%", "x") else "")


def load_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, ValueError):
        return {}


def save_json(path, obj):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


# ---------- šaltinių nuorodos (išnašos) ----------
import re as _re  # noqa: E402

CITE = _re.compile(r"\s*\[([^\[\]]+)\]")
PART = _re.compile(r"^\s*([A-Z]{1,2}\d{1,2})(?:\s*p\.\s*(\d+(?:-\d+)?(?:,\s*p\.\s*\d+(?:-\d+)?)*))?\s*(.*?)\s*$")


class Citations:
    """Išnašų registras vienam dokumentui: [S1 p.24; S3] -> numeriai 1, 2 su tiesiogine nuoroda į šaltinį (PDF - į puslapį)."""

    def __init__(self, sources):
        self.sources = sources or {}
        self.order = []  # [(key, page, note)]
        self.index = {}

    def add(self, key, page="", note=""):
        k = (key, page.replace("p.", "").replace(" ", "").replace(",", ", "), note)
        if k not in self.index:
            self.order.append(k)
            self.index[k] = len(self.order)
        return self.index[k]

    def url(self, key, page=""):
        u = self.sources.get(key, {}).get("url", "")
        first = _re.match(r"\d+", page or "")
        if u and first and ".pdf" in u.lower():
            u += f"#page={first.group(0)}"
        return u

    def parse(self, text):
        """Tekstas -> segmentai: ("t", tekstas, paryškinta) ir ("c", numeris, url). **x** - paryškinimas."""
        out = []
        pos = 0
        for m in CITE.finditer(text):
            parts = [PART.match(p) for p in m.group(1).split(";")]
            if not parts or not all(p and p.group(1) in self.sources for p in parts):
                continue
            out += _bold(text[pos:m.start()])
            for p in parts:
                n = self.add(p.group(1), p.group(2) or "", (p.group(3) or "").strip(" –-"))
                out.append(("c", n, self.url(p.group(1), p.group(2) or "")))
            pos = m.end()
        out += _bold(text[pos:])
        return out

    def notes(self):
        """[(numeris, šaltinio pavadinimas, puslapis, pastaba, url)]"""
        return [(i + 1, self.sources[k]["title"], pg, note, self.url(k, pg), k, str(self.sources[k].get("date", "")))
                for i, (k, pg, note) in enumerate(self.order)]


def _bold(text):
    segs = []
    for i, chunk in enumerate(_re.split(r"\*\*", text)):
        if chunk:
            segs.append(("t", chunk, i % 2 == 1))
    return segs


def lead_bold(segs, limit=60):
    """Punkto pradžia iki pirmo dvitaškio (pvz. „Skola stabili:“) paryškinama, jei trumpa."""
    if segs and segs[0][0] == "t" and not segs[0][2]:
        t = segs[0][1]
        i = t.find(": ")
        if 0 < i <= limit:
            return [("t", t[:i + 1], True), ("t", t[i + 1:], False)] + segs[1:]
    return segs


def plain(text):
    """Tekstas be nuorodų ir paryškinimų (kortelėms, Excel)."""
    return CITE.sub("", text or "").replace("**", "")


NEWS_ORDER = ["review", "reitingas", "sandoris", "obligacijos", "dividendai", "vadovai", "ataskaita", "kita"]


def _words(title):
    return {w for w in _re.findall(r"\w{4,}", title.lower())}


def news_digest(issuer=None, days=90, end=None, lang="en"):
    """Naujienos be pasikartojimų: tos pačios temos pranešimai per kelias dienas sujungiami į vieną (rodomas vienas, kiti - „dar N“).
    Tema: review (antraštėje neigiamas žodis) > reitingas > sandoris > obligacijos > dividendai > ataskaita > kita.
    Grąžina [{issuer, topic, date, title, url, source, lang, more, flag}] svarbiausi pirma."""
    import datetime as _dt
    end = end or TODAY
    start = (_dt.date.fromisoformat(end) - _dt.timedelta(days=days)).isoformat()
    rows = [r for r in read_csv(DATA / "news.csv") if (issuer is None or r["issuer"] == issuer) and start <= r["date"] <= end]
    clusters = []
    for r in sorted(rows, key=lambda r: r["date"], reverse=True):
        topics = [t for t in r.get("topics", "").split(";") if t]
        topic = "review" if r.get("flag") == "review" else min(topics, key=NEWS_ORDER.index) if topics else "kita"
        d, w = _dt.date.fromisoformat(r["date"]), _words(r["title"])
        for c in clusters:
            if c["issuer"] != r["issuer"] or c["topic"] != topic:
                continue
            near = any(abs((d - x).days) <= 4 for x in c["_dates"])
            similar = any(len(w & y) / max(1, len(w | y)) >= 0.3 for y in c["_words"])
            if near and (topic != "kita" or similar):
                c["items"].append(r)
                c["_dates"].append(d)
                c["_words"].append(w)
                break
        else:
            clusters.append({"issuer": r["issuer"], "topic": topic, "items": [r], "_dates": [d], "_words": [w]})
    out = []
    for c in clusters:
        items = sorted(c["items"], key=lambda r: r["date"], reverse=True)
        best = sorted(items, key=lambda r: {lang: 0, "en": 1}.get(r["lang"], 2))[0]  # skaitytojo kalba, tada anglų, naujausia
        out.append({"id": best.get("id", ""), "issuer": c["issuer"], "topic": c["topic"], "date": items[0]["date"], "title": best["title"], "url": best["url"],
                    "source": best["source"], "lang": best["lang"], "more": len(items) - 1, "flag": best["flag"], "keywords": best.get("keywords", "")})
    out.sort(key=lambda x: x["date"], reverse=True)
    out.sort(key=lambda x: NEWS_ORDER.index(x["topic"]))
    return out, start, end


MARKET_SERIES = ["ecb_dfr", "euribor_3m", "euribor_6m", "euribor_12m", "ea_10y", "ea_hicp"]


def market_table():
    """Rinkos rodikliai iš data/market.csv: [(serija, paskutinis mėnuo, dabar, prieš 3 mėn., prieš 12 mėn., šaltinis)]."""
    rows = read_csv(DATA / "market.csv")
    out = []
    for sid in MARKET_SERIES:
        vals = {r["period"]: r for r in rows if r["series"] == sid}
        if not vals:
            continue
        last = max(vals)
        y, m = int(last[:4]), int(last[5:7])
        back = lambda k: f"{y + (m - 1 - k) // 12:04d}-{(m - 1 - k) % 12 + 1:02d}"  # noqa: E731
        get = lambda p: float(vals[p]["value"]) if p in vals else None  # noqa: E731
        out.append((sid, last, get(last), get(back(3)), get(back(12)), vals[last]["source"]))
    return out
