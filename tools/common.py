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
    elif a:
        txt = f"{arrow} {fmt_num(f'{abs((b - a) / abs(a) * 100):.1f}', '', lang)} %"
    else:
        txt = ""
    trend, level = judge(metric, b, a)
    return last, prev, txt, trend, level


_LABEL = re.compile(r"\*\*([^*]{1,60}?):?\*\*:?\s*")
_SEMI = re.compile(r";\s+(?![^\[]*\])(?![^(]*\))")
_SENT = re.compile(r"(?<=[.!?])\s+(?![^\[]*\])(?=[A-ZĄČĘĖĮŠŲŪŽ0-9„\"(])")


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
    return rows


def history(issuer):
    """{metric: {period: row}} iš data/history/<issuer>.csv."""
    out = {}
    for r in read_csv(DATA / "history" / f"{issuer}.csv"):
        out.setdefault(r["metric"], {})[r["period"]] = r
    return out


def reports(issuer):
    """Visos emitento ataskaitos (naujausia paskutinė pagal analizės datą)."""
    items = []
    for p in sorted((REPORTS / issuer).glob("*.yaml")):
        if p.name.startswith("_"):
            continue
        d = load_yaml(p)
        d["_file"] = p
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
