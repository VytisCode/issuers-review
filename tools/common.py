"""Bendros funkcijos: konfigūracijos, istorijos ir ataskaitų skaitymas, skaičių formatavimas."""
import csv
import json
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
    """Vertybiniai popieriai su apskaičiuota būsena: active / matured (išpirkimo data praėjo)."""
    rows = read_csv(CONFIG / "securities.csv")
    for r in rows:
        m = r.get("maturity_date") or ""
        r["status"] = "matured" if m and m < TODAY else "active"
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
