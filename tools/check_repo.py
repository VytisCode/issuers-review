"""Greita repo patikra (be tinklo, be AI): ar duomenų failai sveiki, prieš kuriant ataskaitas ir svetainę.
Paleidžiama automatiškai (.github/workflows/checks.yml) ir ranka: python tools/check_repo.py
KLAIDA – sustabdo (išėjimo kodas 1); PERSPĖJIMAS – tik parodomas (pvz. žinomas negaliojantis ISIN)."""
import csv
import re
import sys
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

errors, warnings = [], []


def isin_ok(isin):
    """ISIN kontrolinis skaitmuo (Luhn)."""
    if not re.fullmatch(r"[A-Z]{2}[A-Z0-9]{9}\d", isin):
        return False
    digits = "".join(str(int(ch, 36)) for ch in isin[:-1])
    total = 0
    for i, d in enumerate(reversed(digits)):
        n = int(d) * (2 if i % 2 == 0 else 1)
        total += n - 9 if n > 9 else n
    return (10 - total % 10) % 10 == int(isin[-1])


def check_yaml():
    """Visi YAML failai skaitomi; greitas C skaitytuvas duoda tą patį rezultatą kaip lėtasis (Python)."""
    class Slow(yaml.SafeLoader):
        pass
    Slow.yaml_implicit_resolvers = C._Loader.yaml_implicit_resolvers
    for p in sorted([*C.CONFIG.glob("*.yaml"), *C.REPORTS.glob("*/*.yaml"), *(C.DATA / "esg").glob("*.yaml")]):
        try:
            fast = C.load_yaml(p)
        except yaml.YAMLError as e:
            errors.append(f"{p.relative_to(C.ROOT)}: YAML klaida: {e}")
            continue
        with open(p, encoding="utf-8") as f:
            slow = yaml.load(f, Loader=Slow) or {}
        if repr(fast) != repr(slow):
            errors.append(f"{p.relative_to(C.ROOT)}: greitas ir lėtas YAML skaitytuvas duoda skirtingą rezultatą")


def check_securities(iss):
    rows = C.read_csv(C.CONFIG / "securities.csv")
    seen = set()
    for i, r in enumerate(rows, start=2):
        isin, where = r.get("isin", "").strip(), f"config/securities.csv eil. {i}"
        if not isin:
            errors.append(f"{where}: tuščias ISIN")
            continue
        if isin in seen:
            errors.append(f"{where}: ISIN {isin} kartojasi")
        seen.add(isin)
        if not isin_ok(isin):
            warnings.append(f"{where}: {isin} – neteisingas ISIN formatas ar kontrolinis skaitmuo")
        if not r.get("issuer"):
            warnings.append(f"{where}: {isin} – nepriskirtas emitentui")
        elif r["issuer"] not in iss:
            errors.append(f"{where}: emitento „{r.get('issuer')}“ nėra config/issuers.yaml")
        for col in ("issue_date", "maturity_date", "call_date", "redeemed_date"):
            v = (r.get(col) or "").strip()
            if v and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
                errors.append(f"{where}: {col} „{v}“ – data turi būti YYYY-MM-DD")
        for col in ("amount_m", "coupon_pct"):
            v = (r.get(col) or "").strip()
            if v:
                try:
                    float(v)
                except ValueError:
                    errors.append(f"{where}: {col} „{v}“ – ne skaičius")


def check_history(iss):
    need = ["metric", "period", "value", "unit", "source", "page", "verified", "note"]
    for p in sorted((C.DATA / "history").glob("*.csv")):
        name = p.relative_to(C.ROOT)
        if p.stem not in iss:
            warnings.append(f"{name}: emitento „{p.stem}“ nėra config/issuers.yaml")
        with open(p, encoding="utf-8", newline="") as f:
            rd = csv.DictReader(f)
            if rd.fieldnames != need:
                errors.append(f"{name}: stulpeliai {rd.fieldnames}, turi būti {need}")
                continue
            for i, r in enumerate(rd, start=2):
                if None in r:
                    errors.append(f"{name} eil. {i}: per daug stulpelių (kablelis tekste be kabučių?)")
                try:
                    float(r["value"])
                except (TypeError, ValueError):
                    errors.append(f"{name} eil. {i}: reikšmė „{r['value']}“ – ne skaičius")
                if r["verified"] not in ("yes", "check", "calc", "no"):
                    warnings.append(f"{name} eil. {i}: verified „{r['verified']}“ (laukiama yes / check / calc / no)")


def check_reports(iss):
    for key in iss:
        if not list((C.REPORTS / key).glob("*.yaml")):
            warnings.append(f"reports/{key}/: nėra nė vienos ataskaitos")


def main():
    iss = C.issuers()
    check_yaml()
    check_securities(iss)
    check_history(iss)
    check_reports(iss)
    for w in warnings:
        print("PERSPĖJIMAS:", w)
    for e in errors:
        print("KLAIDA:", e)
    print(f"Patikra: {len(errors)} klaidų, {len(warnings)} perspėjimų")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
