"""Obligacijų fondų duration – automatinis mėnesio surinkimas (be AI, nemokama).

Kas mėnesį 19 d. (.github/workflows/duration.yml): kiekvienam fondui iš config/duration_funds.json atsisiunčia
faktų lapą pagal išsaugotas nuorodas, ištraukia duration reikšmę ir dokumento datą, patikrina ISIN ir datą.
Rezultatas – data/duration/auto_<mėnesio pabaiga>.json (dashboard'e „Automatic collection“), PDF kopijos – šaka fund-docs.
Būsenos: found (data = mėnesio pabaiga ar paskutinė darbo diena), old (dokumentas senesnis), not_found (neatsisiuntė /
nerado skaičiaus), manual (viešos nuorodos nėra). Reikšmės netikrinamos AI – jas peržiūri savininkas arba Claude paprašius.

  python tools/duration_auto.py out 2026-11-30            # out = fund-docs šakos katalogas
  python tools/duration_auto.py --local <pdf katalogas> 2026-10-31   # bandymas su turimais PDF (<ISIN>_*.pdf)
"""
import calendar
import hashlib
import json
import re
import subprocess
import sys
import time
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CFG = ROOT / "config" / "duration_funds.json"
OUTD = ROOT / "data" / "duration"
MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_name) if m} | {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}
MONTHS |= {"januar": 1, "februar": 2, "märz": 3, "maerz": 3, "mai": 5, "juni": 6, "juli": 7, "oktober": 10, "dezember": 12}
NUM = re.compile(r"(?<![\d.,])(-?\d{1,2}[.,]\d{1,3})(?![\d.,]*\d{4})")


def prev_month_end(today=None):
    t = today or date.today()
    return (t.replace(day=1) - timedelta(days=1)).isoformat()


def fill(url, me):
    return url.replace("{YYYYMMDD}", me.replace("-", "")).replace("{YYYYMM}", me[:7].replace("-", ""))


def parse_date(s):
    """'31.10.2026' / '31/10/2026' / '31 October 2026' / '31OCTOBER2026' / 'August 2026' (-> month end) -> ISO."""
    s = s.strip()
    m = re.fullmatch(r"(\d{1,2})[./](\d{1,2})[./](\d{4})", s)
    if m:
        return date(int(m[3]), int(m[2]), int(m[1])).isoformat()
    m = re.fullmatch(r"(\d{1,2})\s*([A-Za-zäÄ]+)\s*(\d{4})", s)
    if m and m[2].lower() in MONTHS:
        return date(int(m[3]), MONTHS[m[2].lower()], int(m[1])).isoformat()
    m = re.fullmatch(r"([A-Za-zäÄ]+)\s+(\d{4})", s)
    if m and m[1].lower() in MONTHS:
        y, mo = int(m[2]), MONTHS[m[1].lower()]
        return date(y, mo, calendar.monthrange(y, mo)[1]).isoformat()
    return ""


def last_business_day(me):
    d = date.fromisoformat(me)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d.isoformat()


def pages(pdf):
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout or "Pages: 0").group(1))
    return [subprocess.run(["pdftotext", "-layout", "-f", str(p), "-l", str(p), str(pdf), "-"], capture_output=True, text=True).stdout
            for p in range(1, n + 1)]


def extract(fund, pg):
    """-> (value, page, as_of, isin_ok). First number after the label on the same line (fund column comes first)."""
    val = page = None
    rx = re.compile(fund["value"])
    for i, t in enumerate(pg, 1):
        for line in t.splitlines():
            m = rx.search(line)
            if m:
                n = NUM.search(line[m.end():])
                if n:
                    val, page = float(n[1].replace(",", ".")), i
                    break
        if val is not None:
            break
    as_of = ""
    for t in pg:
        src = re.sub(r"\s+", "", t.upper()) if fund.get("date_collapse") else t
        m = re.search(fund["date"], src)
        if m and parse_date(m[1]):
            as_of = parse_date(m[1])
            break
    allt = "\n".join(pg)
    isin_ok = fund["isin"] in allt or bool(fund.get("class_isin") and fund["class_isin"] in allt)
    return val, page, as_of, isin_ok


def judge(fund, me, val, page, as_of, isin_ok):
    if val is None:
        return "not_found", "duration figure not found in the document"
    if not isin_ok:
        return "not_found", "ISIN not in the document – may be another share class"
    if not as_of:
        return "check", "value found, document date not recognised – check the date"
    if as_of >= (date.fromisoformat(last_business_day(me)) - timedelta(days=1)).isoformat() and as_of <= me:
        return "found", ""
    if as_of > me:
        return "check", f"document dated {as_of}, after the month end"
    return "old", f"document dated {as_of} – this month's document not published at this link yet"


def collect(me, getter, store=None):
    cfg = json.loads(CFG.read_text())
    rows = []
    for f in cfg["funds"]:
        rec = dict(isin=f["isin"], fund=f["name"], type=f.get("type", ""), value=None, as_of="", page=None, url="",
                   copy=False, status="manual" if not f.get("urls") else "not_found", note=f.get("hint", ""), tried=[])
        best = None
        for u in f.get("urls", []):
            u = u if isinstance(u, dict) else {"url": u}
            url = fill(u["url"], me)
            try:
                pdf = getter(f["isin"], url)
                val, page, as_of, ok = extract(f, pages(pdf))
                st, why = judge(f, me, val, page, as_of, ok)
                rec["tried"].append(dict(url=url, result=st, note=why))
                cand = dict(value=val, as_of=as_of, page=page, url=url, copy=bool(u.get("copy")), status=st,
                            note=why or ("broker copy" if u.get("copy") else ""), sha256=hashlib.sha256(pdf.read_bytes()).hexdigest())
                if store:
                    cand["file"] = store(f["isin"], pdf)
                rank = {"found": 0, "check": 1, "old": 2, "not_found": 3}[st]
                if best is None or rank < best[0]:
                    best = (rank, cand)
                if st == "found":
                    break
            except Exception as e:  # noqa: BLE001
                rec["tried"].append(dict(url=url, result="error", note=str(e)[:160]))
        if best:
            rec.update(best[1])
            if f.get("hint") and best[1]["status"] != "found":
                rec["note"] = "; ".join(x for x in (best[1]["note"], f["hint"]) if x)
        elif f.get("urls"):
            rec["note"] = "; ".join(x for x in ("download failed: " + rec["tried"][-1]["note"] if rec["tried"] else "", f.get("hint", "")) if x)
        rows.append(rec)
        print(f"{rec['status']:9} {f['isin']} {rec['value']} {rec['as_of']} {rec['note'][:90]}")
    return dict(month_end=me, run=time.strftime("%Y-%m-%d"), source="automatic, no AI (tools/duration_auto.py)", rows=rows,
                summary={s: sum(r["status"] == s for r in rows) for s in ("found", "check", "old", "not_found", "manual")})


def main(argv):
    if argv[0] == "--local":
        d, me = Path(argv[1]), argv[2]
        idx = {}
        for p in sorted(d.glob("*.pdf")):
            idx.setdefault(p.name.split("_")[0], []).append(p)
        queue = {}

        def getter(isin, url):
            q = queue.setdefault(isin, list(idx.get(isin, [])))
            if not q:
                raise FileNotFoundError("no local file")
            return q.pop(0)
        res = collect(me, getter)
        print(json.dumps(res["summary"]))
        return
    out, me = Path(argv[0]), (argv[1] if len(argv) > 1 and argv[1] else prev_month_end())
    sys.path.insert(0, str(ROOT / "tools"))
    from fetch_fund_docs import fetch  # noqa: E402
    day = time.strftime("%Y-%m-%d")
    tmp = Path("/tmp/duration_auto"); tmp.mkdir(exist_ok=True)

    def getter(isin, url):
        data, ctype, _ = fetch(url)
        if data[:5] != b"%PDF-":
            raise ValueError(f"not a PDF ({ctype}, {len(data)} B)")
        p = tmp / f"{isin}_{hashlib.sha256(data).hexdigest()[:10]}.pdf"
        p.write_bytes(data)
        time.sleep(1)
        return p

    def store(isin, pdf):
        d = out / day / isin; d.mkdir(parents=True, exist_ok=True)
        dst = d / f"duration_factsheet_{pdf.stem.split('_')[1]}.pdf"
        dst.write_bytes(pdf.read_bytes())
        return f"fund-docs:{dst.relative_to(out)}"
    res = collect(me, getter, store)
    OUTD.mkdir(parents=True, exist_ok=True)
    (OUTD / f"auto_{me}.json").write_text(json.dumps(res, ensure_ascii=False, indent=1))
    print(json.dumps(res["summary"]))


if __name__ == "__main__":
    main(sys.argv[1:])
