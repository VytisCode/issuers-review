"""Dashboard'o duomenys: docs/data.js iš config/, data/ ir reports/; nukopijuoja site/ -> docs/.

Svetainė - GitHub Pages iš docs/ (docs/ ranka netaisyti; taisyti site/ ir paleisti šį skriptą).
Naudojimas: python tools/build_site.py
"""
import json
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

REPO = "VytisCode/issuers-review"


def main():
    lang = C.settings().get("site_lang", "lt")
    langs = C.settings().get("languages", ["lt"])
    M = C.metrics()
    secs = C.securities()
    news = C.read_csv(C.DATA / "news.csv")
    filings = C.read_csv(C.DATA / "filings.csv")
    out = {"generated": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"), "repo": REPO, "today": C.TODAY,
           "status": C.load_json(C.DATA / "source_status.json"), "issuers": [], "securities": secs,
           "metrics": {k: C.tr(v, lang) for k, v in M.items()},
           "news": news[:600], "filings": [f for f in filings if f.get("baseline") != "yes"]}
    for key, iss in C.issuers().items():
        reps = C.reports(key)
        last = reps[-1] if reps else {}
        files = []
        for r in reps:
            for lg in langs:
                for ext in ("docx", "pdf", "xlsx"):
                    p = C.DOCS / "files" / key / f"{key}_{r['period']}_{lg}.{ext}"
                    if p.exists():
                        files.append({"period": r["period"], "lang": lg, "ext": ext, "url": f"files/{key}/{p.name}", "date": str(r["analysis_date"])})
        hist = C.history(key)
        last_date = str(last.get("analysis_date", ""))
        new_filings = [f for f in out["filings"] if f["issuer"] == key and f["first_seen"] > last_date]
        esg = last.get("esg", {}).get("items", {})
        out["issuers"].append({
            "id": key, "name": iss["name"], "legal_name": iss.get("legal_name", ""), "country": iss.get("country", ""), "sector": iss.get("sector", ""),
            "owner": iss.get("owner", ""), "website": iss.get("website", ""), "ir_page": iss.get("ir_page", ""), "calendar": iss.get("report_calendar", ""),
            "rating": last.get("rating", ""), "period": last.get("period", ""), "analysis_date": last_date, "next_report": last.get("next_report", ""),
            "summary": [C.tr(x, lang) for x in last.get("quick_summary", [])],
            "investor": [C.tr(x, lang) for x in last.get("investor", [])],
            "plans": [{"plan": C.tr(p["plan"], lang), "outcome": C.tr(p["outcome"], lang), "status": p["status"], "src": p.get("src", "")} for p in last.get("plans", [])],
            "esg_review": sum(1 for v in esg.values() if v.get("status") == "review"), "esg_no": sum(1 for v in esg.values() if v.get("status") == "no"),
            "esg_conclusion": C.tr(last.get("esg", {}).get("conclusion"), lang),
            "checks": {k: {"status": v.get("status"), "note": C.tr(v.get("note"), lang)} for k, v in last.get("checks", {}).items()},
            "sources": last.get("sources", {}),
            "files": files, "new_filings": new_filings,
            "periods": last.get("periods", []), "ratios": last.get("ratios", []),
            "history": {m: {p: {"v": r["value"], "u": r["unit"], "s": r["source"], "pg": r.get("page", ""), "ok": r.get("verified", "")} for p, r in v.items()}
                        for m, v in hist.items()},
            "yaml": f"reports/{key}/{last['_file'].name}" if last else "",
        })
    site = C.ROOT / "site"
    C.DOCS.mkdir(exist_ok=True)
    for f in site.iterdir():
        if f.is_file():
            shutil.copy2(f, C.DOCS / f.name)
    (C.DOCS / ".nojekyll").write_text("")
    js = "window.DATA = " + json.dumps(out, ensure_ascii=False, default=str) + ";\n"
    (C.DOCS / "data.js").write_text(js, encoding="utf-8")
    print(f"docs/data.js: {len(out['issuers'])} emitentai, {len(secs)} VP, {len(news)} naujienos, {len(out['filings'])} nauji dokumentai")


if __name__ == "__main__":
    main()
