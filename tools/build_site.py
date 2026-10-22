"""Dashboard'o duomenys: docs/data.js iš config/, data/ ir reports/; nukopijuoja site/ -> docs/.

Tekstai pateikiami abiem svetainės kalbomis (lt, en) su šaltinių išnašomis (HTML).
Svetainė - GitHub Pages (šakninis index.html nukreipia į docs/). docs/ ranka netaisyti; taisyti site/ ir paleisti šį skriptą.
Naudojimas: python tools/build_site.py
"""
import html
import json
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

REPO = "VytisCode/issuers-review"
SITE_LANGS = ["en", "lt"]


def segs_html(segs):
    out, prev_c = [], False
    for kind, a, b in segs:
        if kind == "t":
            out.append(f"<b>{html.escape(a)}</b>" if b else html.escape(a))
            prev_c = False
        else:
            link = f'<a href="{html.escape(b)}" target="_blank" rel="noopener">{a}</a>' if b else str(a)
            out.append(f'<sup class="cite">{"," if prev_c else ""}{link}</sup>')
            prev_c = True
    return "".join(out)


def texts(rep, lang):
    cites = C.Citations(rep.get("sources", {}))
    h = lambda x, lead=False: segs_html(C.lead_bold(cites.parse(C.tr(x, lang))) if lead else cites.parse(C.tr(x, lang)))  # noqa: E731
    tpl = C.seb()["esg_template"]
    esg = rep.get("esg", {})
    items = C.esg_items(rep)
    out = {
        "verdict": h(rep.get("verdict", "")),
        "summary": [h(x, True) for x in rep.get("quick_summary", [])],
        "business": [h(x, True) for x in rep.get("business", [])],
        "financials": [h(x, True) for x in rep.get("financials", [])],
        "investor": [h(x, True) for x in rep.get("investor", [])],
        "press": [h(x, True) for x in rep.get("press", [])],
        "market": [h(x, True) for x in rep.get("market", [])],
        "plans": [{"plan": h(p["plan"]), "outcome": h(C.tr(p["outcome"], lang) + (f" [{p['src']}]" if p.get("src") else "")),
                   "why": h(p.get("why", "")), "status": p["status"]}
                  for p in rep.get("plans", [])],
        "rating": h(rep.get("rating", "")), "next_report": h(rep.get("next_report", "")),
        "esg": [{"heading": s["heading"], "note": s.get("note", ""),
                 "items": [{"id": q["id"], "q": q["q"], "status": items.get(q["id"], {}).get("status", "review"),
                            "points": [{"label": lab, "items": [h(x) for x in its]} for lab, its in C.points(C.tr(items.get(q["id"], {}).get("evidence", ""), lang))]} for q in s["items"]]} for s in tpl["sections"]],
        "esg_conclusion": h(esg.get("conclusion", "")),
    }
    out["sources"] = [{"n": n, "title": t, "page": pg, "note": note, "url": u, "date": d} for n, t, pg, note, u, k, d in cites.notes()]
    return out


def main():
    langs = C.settings().get("languages", ["lt"])
    M = C.metrics()
    secs = C.securities()
    news = C.read_csv(C.DATA / "news.csv")
    filings = C.read_csv(C.DATA / "filings.csv")
    out = {"generated": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"), "repo": REPO, "today": C.TODAY, "isin_store_url": C.settings().get("isin_store_url", "") or "",
           "status": C.load_json(C.DATA / "source_status.json"), "issuers": [], "securities": secs,
           "metrics": {lg: {k: C.tr(v, lg) for k, v in M.items()} for lg in SITE_LANGS},
           "news": news[:600],
           "market": [{"id": sid, "label": {lg: C.tr(C.i18n()["mk_" + sid], lg) for lg in SITE_LANGS}, "last": last, "now": now, "m3": m3, "m12": m12, "src": src}
                      for sid, last, now, m3, m12, src in C.market_table()],
           # svarbiausios naujienos (30 d.): tos pačios temos pranešimai sujungti, „kita“ tema nerodoma, nebent antraštėje neigiamas žodis
           "news_highlights": {lg: [d for d in C.news_digest(None, days=30, lang=lg)[0] if d["topic"] != "kita"][:40] for lg in SITE_LANGS},
           "filings": [f for f in filings if f.get("baseline") != "yes"]}
    archive = C.load_json(C.DATA / "sources_archive.json")
    for key, iss in C.issuers().items():
        reps = C.reports(key)
        last = reps[-1] if reps else {}
        files = []
        for r in reps:
            z = archive.get(f"{key}_{r['period']}_sources.zip")
            if z:
                files.append({"period": r["period"], "lang": "", "ext": "zip", "url": z["url"], "date": z["date"], "n": z["files"], "kb": z["size_kb"]})
            for lg in langs:
                for ext in ("pdf", "docx"):
                    p = C.DOCS / "files" / key / f"{key}_{r['period']}_{lg}.{ext}"
                    if p.exists():
                        files.append({"period": r["period"], "lang": lg, "ext": ext, "url": f"files/{key}/{p.name}", "date": str(r["analysis_date"])})
        hist = C.history(key)
        last_date = str(last.get("analysis_date", ""))
        items = C.esg_items(last) if last else {}
        out["issuers"].append({
            "id": key, "name": iss["name"], "legal_name": iss.get("legal_name", ""), "country": iss.get("country", ""), "sector": iss.get("sector", ""),
            "owner": iss.get("owner", ""), "website": iss.get("website", ""), "ir_page": iss.get("ir_page", ""), "calendar": iss.get("report_calendar", ""),
            "rating_plain": C.plain(last.get("rating", "")), "period": last.get("period", ""), "analysis_date": last_date,
            "verdict": (last.get("verdict") or {}).get("status", ""), "ratings": last.get("ratings", []),
            "esg_review": sum(1 for v in items.values() if v.get("status") == "review"),
            "files": files, "new_filings": [f for f in out["filings"] if f["issuer"] == key and f["first_seen"] > last_date],
            "periods": last.get("periods", []), "ratios": last.get("ratios", []), "highlights": last.get("highlights", []),
            "history": {m: {p: {"v": r["value"], "u": r["unit"], "s": r["source"], "pg": r.get("page", ""), "ok": r.get("verified", "")} for p, r in v.items()}
                        for m, v in hist.items()},
            "judge": {m: dict(zip(("last", "prev", "txt", "trend", "level"), C.ratio_change(m, hist[m], last.get("periods", []), "en"))) | {"txt_lt": C.ratio_change(m, hist[m], last.get("periods", []), "lt")[2]}
                      for m in last.get("ratios", []) if m in hist} if last else {},
            "src_urls": {k: v.get("url", "") for k, v in last.get("sources", {}).items()},
            "yaml": f"reports/{key}/{last['_file'].name}" if last else "",
            "t": {lg: texts(last, lg) for lg in SITE_LANGS} if last else {},
        })
    site = C.ROOT / "site"
    C.DOCS.mkdir(exist_ok=True)
    for f in site.iterdir():
        if f.is_file():
            shutil.copy2(f, C.DOCS / f.name)
    (C.DOCS / ".nojekyll").write_text("")
    (C.DOCS / "data.js").write_text("window.DATA = " + json.dumps(out, ensure_ascii=False, default=str) + ";\n", encoding="utf-8")
    print(f"docs/data.js: {len(out['issuers'])} emitentai, {len(secs)} VP, {len(news)} naujienos, {len(out['filings'])} nauji dokumentai")


if __name__ == "__main__":
    main()
