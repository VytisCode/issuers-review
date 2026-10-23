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
import build_esg  # noqa: E402

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
        # svetainėje - paprasta kalba (verdict.plain), jei yra; Word/PDF lieka verdict
        "verdict": h((rep.get("verdict") or {}).get("plain") or rep.get("verdict", "")),
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


# laikotarpio žymės ataskaitų pavadinimuose (kad būtų galima atpažinti, už kokį ketvirtį paskelbta ataskaita)
PERIOD_RX = {3: r"\bQ1\b|\b1Q\b|3M|three months|first quarter|I ketv|3 mėn|3 kuu|3 mēn|january.?march|sausio.?kovo",
             6: r"\bQ2\b|\b2Q\b|6M|\bH1\b|\b1H\b|six months|half.?year|first half|II ketv|6 mėn|pusmeč|6 kuu|poolaasta|6 mēn|pusgad|january.?june|sausio.?birželio",
             9: r"\bQ3\b|\b3Q\b|9M|nine months|third quarter|III ketv|9 mėn|9 kuu|9 mēn|january.?september|sausio.?rugsėjo|jaanuar.?september",
             12: r"\bQ4\b|\b4Q\b|12M|annual|full.?year|metin|aastaaruanne|gada pārskats|twelve months|IV ketv|12 mėn"}


NOT_REPORT_RX = r"tegevusmahud|volumes|traffic|passenger|krovin|calendar|kalend|will publish|to publish|avaldab|paskelbs|publicēs|webinar|investor call|conference call|audiokonver|veebiseminar"


def readiness(news, filings):
    """Ar jau metas generuoti naujas apžvalgas: kokio laikotarpio dabartinės analizės ir kiek emitentų jau paskelbė naujesnę ataskaitą.
    Tikslinis laikotarpis - paskutinio pasibaigusio ketvirčio pabaiga. Paskelbimas atpažįstamas be AI: nauji emitentų puslapių dokumentai
    (data/filings.csv) ir naujienos su to ketvirčio žyme pavadinime."""
    import re
    from datetime import date, timedelta
    today = date.fromisoformat(C.TODAY)
    qe = [date(today.year - (1 if m > today.month or (m == today.month and today.day < 28) else 0), m, 1) for m in (3, 6, 9, 12)]
    target_m = max(qe)
    target = f"{target_m.year}-{target_m.month:02d}-{30 if target_m.month in (6, 9) else 31}"
    rows = []
    for key, iss in C.issuers().items():
        reps = C.reports(key)
        last = reps[-1] if reps else {}
        fy = iss.get("fiscal_year_end", "12-31")
        end = C.period_end(last.get("period", ""), fy) if last else ""
        adate = str(last.get("analysis_date", ""))
        exp, exact = C.expected_date(last.get("next_report") or iss.get("report_calendar", ""), adate)
        # pusmetinės ataskaitos (pvz. LATRAPS): emitento tikslinis laikotarpis - paskutinis pasibaigęs pusmetis
        months = [int(fy[:2]), (int(fy[:2]) + 6 - 1) % 12 + 1] if iss.get("report_frequency") == "semiannual" else [3, 6, 9, 12]
        own = max(f"{int(target[:4]) - (1 if mo > int(target[5:7]) else 0)}-{mo:02d}-{30 if mo in (6, 9) else 31}" for mo in months)
        found = []
        if not end or end < own:
            rx = re.compile(PERIOD_RX[int(own[5:7])] if fy == "12-31" else r"(?!)", re.I)
            year = own[:4]
            for f in filings:
                if f["issuer"] == key and f.get("baseline") != "yes" and f["first_seen"] > adate and f.get("kind") in ("interim", "annual"):
                    found.append({"date": f["first_seen"], "title": f["title"], "url": f["url"], "sure": bool(rx.search(f["title"] + " " + f["url"]))})
            for n in news:
                if n["issuer"] == key and n["date"] > own and rx.search(n["title"]) and not re.search(NOT_REPORT_RX, n["title"], re.I) and (year in n["title"] or n["date"][:4] == year):
                    found.append({"date": n["date"], "title": n["title"], "url": n["url"], "sure": True})
            found.sort(key=lambda x: (not x["sure"], x["date"]))
        if end and end >= own:
            st = "done"
        elif iss.get("manual_reports"):
            st = "manual"
        elif any(x["sure"] for x in found):
            st = "published"
        elif found:
            st = "maybe"
        elif exp and exp < C.TODAY:
            st = "late"
        else:
            st = "waiting"
        rows.append({"id": key, "name": iss["name"], "period": last.get("period", ""), "end": end, "analysis_date": adate,
                     "expected": exp, "exact": exact, "status": st, "found": found[:3]})
    periods = {}  # pagal duomenų datą (laikotarpio pabaigą): 6M2026, 1H2026 ir LATRAPS FY2026 – visi „6M 2026 (Q2)“
    for r in rows:
        if r["end"]:
            lab = C.quarter_label(r["end"])
            periods[lab] = periods.get(lab, 0) + 1
    return {"target": target, "target_label": C.quarter_label(target), "rows": rows,
            "periods": sorted(periods.items(), key=lambda kv: (-kv[1], kv[0])),
            "last_expected": max([r["expected"] for r in rows if r["status"] in ("waiting", "late") and r["expected"]] or [""])}


EMAIL = {
    "en": {"subject": "Issuer review completed – {period}", "hi": "Hi all,",
           "intro": "The issuer review based on the latest published reports ({periods}) is complete: {n} issuers, {h} securities held.",
           "files": "Full reviews (PDF / Word) and source files: [link]",
           "glance": "At a glance: {a} need attention · {w} on watch · {o} stable",
           "action": "Needs attention", "watch": "Watch list", "ok": "Stable – no action", "mat": "Upcoming maturities and calls (next 12 months)",
           "matures": "matures", "call": "call date", "none_mat": "No maturities or calls in the next 12 months.", "src": "source",
           "bye": "Questions and comments are welcome."},
    "lt": {"subject": "Emitentų apžvalga atlikta – {period}", "hi": "Sveiki,",
           "intro": "Atlikta emitentų apžvalga pagal naujausias paskelbtas ataskaitas ({periods}): {n} emitentai, {h} turimi vertybiniai popieriai.",
           "files": "Išsamios apžvalgos (PDF / Word) ir šaltinių failai: [nuoroda]",
           "glance": "Trumpai: {a} reikia dėmesio · {w} stebime · {o} stabilūs",
           "action": "Reikia dėmesio", "watch": "Stebime", "ok": "Stabilūs – veiksmų nereikia", "mat": "Artėjantys išpirkimai ir call datos (12 mėn.)",
           "matures": "išpirkimas", "call": "call data", "none_mat": "Per artimiausius 12 mėn. išpirkimų ir call datų nėra.", "src": "šaltinis",
           "bye": "Klausimai ir pastabos – laukiami."},
}


def _sentences(text, n):
    import re
    parts = re.split(r"(?<=[.!?])\s+(?=[A-ZĄČĘĖĮŠŲŪŽ0-9])", text.strip())
    return " ".join(parts[:n]).strip()


def email(lang, readiness_info):
    """Laiškas kolegoms (be AI): sudaromas iš kiekvieno emitento verdikto ir santraukos. Daugiau vietos – reikia dėmesio / stebimiems,
    stabiliems – viena eilutė. Pabaigoje - išpirkimai ir call datos per 12 mėn. Vieta nuorodai į failus – laiško pabaigoje (įrašo savininkas)."""
    import re
    from datetime import date, timedelta
    E, out = EMAIL[lang], {"action": [], "watch": [], "ok": []}
    secs = [x for x in C.securities() if x.get("held") == "yes" and x["status"] == "active"]
    for key, iss in C.issuers().items():
        reps = C.reports(key)
        if not reps:
            continue
        rep = reps[-1]
        v = rep.get("verdict") or {}
        st = v.get("status") or "ok"
        cites = C.Citations(rep.get("sources", {}))
        segs = cites.parse(C.tr(v, lang))
        url = next((b for k, a, b in segs if k == "c" and b), "")
        txt = C.plain(C.tr(v, lang))
        first, _, rest_ = txt.partition(". ")
        if rest_ and len(first.split()) <= 3:  # „Position OK.“ / „Worth watching.“ – žyma, ne turinys
            txt = rest_
        item = {"name": iss["name"], "period": rep.get("period", ""), "url": url, "green": any(x["green"] == "yes" and x["issuer"] == key for x in secs)}
        item["text"] = _sentences(txt, 2) if len(_sentences(txt, 2)) <= 230 else _sentences(txt, 1)
        if st != "ok":
            pts = []
            for q in rep.get("quick_summary", [])[:2]:
                q = C.plain(C.tr(q, lang))
                head, _, rest = q.partition(":")
                pts.append((head.strip(), _sentences(rest, 1)) if rest else ("", _sentences(q, 1)))
            item["points"] = pts
        out.setdefault(st, []).append(item)
    today = date.fromisoformat(C.TODAY)
    soon = []
    for x in secs:
        for k, lab in (("maturity_date", E["matures"]), ("call_date", E["call"])):
            d = x.get(k)
            if d and C.TODAY <= d <= (today + timedelta(days=365)).isoformat():
                soon.append((d, x, lab))
    soon.sort(key=lambda t: t[0])
    per = ", ".join(p for p, _ in readiness_info["periods"][:3])
    main_period = readiness_info["periods"][0][0] if readiness_info["periods"] else ""
    return {"subject": E["subject"].format(period=main_period), "E": E, "groups": out, "soon": [{"date": d, "name": x["name"], "isin": x["isin"], "what": lab,
            "amount": x.get("amount_m", ""), "green": x["green"] == "yes", "issuer": C.issuers().get(x["issuer"], {}).get("name", x["issuer"])} for d, x, lab in soon],
            "periods": per, "n": len(C.issuers()), "h": len(secs)}


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
           # antraščių vertimai (tools/translate_news.py): {kalba: {naujienos id: tekstas}}
           "news_tr": {lg: {r["id"]: r["text"] for r in C.read_csv(C.DATA / "news_tr.csv") if r["to"] == lg} for lg in SITE_LANGS} if (C.DATA / "news_tr.csv").exists() else {},
           "market": [{"id": sid, "label": {lg: C.tr(C.i18n()["mk_" + sid], lg) for lg in SITE_LANGS}, "last": last, "now": now, "m3": m3, "m12": m12, "src": src}
                      for sid, last, now, m3, m12, src in C.market_table()],
           # svarbiausios naujienos (30 d.): tos pačios temos pranešimai sujungti, „kita“ tema nerodoma, nebent antraštėje neigiamas žodis
           "news_highlights": {lg: [d for d in C.news_digest(None, days=30, lang=lg)[0] if d["topic"] != "kita"][:40] for lg in SITE_LANGS},
           "filings": [f for f in filings if f.get("baseline") != "yes"],
           "readiness": readiness(news, filings)}
    out["email"] = {lg: email(lg, out["readiness"]) for lg in SITE_LANGS}
    out["esg"] = {lg: build_esg.site(lg) for lg in SITE_LANGS}
    build_esg.main()
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
            "owner": iss.get("owner", ""), "website": iss.get("website", ""), "ir_page": iss.get("ir_page", ""), "calendar": iss.get("report_calendar", ""), "manual": bool(iss.get("manual_reports")),
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
    # užšifruotas DD archyvas (tools/dd_vault.py): raktas ir įrašai - į data.js, failai - į docs/dd/
    ddir = C.DATA / "dd"
    if (ddir / "key.json").exists():
        out["dd"] = {"key": json.loads((ddir / "key.json").read_text()), "items": json.loads((ddir / "index.json").read_text())}
        (C.DOCS / "dd").mkdir(exist_ok=True)
        for f in ddir.glob("*.bin"):
            shutil.copy2(f, C.DOCS / "dd" / f.name)
    (C.DOCS / "data.js").write_text("window.DATA = " + json.dumps(out, ensure_ascii=False, default=str) + ";\n", encoding="utf-8")
    print(f"docs/data.js: {len(out['issuers'])} emitentai, {len(secs)} VP, {len(news)} naujienos, {len(out['filings'])} nauji dokumentai")


if __name__ == "__main__":
    main()
