"""ESG Engagement: emitentų tvarumo pažanga pagal metus (data/esg/<emitentas>.yaml) -> svetainės duomenys ir Excel.

Kiekvienas emitentas vertinamas tik pagal savo paties tikslus ir ankstesnius metus (emitentai tarpusavyje nelyginami).
Tvarumo ataskaitos metinės, todėl stulpeliai - ataskaitiniai metai (FY2024, FY2025, ...). Kiekvienas teiginys - su šaltiniu
[E# p.N] (p.N - PDF puslapis, nuoroda atidaro šaltinį tame puslapyje). Excel: docs/files/esg_engagement.xlsx.
Engagement (laiškas emitentui) siūlomas tik realioms spragoms, įvertinus SEB sektorių politikas (config/seb_policies.yaml).
Naudojimas: python tools/build_esg.py (kviečiama ir iš build_site.py)
"""
import html
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

ESG_DIR = C.DATA / "esg"
TREND = {"improving": {"lt": "Gerėja", "en": "Improving"}, "stable": {"lt": "Stabilu", "en": "Stable"},
         "mixed": {"lt": "Nevienareikšmiška", "en": "Mixed"}, "worsening": {"lt": "Blogėja", "en": "Worsening"},
         "limited": {"lt": "Mažai duomenų", "en": "Limited data"}, "baseline": {"lt": "Atskaitos metai", "en": "Baseline year"}}
TSTATUS = {"met": {"lt": "Pasiekta", "en": "Met"}, "on_track": {"lt": "Pagal planą", "en": "On track"},
           "behind": {"lt": "Atsilieka", "en": "Behind"}, "missed": {"lt": "Nepasiekta", "en": "Missed"},
           "new": {"lt": "Naujas tikslas", "en": "New target"}, "no_data": {"lt": "Nėra duomenų", "en": "No data"},
           "changed": {"lt": "Pakeistas", "en": "Revised"}}
ENGAGE = {"none": {"lt": "Nereikia", "en": "Not needed"}, "watch": {"lt": "Stebėti", "en": "Monitor"},
          "engage": {"lt": "Reikia", "en": "Needed"}}
PILLAR = {"E": {"lt": "Aplinka", "en": "Environment"}, "S": {"lt": "Socialinė sritis", "en": "Social"}, "G": {"lt": "Valdysena", "en": "Governance"}}


def issuers():
    """[(raktas, esg duomenys)] tokia tvarka kaip config/issuers.yaml."""
    out = []
    for key in C.issuers():
        p = ESG_DIR / f"{key}.yaml"
        if p.exists():
            out.append((key, C.load_yaml(p)))
    return out


def years(all_esg):
    ys = sorted({y for _, e in all_esg for y in (e.get("years") or {})})
    return ys


def _segs_html(segs):
    out, prev = [], False
    for kind, a, b in segs:
        if kind == "t":
            out.append(f"<b>{html.escape(a)}</b>" if b else html.escape(a))
            prev = False
        else:
            link = f'<a href="{html.escape(b)}" target="_blank" rel="noopener">{a}</a>' if b else str(a)
            out.append(f'<sup class="cite">{"," if prev else ""}{link}</sup>')
            prev = True
    return "".join(out)


def site(lang):
    """Svetainės ESG puslapio duomenys viena kalba (tekstai jau HTML su išnašomis)."""
    secs = C.securities()
    all_esg = issuers()
    ys = years(all_esg)
    rows = []
    for key, e in all_esg:
        iss = C.issuers()[key]
        cites = C.Citations(e.get("sources", {}))
        h = lambda x: _segs_html(C.lead_bold(cites.parse(C.tr(x, lang)))) if x else ""  # noqa: E731
        held = [s for s in secs if s["issuer"] == key and s.get("held") == "yes" and s.get("status") != "matured"]
        yr = {}
        for y, d in (e.get("years") or {}).items():
            yr[y] = {"trend": d.get("trend", "limited"), "summary": h(d.get("summary")), "framework": h(d.get("framework")),
                     "assurance": d.get("assurance", ""),
                     "evidence": [{"p": x.get("p", ""), "t": h(x)} for x in d.get("evidence", [])],
                     "concerns": [h(x) for x in d.get("concerns", [])]}
        targets = [{"p": t.get("p", ""), "name": h(t.get("name")), "status": t.get("status", "no_data"),
                    "y": {y: h(t.get(y)) for y in ys if t.get(y)}} for t in e.get("targets", [])]
        metrics = []
        for m in e.get("metrics", []):
            vals = {}
            for y in ["FY2023"] + ys:
                v = m.get(y)
                if v is None:
                    continue
                v = v if isinstance(v, dict) else {"v": v}
                src = v.get("src", "")
                link = ""
                if src:
                    sg = cites.parse(f"[{src}]")
                    link = next((b for k, a, b in sg if k == "c"), "")
                    num = next((a for k, a, b in sg if k == "c"), "")
                else:
                    num = ""
                vals[y] = {"v": v.get("v"), "src": src, "url": link, "n": num, "note": C.tr(v.get("note", ""), lang), "calc": bool(v.get("calc"))}
            metrics.append({"p": m.get("p", ""), "name": C.tr(m.get("name"), lang), "unit": C.tr(m.get("unit", ""), lang),
                            "better": m.get("better", ""), "vals": vals, "note": h(m.get("note"))})
        eng = e.get("engagement") or {}
        rows.append({
            "id": key, "name": iss["name"], "country": iss.get("country", ""), "sector": iss.get("sector", ""),
            "instruments": sorted({s["type"] or "bond" for s in held}), "green": any(s.get("green") == "yes" for s in held),
            "held": [{"isin": s["isin"], "name": s["name"], "type": s["type"], "green": s.get("green")} for s in held],
            "trend": e.get("trend", "limited"), "summary": h(e.get("summary")),
            "years": yr, "targets": targets, "metrics": metrics,
            "engagement": {"level": eng.get("level", "none"), "reason": h(eng.get("reason")),
                           "policies": [C.seb()["policies"].get(p, {}).get("name", p) for p in eng.get("policies", [])],
                           "questions": eng.get("questions", []), "subject": eng.get("subject", ""),
                           "contact": eng.get("contact", "")},
            "analysis_date": str(e.get("analysis_date", "")),
            "yaml": f"data/esg/{key}.yaml",
            "sources": [{"n": n, "title": t, "page": pg, "url": u} for n, t, pg, note, u, k, d in cites.notes()],
        })
    return {"years": ys, "rows": rows,
            "labels": {"trend": {k: v[lang] for k, v in TREND.items()}, "tstatus": {k: v[lang] for k, v in TSTATUS.items()},
                       "engage": {k: v[lang] for k, v in ENGAGE.items()}, "pillar": {k: v[lang] for k, v in PILLAR.items()}}}


def plain(text, cites):
    """Tekstas be išnašų žymių, šaltiniai - atskirai („Pavadinimas p.N“)."""
    segs = cites.parse(text)
    t = "".join(a for k, a, b in segs if k == "t")
    return re.sub(r"\s+", " ", t).strip()


def xlsx(path=None, lang="lt"):
    """Excel: 1 lapas - emitentai × metai (būsena, santrauka), 2 - rodikliai × metai su šaltiniais, 3 - tikslai × metai."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    all_esg = issuers()
    ys = years(all_esg)
    L = {"lt": {"s1": "Emitentai pagal metus", "s2": "Rodikliai", "s3": "Tikslai", "issuer": "Emitentas", "sector": "Sektorius",
                "trend": "Bendra kryptis", "engage": "Engagement", "reason": "Priežastis", "status": "būsena", "summary": "santrauka",
                "metric": "Rodiklis", "unit": "Vnt.", "src": "šaltinis", "target": "Tikslas", "tstatus": "Būsena", "pillar": "Sritis",
                "sources": "Šaltiniai"},
         "en": {"s1": "Issuers by year", "s2": "Metrics", "s3": "Targets", "issuer": "Issuer", "sector": "Sector",
                "trend": "Overall direction", "engage": "Engagement", "reason": "Reason", "status": "status", "summary": "summary",
                "metric": "Metric", "unit": "Unit", "src": "source", "target": "Target", "tstatus": "Status", "pillar": "Pillar",
                "sources": "Sources"}}[lang]
    wb = Workbook()
    head = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="12355B")
    wrap = Alignment(wrap_text=True, vertical="top")

    def src_txt(cites):
        return "\n".join(f"{n}. {t}{' p.' + pg if pg else ''} – {u}" for n, t, pg, note, u, k, d in cites.notes())

    ws = wb.active
    ws.title = L["s1"]
    hdr = [L["issuer"], L["sector"]] + [f"{y} {L['status']}" for y in ys] + [f"{y} {L['summary']}" for y in ys] + [L["trend"], L["engage"], L["reason"], L["sources"]]
    ws.append(hdr)
    for key, e in all_esg:
        cites = C.Citations(e.get("sources", {}))
        yd = e.get("years") or {}
        row = [C.issuers()[key]["name"], C.issuers()[key].get("sector", "")]
        row += [TREND.get(yd.get(y, {}).get("trend", ""), {}).get(lang, "") for y in ys]
        row += [plain(C.tr(yd.get(y, {}).get("summary", ""), lang), cites) for y in ys]
        eng = e.get("engagement") or {}
        row += [TREND.get(e.get("trend", ""), {}).get(lang, ""), ENGAGE.get(eng.get("level", "none"), {}).get(lang, ""),
                plain(C.tr(eng.get("reason", ""), lang), cites)]
        # šaltinių sąrašas pagal visų tekstų išnašas
        for y in ys:
            cites.parse(C.tr(yd.get(y, {}).get("summary", ""), lang))
        row.append(src_txt(cites))
        ws.append(row)
    widths = [22, 26] + [14] * len(ys) + [60] * len(ys) + [16, 14, 50, 60]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws2 = wb.create_sheet(L["s2"])
    ws2.append([L["issuer"], L["pillar"], L["metric"], L["unit"]] + [x for y in ["FY2023"] + ys for x in (y, f"{y} {L['src']}")])
    for key, e in all_esg:
        srcs = e.get("sources", {})
        for m in e.get("metrics", []):
            row = [C.issuers()[key]["name"], m.get("p", ""), C.tr(m.get("name"), lang), C.tr(m.get("unit", ""), lang)]
            for y in ["FY2023"] + ys:
                v = m.get(y)
                v = v if isinstance(v, dict) or v is None else {"v": v}
                if not v:
                    row += ["", ""]
                    continue
                s = v.get("src", "")
                mm = re.match(r"([A-Z]{1,2}\d{1,2})\s*(?:p\.\s*([\d,\- ]+))?", s or "")
                st = ""
                if mm and mm.group(1) in srcs:
                    st = f"{srcs[mm.group(1)]['title']}{' p.' + mm.group(2) if mm.group(2) else ''} – {C.Citations(srcs).url(mm.group(1), mm.group(2) or '')}"
                row += [v.get("v"), st + (" (" + C.tr(v.get("note"), lang) + ")" if v.get("note") else "") + (" [calc]" if v.get("calc") else "")]
            ws2.append(row)
    for i, w in enumerate([22, 8, 44, 14] + [14, 40] * (len(ys) + 1), 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    ws3 = wb.create_sheet(L["s3"])
    ws3.append([L["issuer"], L["pillar"], L["target"]] + ys + [L["tstatus"]])
    for key, e in all_esg:
        cites = C.Citations(e.get("sources", {}))
        for t in e.get("targets", []):
            ws3.append([C.issuers()[key]["name"], t.get("p", ""), plain(C.tr(t.get("name"), lang), cites)]
                       + [plain(C.tr(t.get(y, ""), lang), cites) for y in ys] + [TSTATUS.get(t.get("status", ""), {}).get(lang, "")])
    for i, w in enumerate([22, 8, 50] + [50] * len(ys) + [14], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    for w in (ws, ws2, ws3):
        for c in w[1]:
            c.font, c.fill = head, fill
        for r in w.iter_rows(min_row=2):
            for c in r:
                c.alignment = wrap
        w.freeze_panes = "B2"
    path = path or C.DOCS / "files" / f"esg_engagement_{lang}.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def main():
    for lg in ("lt", "en"):
        p = xlsx(lang=lg)
        print(f"{p.relative_to(C.ROOT)}: {len(issuers())} emitentai")


if __name__ == "__main__":
    main()
