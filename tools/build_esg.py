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
         "mixed": {"lt": "Vieni geriau, kiti blogiau", "en": "Some better, some worse"}, "worsening": {"lt": "Blogėja", "en": "Worsening"},
         "limited": {"lt": "Mažai duomenų", "en": "Limited data"}, "baseline": {"lt": "Atskaitos metai", "en": "Baseline year"}}
TSTATUS = {"met": {"lt": "Pasiekta", "en": "Met"}, "on_track": {"lt": "Pagal planą", "en": "On track"},
           "behind": {"lt": "Atsilieka", "en": "Behind"}, "missed": {"lt": "Nepasiekta", "en": "Missed"},
           "new": {"lt": "Naujas tikslas", "en": "New target"}, "no_data": {"lt": "Nėra duomenų", "en": "No data"},
           "changed": {"lt": "Pakeistas", "en": "Revised"}}
ENGAGE = {"none": {"lt": "Viskas gerai", "en": "All fine"}, "watch": {"lt": "Tik stebėti", "en": "Just monitor"},
          "engage": {"lt": "Siūloma parašyti", "en": "Suggest writing"}}
# Savininkui: kiek rimta spraga (įrankio vertinimas, ne emitento)
SEVERITY = {"minor": {"lt": "Smulkmena", "en": "Minor"}, "moderate": {"lt": "Vidutinė", "en": "Moderate"},
            "major": {"lt": "Didelė", "en": "Major"}}
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


def eng_text(eng, lang):
    """Excel: svarba + verdiktas + spragos punktais (kas / kodėl / kaip pataisyti); be jų - reason."""
    pl = lambda x: C.plain(C.tr(x, lang)).strip() if x else ""  # noqa: E731
    if not eng.get("issues"):
        return pl(eng.get("reason"))
    lab = {"lt": ("Svarba", "Kodėl", "Kas padėtų"), "en": ("Severity", "Why", "What would help")}[lang if lang == "lt" else "en"]
    out = []
    if eng.get("severity"):
        out.append(f"{lab[0]}: {SEVERITY.get(eng['severity'], {}).get(lang, eng['severity'])}. {pl(eng.get('verdict'))}".strip())
    for i in eng["issues"]:
        out.append(f"• {pl(i.get('what'))}\n  {lab[1]}: {pl(i.get('why'))}\n  {lab[2]}: {pl(i.get('fix'))}")
    return "\n".join(out)


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
            # paprasta kalba: kas gerai / kas negerai / ką tai reiškia mums
            "brief": {"good": [h(x) for x in (e.get("brief") or {}).get("good", [])], "bad": [h(x) for x in (e.get("brief") or {}).get("bad", [])],
                      "means": h((e.get("brief") or {}).get("means"))},
            "years": yr, "targets": targets, "metrics": metrics,
            "engagement": {"level": eng.get("level", "none"), "reason": h(eng.get("reason")),
                           "severity": eng.get("severity", ""), "verdict": h(eng.get("verdict")),
                           "issues": [{k: h(i.get(k)) for k in ("what", "why", "fix")} for i in eng.get("issues", [])],
                           "policies": [C.seb()["policies"].get(p, {}).get("name", p) for p in eng.get("policies", [])],
                           # Klausimai: laiške - angliškai (q), vidinė užuomina savininkui (hint) - pasirinkta kalba
                           "questions": [q if isinstance(q, dict) else {"q": q} for q in eng.get("questions", [])],
                           "hints": [h(q.get("hint")) if isinstance(q, dict) else "" for q in eng.get("questions", [])],
                           "noticed": C.tr(eng.get("noticed", ""), "en"),
                           "subject": eng.get("subject", ""), "contact": eng.get("contact", "")},
            "analysis_date": str(e.get("analysis_date", "")),
            "yaml": f"data/esg/{key}.yaml",
            "sources": [{"n": n, "title": t, "page": pg, "url": u} for n, t, pg, note, u, k, d in cites.notes()],
        })
    return {"years": ys, "rows": rows,
            "labels": {"trend": {k: v[lang] for k, v in TREND.items()}, "tstatus": {k: v[lang] for k, v in TSTATUS.items()},
                       "engage": {k: v[lang] for k, v in ENGAGE.items()}, "pillar": {k: v[lang] for k, v in PILLAR.items()},
                       "severity": {k: v[lang] for k, v in SEVERITY.items()}}}


def plain(text, cites):
    """Tekstas be išnašų žymių, šaltiniai - atskirai („Pavadinimas p.N“)."""
    segs = cites.parse(text)
    t = "".join(a for k, a, b in segs if k == "t")
    return re.sub(r"\s+", " ", t).strip()


def refs(text):
    """[(raktas, puslapis)] iš teksto išnašų [E1 p.12; E2]."""
    out = []
    for m in C.CITE.finditer(text or ""):
        for part in m.group(1).split(";"):
            pm = C.PART.match(part)
            if pm:
                out.append((pm.group(1), (pm.group(2) or "").replace(" ", "")))
    return out


# Excel spalvos (tos pačios reikšmės kaip svetainės žetonuose)
XL_TONE = {"good": ("E6F4EA", "1E7B34"), "warn": ("FFF4E5", "8A5300"), "bad": ("FDECEA", "B3261E"),
           "info": ("E8F0FE", "1A56B8"), "mute": ("F1F3F4", "5F6368")}
XL_TREND = {"improving": "good", "stable": "info", "mixed": "warn", "worsening": "bad", "limited": "mute", "baseline": "info"}
XL_TSTATUS = {"met": "good", "on_track": "good", "behind": "warn", "missed": "bad", "new": "info", "no_data": "mute", "changed": "info"}
XL_ENGAGE = {"none": "good", "watch": "warn", "engage": "bad"}


def xlsx(path=None, lang="lt"):
    """Excel savininko ataskaitai: Apžvalga (emitentai × metai), Rodikliai × metai su šaltiniais (nuorodos į PDF puslapį),
    Tikslai, Įrodymai (kiekvienas faktas su šaltiniu), Šaltiniai. Emitentai tarpusavyje nelyginami."""
    import math
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    all_esg = issuers()
    ys = years(all_esg)
    mys = ["FY2023"] + ys
    lt = lang == "lt"
    L = {"lt": {"s1": "Apžvalga", "s2": "Rodikliai", "s3": "Tikslai", "s4": "Įrodymai", "s5": "Šaltiniai",
                "title": "ESG pažanga pagal emitentus", "issuer": "Emitentas", "country": "Šalis", "held": "Turimi VP",
                "trend": "Bendra kryptis", "engage": "Engagement", "reason": "Vertinimas ir spragos", "status": "būsena", "summary": "santrauka",
                "metric": "Rodiklis", "unit": "Vnt.", "src": "šaltinis", "target": "Tikslas", "tstatus": "Būsena", "pillar": "Sritis",
                "change": "Pokytis per metus", "year": "Metai", "kind": "Tipas", "fact": "Faktas", "ev": "Įrodymas", "gap": "Spraga",
                "fw": "Ataskaita ir patikra", "srcs": "Šaltiniai", "key": "Žymė", "doc": "Dokumentas", "date": "Data", "link": "Nuoroda",
                "open": "Atidaryti", "note": "Pastaba", "more": "Kiti šaltiniai",
                "sub": "Kiekvienas emitentas vertinamas pagal savo tikslus ir ankstesnius metus (tarpusavyje nelyginami). "
                       "Kiekvienas faktas turi šaltinį su PDF puslapiu – spauskite nuorodą.",
                "legend": "Spalvos: žalia – gerai / pagal planą, geltona – nevienareikšmiška / atsilieka / stebėti, raudona – blogėja / nepasiekta / reikia kreiptis, pilka – mažai duomenų.",
                "made": "Parengta", "calc": "mūsų skaičiavimas"},
         "en": {"s1": "Overview", "s2": "Metrics", "s3": "Targets", "s4": "Evidence", "s5": "Sources",
                "title": "ESG progress by issuer", "issuer": "Issuer", "country": "Country", "held": "Holdings",
                "trend": "Overall direction", "engage": "Engagement", "reason": "Assessment and gaps", "status": "status", "summary": "summary",
                "metric": "Metric", "unit": "Unit", "src": "source", "target": "Target", "tstatus": "Status", "pillar": "Pillar",
                "change": "Change y/y", "year": "Year", "kind": "Type", "fact": "Fact", "ev": "Evidence", "gap": "Gap",
                "fw": "Reporting and assurance", "srcs": "Sources", "key": "Ref", "doc": "Document", "date": "Date", "link": "Link",
                "open": "Open", "note": "Note", "more": "Other sources",
                "sub": "Each issuer is assessed against its own targets and prior years (issuers are not compared with each other). "
                       "Every fact has a source with the PDF page – click the link.",
                "legend": "Colours: green – good / on track, amber – mixed / behind / monitor, red – worsening / missed / engagement needed, grey – limited data.",
                "made": "Prepared", "calc": "own calculation"}}[lang]
    NAVY, INK, MUTED, LINE, BAND = "12355B", "1F2933", "6B7280", "D9DEE7", "F6F8FB"
    F = "Calibri"
    f_body, f_head = Font(name=F, size=10, color=INK), Font(name=F, size=10, bold=True, color="FFFFFF")
    f_bold, f_link = Font(name=F, size=10, bold=True, color=INK), Font(name=F, size=10, color="1A56B8", underline="single")
    f_muted = Font(name=F, size=9, color=MUTED)
    fill_head, fill_band = PatternFill("solid", fgColor=NAVY), PatternFill("solid", fgColor=BAND)
    top = Alignment(wrap_text=True, vertical="top")
    mid = Alignment(wrap_text=True, vertical="center")
    centre = Alignment(horizontal="center", vertical="top", wrap_text=True)
    line = Border(bottom=Side(style="thin", color=LINE))
    secs = C.securities()
    names = {k: C.issuers()[k]["name"] for k, _ in all_esg}

    def src_label(srcs, key, pg):
        """Trumpas šaltinio vardas langeliui: „Metinė ataskaita 2025, p. 122“ (pilnas pavadinimas - lape Šaltiniai)."""
        t = srcs.get(key, {}).get("title", key)
        yr = re.findall(r"20\d\d", t)
        low = t.lower()
        kind = ("Metinė ataskaita" if lt else "Annual report") if re.search(r"annual|metin|gada|aastaaruanne|integrated|management report|vadovybės", low) else \
            ("Darnumo ataskaita" if lt else "Sustainability report") if re.search(r"sustainab|ilgtspēj|tvarumo ataskait|esg", low) else t[:28]
        return f"{kind}{' ' + yr[0] if yr else ''}{', p. ' + pg if pg else ''}"

    def link_for(srcs, key, pg):
        return C.Citations(srcs).url(key, pg)

    def tone(cell, kind):
        bg, fg = XL_TONE[kind]
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(name=F, size=10, bold=True, color=fg)
        cell.alignment = centre

    def sheet(ws, title, headers, widths, first_row=4):
        ws.sheet_view.showGridLines = False
        ws["A1"] = title
        ws["A1"].font = Font(name=F, size=16, bold=True, color=NAVY)
        ws["A2"] = f"{L['sub']}  {L['made']}: {max(str(e.get('analysis_date', '')) for _, e in all_esg)}"
        ws["A2"].font = f_muted
        for i, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(row=first_row, column=i, value=h)
            c.font, c.fill, c.alignment = f_head, fill_head, mid
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[first_row].height = 30
        return first_row

    def finish(ws, hrow, ncols, freeze_col="B"):
        last = ws.max_row
        for r in range(hrow + 1, last + 1):
            h = 1
            for c in range(1, ncols + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = line
                if cell.font is None or cell.font == Font():
                    cell.font = f_body
                if cell.alignment is None or not cell.alignment.wrap_text:
                    cell.alignment = top
                v = cell.value
                if isinstance(v, str) and v:
                    w = ws.column_dimensions[get_column_letter(c)].width or 10
                    h = max(h, sum(math.ceil(max(len(p), 1) / max(w * 0.95, 1)) for p in v.split("\n")))
            ws.row_dimensions[r].height = max(16, min(h * 13 + 4, 360))
        ws.freeze_panes = f"{freeze_col}{hrow + 1}"
        ws.auto_filter.ref = f"A{hrow}:{get_column_letter(ncols)}{last}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = f"{hrow}:{hrow}"
        ws.oddFooter.center.text = "&P / &N"

    def band(ws, row, ncols, on):
        if on:
            for c in range(1, ncols + 1):
                cell = ws.cell(row=row, column=c)
                if cell.fill is None or cell.fill.fgColor.rgb in (None, "00000000"):
                    cell.fill = fill_band

    def put(ws, row, col, value, font=None, link=None, align=None):
        if isinstance(value, str):
            value = re.sub(r"(?<![.])\.\.(?![.])", ".", value)
        c = ws.cell(row=row, column=col, value=value)
        c.font = font or f_body
        c.alignment = align or top
        if link:
            c.hyperlink = link
            c.font = f_link
        return c

    wb = Workbook()
    # ---- 4. Įrodymai (pirmiau, kad Apžvalga galėtų nurodyti eilutę) ----
    ws4 = wb.active
    ws4.title = L["s4"]
    h4 = [L["issuer"], L["year"], L["kind"], L["pillar"], L["fact"], L["src"].capitalize(), L["more"]]
    hr4 = sheet(ws4, L["title"], h4, [22, 9, 13, 14, 80, 30, 30])
    first_ev = {}
    r = hr4
    for gi, (key, e) in enumerate(all_esg):
        srcs = e.get("sources", {})
        for y in sorted((e.get("years") or {}), reverse=True):
            d = e["years"][y]
            items = [(L["fw"], "", d.get("framework"))] + [(L["ev"], x.get("p", ""), x) for x in d.get("evidence", [])] \
                + [(L["gap"], "", x) for x in d.get("concerns", [])]
            for kind, p, x in items:
                txt = C.tr(x, lang) if x else ""
                if not txt:
                    continue
                r += 1
                first_ev.setdefault(key, r)
                rf = refs(txt)
                put(ws4, r, 1, names[key], f_bold)
                put(ws4, r, 2, y, align=centre)
                c = put(ws4, r, 3, kind, align=centre)
                if kind == L["gap"]:
                    tone(c, "bad")
                elif kind == L["fw"]:
                    tone(c, "info")
                put(ws4, r, 4, PILLAR.get(p, {}).get(lang, ""))
                put(ws4, r, 5, C.plain(txt).strip())
                if rf:
                    put(ws4, r, 6, src_label(srcs, *rf[0]), link=link_for(srcs, *rf[0]))
                    put(ws4, r, 7, "; ".join(src_label(srcs, k, pg) for k, pg in rf[1:]), f_muted)
                band(ws4, r, len(h4), gi % 2)
    finish(ws4, hr4, len(h4), "C")

    # ---- 1. Apžvalga ----
    ws = wb.create_sheet(L["s1"], 0)
    h1 = [L["issuer"], L["country"], L["held"]] + [x for y in ys for x in (f"{y} {L['status']}", f"{y} {L['summary']}")] + [L["trend"], L["engage"], L["reason"]]
    hr1 = sheet(ws, L["title"], h1, [24, 7, 18] + [17, 52] * len(ys) + [17, 13, 56])
    ws["A3"] = L["legend"]
    ws["A3"].font = f_muted
    for gi, (key, e) in enumerate(all_esg):
        r = hr1 + 1 + gi
        iss = C.issuers()[key]
        held = [s["isin"] for s in secs if s["issuer"] == key and s.get("held") == "yes" and s.get("status") != "matured"]
        put(ws, r, 1, names[key], f_link if key in first_ev else f_bold,
            link=f"#'{L['s4']}'!A{first_ev[key]}" if key in first_ev else None)
        put(ws, r, 2, iss.get("country", ""), align=centre)
        put(ws, r, 3, "\n".join(held), f_muted)
        yd = e.get("years") or {}
        col = 4
        for y in ys:
            d = yd.get(y)
            if d:
                tone(put(ws, r, col, TREND.get(d.get("trend", ""), {}).get(lang, "")), XL_TREND.get(d.get("trend"), "mute"))
                put(ws, r, col + 1, C.plain(C.tr(d.get("summary", ""), lang)).strip())
            else:
                put(ws, r, col, "–", f_muted, align=centre)
            col += 2
        tone(put(ws, r, col, TREND.get(e.get("trend", ""), {}).get(lang, "")), XL_TREND.get(e.get("trend"), "mute"))
        eng = (e.get("engagement") or {}).get("level", "none")
        tone(put(ws, r, col + 1, ENGAGE.get(eng, {}).get(lang, "")), XL_ENGAGE.get(eng, "mute"))
        put(ws, r, col + 2, eng_text(e.get("engagement") or {}, lang))
        band(ws, r, len(h1), gi % 2)
    finish(ws, hr1, len(h1), "B")

    # ---- 2. Rodikliai ----
    ws2 = wb.create_sheet(L["s2"], 1)
    h2 = [L["issuer"], L["pillar"], L["metric"], L["unit"]] + mys + [L["change"]] + [f"{y} {L['src']}" for y in mys] + [L["note"]]
    hr2 = sheet(ws2, L["title"], h2, [22, 12, 38, 14] + [11] * len(mys) + [11] + [26] * len(mys) + [34])
    r = hr2
    for gi, (key, e) in enumerate(all_esg):
        srcs = e.get("sources", {})
        for m in e.get("metrics", []):
            r += 1
            put(ws2, r, 1, names[key], f_bold)
            put(ws2, r, 2, PILLAR.get(m.get("p", ""), {}).get(lang, ""))
            put(ws2, r, 3, C.tr(m.get("name"), lang))
            put(ws2, r, 4, C.tr(m.get("unit", ""), lang), f_muted)
            notes, vals = [], []
            for i, y in enumerate(mys):
                v = m.get(y)
                v = v if isinstance(v, dict) or v is None else {"v": v}
                if not v:
                    vals.append(None)
                    continue
                val = v.get("v")
                vals.append(val if isinstance(val, (int, float)) else None)
                c = put(ws2, r, 5 + i, val, align=Alignment(horizontal="right", vertical="top"))
                if isinstance(val, float) and not val.is_integer():
                    c.number_format = "#,##0." + "0" * min(3, len(repr(val).split(".")[1]))
                elif isinstance(val, (int, float)):
                    c.number_format = "#,##0"
                rf = refs(f"[{v.get('src', '')}]")
                if rf:
                    put(ws2, r, 6 + len(mys) + i, src_label(srcs, *rf[0]), link=link_for(srcs, *rf[0]))
                if v.get("note"):
                    notes.append(f"{y}: {C.tr(v.get('note'), lang)}")
                if v.get("calc"):
                    notes.append(f"{y}: {L['calc']}")
            nums = [(i, x) for i, x in enumerate(vals) if x is not None]
            if len(nums) >= 2 and nums[-2][1]:
                ch = (nums[-1][1] - nums[-2][1]) / abs(nums[-2][1])
                c = put(ws2, r, 5 + len(mys), ch, align=Alignment(horizontal="right", vertical="top"))
                c.number_format = "+0.0%;-0.0%;0.0%"
                good = (ch < 0) if m.get("better") == "down" else (ch > 0) if m.get("better") == "up" else None
                if good is not None and abs(ch) > 0.0005:
                    c.font = Font(name=F, size=10, bold=True, color=XL_TONE["good" if good else "bad"][1])
            if notes:
                put(ws2, r, 6 + 2 * len(mys), "\n".join(notes), f_muted)
            band(ws2, r, len(h2), gi % 2)
    finish(ws2, hr2, len(h2), "D")

    # ---- 3. Tikslai ----
    ws3 = wb.create_sheet(L["s3"], 2)
    h3 = [L["issuer"], L["pillar"], L["target"]] + ys + [L["tstatus"], L["src"].capitalize()]
    hr3 = sheet(ws3, L["title"], h3, [22, 12, 46] + [40] * len(ys) + [16, 28])
    r = hr3
    for gi, (key, e) in enumerate(all_esg):
        srcs = e.get("sources", {})
        for t in e.get("targets", []):
            r += 1
            put(ws3, r, 1, names[key], f_bold)
            put(ws3, r, 2, PILLAR.get(t.get("p", ""), {}).get(lang, ""))
            put(ws3, r, 3, C.plain(C.tr(t.get("name"), lang)).strip())
            allrefs = refs(C.tr(t.get("name"), lang))
            for i, y in enumerate(ys):
                txt = C.tr(t.get(y, ""), lang)
                allrefs += refs(txt)
                put(ws3, r, 4 + i, C.plain(txt).strip() or "–")
            st = t.get("status", "no_data")
            tone(put(ws3, r, 4 + len(ys), TSTATUS.get(st, {}).get(lang, "")), XL_TSTATUS.get(st, "mute"))
            if allrefs:
                k, pg = allrefs[-1]
                put(ws3, r, 5 + len(ys), src_label(srcs, k, pg), link=link_for(srcs, k, pg))
            band(ws3, r, len(h3), gi % 2)
    finish(ws3, hr3, len(h3), "D")

    # ---- 5. Šaltiniai ----
    ws5 = wb.create_sheet(L["s5"])
    h5 = [L["issuer"], L["key"], L["doc"], L["date"], L["link"]]
    hr5 = sheet(ws5, L["title"], h5, [22, 7, 60, 10, 70])
    r = hr5
    for gi, (key, e) in enumerate(all_esg):
        for k, s in (e.get("sources") or {}).items():
            r += 1
            put(ws5, r, 1, names[key], f_bold)
            put(ws5, r, 2, k, align=centre)
            put(ws5, r, 3, s.get("title", ""))
            put(ws5, r, 4, str(s.get("date", "")), align=centre)
            put(ws5, r, 5, s.get("url", ""), link=s.get("url") or None)
            band(ws5, r, len(h5), gi % 2)
    finish(ws5, hr5, len(h5), "C")

    wb.active = 0
    for w in wb.worksheets:
        w.sheet_view.zoomScale = 100
    ws.sheet_properties.tabColor = NAVY
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
