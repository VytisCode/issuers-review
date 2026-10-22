"""Generuoja ataskaitų failus (Word, PDF) iš reports/<emitentas>/<laikotarpis>.yaml ir data/history/*.csv.

Šaltiniai pateikiami kaip išnašos: tekste mažas numeris (nuoroda tiesiai į šaltinį, PDF - į puslapį),
pabaigoje - numeruotas šaltinių sąrašas su nuorodomis. ESG klausimynas - tiksliai pagal config/seb_policies.yaml esg_template.
Kalbos ir formatai - config/settings.yaml. Rezultatas: docs/files/<emitentas>/<emitentas>_<laikotarpis>_<kalba>.<docx|pdf>
Naudojimas: python tools/build_reports.py            - tik pasikeitusios ataskaitos
            python tools/build_reports.py ignitis    - tik vienas emitentas
            python tools/build_reports.py --force    - visos iš naujo
"""
import hashlib
import sys
import tempfile
import textwrap
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from docx import Document  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.opc.constants import RELATIONSHIP_TYPE as RT  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt, RGBColor  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle  # noqa: E402
from reportlab.lib.units import cm  # noqa: E402
from reportlab.pdfbase import pdfmetrics  # noqa: E402
from reportlab.pdfbase.ttfonts import TTFont  # noqa: E402
from reportlab.platypus import Image, Paragraph, Spacer, Table, TableStyle  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

FONT_DIR = C.ROOT / "assets" / "fonts"
pdfmetrics.registerFont(TTFont("DejaVu", str(FONT_DIR / "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", str(FONT_DIR / "DejaVuSans-Bold.ttf")))
pdfmetrics.registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold", italic="DejaVu", boldItalic="DejaVu-Bold")
for _n, _f in (("Carlito", "Regular"), ("Carlito-Bold", "Bold"), ("Carlito-Italic", "Italic"), ("Carlito-BoldItalic", "BoldItalic")):
    pdfmetrics.registerFont(TTFont(_n, str(FONT_DIR / f"Carlito-{_f}.ttf")))
pdfmetrics.registerFontFamily("Carlito", normal="Carlito", bold="Carlito-Bold", italic="Carlito-Italic", boldItalic="Carlito-BoldItalic")
# PDF šriftas - Inter (SIL OFL); Word naudoja Calibri, nes Inter dažnai neįdiegtas skaitytojo kompiuteryje
for _n in ("Inter-Regular", "Inter-Medium", "Inter-SemiBold", "Inter-Bold", "Inter-Italic", "InterDisplay-SemiBold", "InterDisplay-Bold"):
    pdfmetrics.registerFont(TTFont(_n, str(FONT_DIR / f"{_n}.ttf")))
pdfmetrics.registerFontFamily("Inter-Regular", normal="Inter-Regular", bold="Inter-SemiBold", italic="Inter-Italic", boldItalic="Inter-SemiBold")
ACCENT, ACCENT_LIGHT, INK, MUTED, LINE, SERIES = "12355B", "EEF3F8", "1B2430", "667085", "D5DCE4", "2A78D6"
HAIR, PANEL = "E6EAF0", "F4F6F9"
VERDICT_FILL = {"ok": ("E3F3E8", "1B7F3B"), "watch": ("FDF1D8", "9A6400"), "action": ("FDE7E5", "B42318")}
STATUS_KEYS = {"met": "st_met", "partly": "st_partly", "not_met": "st_not_met", "pending": "st_pending", "changed": "st_changed"}
STATUS_FILL = {"met": "E3F3E8", "partly": "FDF1D8", "not_met": "FDE7E5", "pending": "EAF1F8", "changed": "FDF1D8"}
LEVEL_FILL = {"good": "E3F3E8", "mid": "FDF1D8", "weak": "FDE7E5"}
TREND_INK = {"good": "1B7F3B", "bad": "B42318", "flat": "5F6B7A", "none": "5F6B7A"}
TOPIC_INK = {"review": ("FDE7E5", "B42318"), "reitingas": ("EDE7F6", "5B3FA0"), "sandoris": ("FDF1D8", "9A6400"), "obligacijos": ("E6F0FB", "1F5FA8"),
             "dividendai": ("E3F3E8", "1B7F3B"), "ataskaita": ("EAF1F8", "1F4E79"), "kita": ("F0F2F5", "5F6B7A")}
LANG_NAMES = {"lt": "LT", "lv": "LV", "et": "EE", "en": "EN"}


# ---------- pagalbinės ----------
def period_key(p):
    """FY2025 > 6M2025 tais pačiais metais; rikiavimui."""
    return (p[-4:], 1 if p.startswith("FY") else 0, p)


def comparable(p, available):
    """Lyginamasis laikotarpis: tas pats tipas prieš metus (6M2026 -> 6M2025), kitaip - ankstesnis turimas."""
    prev = p[:-4] + str(int(p[-4:]) - 1)
    if prev in available:
        return prev
    earlier = [x for x in sorted(available, key=period_key) if period_key(x) < period_key(p)]
    return earlier[-1] if earlier else None


def highlights(rep, lang):
    hist, M = C.history(rep["issuer"]), C.metrics()
    out = []
    for m in rep.get("highlights", []):
        vals = hist.get(m)
        if not vals:
            continue
        last = sorted(vals, key=period_key)[-1]
        r, unit = vals[last], vals[last]["unit"]
        prev = comparable(last, vals)
        delta = ""
        if prev:
            a, b = float(vals[prev]["value"]), float(r["value"])
            if unit in ("%", "x"):
                d = b - a
                delta = f"{'▲' if d > 0 else '▼' if d < 0 else '•'} {C.fmt_num(f'{abs(d):.2f}'.rstrip('0').rstrip('.'), 'x' if unit == 'x' else '', lang)}{' p. p.' if unit == '%' else ''} vs {prev}"
            elif a:
                d = (b - a) / abs(a) * 100
                delta = f"{'▲' if d > 0 else '▼' if d < 0 else '•'} {C.fmt_num(f'{abs(d):.1f}', '', lang)} % vs {prev}"
        ul = C.unit_label(unit, lang)
        out.append((C.tr(M.get(m, m), lang), C.fmt_num(r["value"], unit, lang) + (f" {ul}" if ul else ""), delta, last))
    return out


def chart_png(rep, lang, path):
    """Mažos diagramos (viena serija kiekvienoje, be dviejų ašių): metiniai ir tarpiniai laikotarpiai atskirti."""
    hist, M = C.history(rep["issuer"]), C.metrics()
    ms = [m for m in rep.get("charts", []) if hist.get(m)][:4]
    if not ms:
        return None
    from matplotlib import font_manager
    for f in ("Inter-Regular.ttf", "Inter-SemiBold.ttf"):
        font_manager.fontManager.addfont(str(FONT_DIR / f))
    plt.rcParams.update({"font.family": ["Inter", "DejaVu Sans"], "font.size": 8})
    fig, axes = plt.subplots(1, len(ms), figsize=(1.9 * len(ms) + 0.6, 2.3), dpi=200)
    axes = axes if len(ms) > 1 else [axes]
    for ax, m in zip(axes, ms):
        vals = hist[m]
        ps = [p for p in rep.get("periods", sorted(vals, key=period_key)) if p in vals]
        unit = vals[ps[0]]["unit"]
        ys = [float(vals[p]["value"]) for p in ps]
        xs, x, prevfy = [], 0, None
        for p in ps:
            fy = p.startswith("FY")
            if prevfy is not None and fy != prevfy:
                x += 0.6  # tarpas tarp metinių ir tarpinių
            xs.append(x)
            x += 1
            prevfy = fy
        # naujausias laikotarpis - tamsiai mėlynas, ankstesni - šviesesni
        last = max(range(len(ps)), key=lambda i: period_key(ps[i]))
        ax.bar(xs, ys, width=0.62, color=["#" + (ACCENT if i == last else "8DB4E2") for i in range(len(ps))], zorder=2)
        for i, (xi, yi) in enumerate(zip(xs, ys)):
            ax.text(xi, yi, C.fmt_num(vals[ps[i]]["value"], unit, lang), ha="center", va="bottom", fontsize=7.5,
                    color="#" + (ACCENT if i == last else INK), fontweight="bold" if i == last else "normal")
        ax.set_xticks(xs, ps, fontsize=7.5, color="#" + MUTED)
        ul = C.unit_label(unit, lang)
        ax.set_title(textwrap.fill(C.tr(M.get(m, m), lang) + (f", {ul}" if ul else ""), 26), fontsize=8.5, color="#" + ACCENT, loc="left", fontweight="bold", pad=8)
        ax.set_yticks([])
        for s in ("top", "right", "left"):
            ax.spines[s].set_visible(False)
        ax.spines["bottom"].set_color("#" + LINE)
        ax.tick_params(length=0)
        lo = min(0, min(ys))
        ax.set_ylim(lo, max(ys) * 1.18 if max(ys) > 0 else 1)
    fig.tight_layout(w_pad=1.2)
    fig.savefig(path, transparent=False, facecolor="white")
    plt.close(fig)
    return path


# ---------- dokumento modelis (bendras Word ir PDF) ----------
def model(rep, lang, tmpdir):
    L, M = C.i18n(), C.metrics()
    t = lambda k: C.tr(L.get(k, k), lang)  # noqa: E731
    cites = C.Citations(rep.get("sources", {}))
    rich = lambda x, lead=False: (C.lead_bold(cites.parse(C.tr(x, lang))) if lead else cites.parse(C.tr(x, lang)))  # noqa: E731
    iss = C.issuers()[rep["issuer"]]
    hist = C.history(rep["issuer"])
    secs = [s for s in C.securities() if s["issuer"] == rep["issuer"] and s.get("held") == "yes" and s["status"] == "active"]

    vst = (rep.get("verdict") or {}).get("status", "")
    country = t("country_" + iss.get("country", "")) if iss.get("country") else ""
    sub = " · ".join(x for x in (country, iss.get("sector", "") if lang == "en" else "") if x)
    blocks = [("cover", iss["name"], f"{t('title')} · {rep['period']}", vst, t("verdict_" + vst) if vst else "", sub)]

    # pirmo puslapio šoninė juosta: pagrindinė informacija ir turimi VP
    held = []
    for s in secs:
        mat = s["maturity_date"] or (t("perpetual") if s["type"] == "bond" else t("share"))
        cpn = C.fmt_num(s["coupon_pct"], "%", lang) if s["coupon_pct"] else ""
        er = "" if s["type"] == "share" else (s.get("issue_rating") or t("not_rated"))
        held.append((s["isin"], s["name"], " · ".join(x for x in (cpn, mat) if x), er))
    blocks.append(("sidebar", t("key_facts"), [(t("rating"), rich(rep.get("rating", ""))), (t("analysis_date"), [("t", str(rep["analysis_date"]), False)]),
                                               (t("next_report"), rich(rep.get("next_report", ""))), (t("country"), [("t", country, False)] if country else [])],
                   t("held_short"), held))
    if rep.get("verdict"):
        v = rep["verdict"]
        blocks.append(("verdict", v.get("status", "ok"), rich(v)))
    hl = highlights(rep, lang)
    if hl:
        blocks.append(("tiles", hl))

    blocks += [("h1", t("quick_summary")), ("bullets", [rich(x, True) for x in rep.get("quick_summary", [])]), ("endfirst",)]

    # priemonės
    rows = []
    for s in secs:
        mat = s["maturity_date"] or (t("perpetual") if s["type"] == "bond" else t("share"))
        cpn = C.fmt_num(s["coupon_pct"], "%", lang) if s["coupon_pct"] else "–"
        note = ("⚠ " if s.get("verified") == "review" else "") + s.get("status_note", "")
        er = "" if s["type"] == "share" else (s.get("issue_rating") or t("not_rated")) + (f" ({s['issue_rating_note']})" if s.get("issue_rating_note") else "")
        rows.append([s["isin"], s["name"], cpn, mat, er, note])
    blocks += [("h2", t("instruments")), ("table", [t("isin"), t("instrument"), t("coupon"), t("maturity"), t("issue_rating"), t("note")], rows,
                                          [2.6, 4.0, 1.9, 2.0, 2.8, 3.9], {"num": [2]})]

    # rodikliai: šaltiniai - išnašos prie rodiklio pavadinimo
    periods = rep.get("periods", [])
    ratios = [m for m in rep.get("ratios", []) if hist.get(m)]
    # tuščių stulpelių nerodome: laikotarpis rodomas, jei turi reikšmę bent pusei rodiklių
    shown = [p for p in periods if sum(1 for m in ratios if p in hist[m]) * 2 >= len(ratios)]
    rows = []
    for m in ratios:
        vals = hist.get(m, {})
        if not vals:
            continue
        unit = next(iter(vals.values()))["unit"]
        ul = C.unit_label(unit, lang)
        label = [("t", C.tr(M.get(m, m), lang) + (f", {ul}" if ul else ""), False)]
        seen = []
        for p in shown:
            r = vals.get(p)
            if r and r["source"] in cites.sources and (r["source"], r.get("page", "")) not in seen:
                seen.append((r["source"], r.get("page", "")))
        for src, pg in seen[:3]:
            label.append(("c", cites.add(src, pg), cites.url(src, pg)))
        cells = [label]
        last, prev, ch, trend, level = C.ratio_change(m, vals, periods, lang)
        for p in shown:
            r = vals.get(p)
            v = "–" if not r else C.fmt_num(r["value"], unit, lang) + ("*" if r.get("verified") == "check" else "ᶜ" if r.get("verified") == "calc" else "")
            if p == last and level:
                v = ("styled", v, LEVEL_FILL[level], None, True)
            cells.append(v)
        cells.append(("styled", f"{ch} vs {prev}" if ch else "–", None, TREND_INK[trend], trend in ("good", "bad")))
        rows.append(cells)
    blocks += [("h2", t("key_ratios")), ("table", [t("metric")] + shown + [t("change")], rows, [5.2] + [2.1] * len(shown) + [3.9],
                                         {"num": list(range(1, len(shown) + 2))}),
               ("small", t("legend_colors")), ("small", t("legend"))]
    png = chart_png(rep, lang, Path(tmpdir) / f"{rep['issuer']}_{lang}.png")
    if png:
        blocks.append(("image", png))

    for key in ("business", "financials"):
        blocks += [("h2", t(key)), ("bullets", [rich(x, True) for x in rep.get(key, [])])]

    rows = [[rich(p["plan"]), rich(C.tr(p["outcome"], lang) + (f" [{p['src']}]" if p.get("src") else "")), rich(p.get("why", "")) or "–",
             ("status", p["status"], t(STATUS_KEYS.get(p["status"], p["status"])))]
            for p in rep.get("plans", [])]
    blocks += [("h2", t("plans")), ("table", [t("plan"), t("outcome"), t("why"), t("status")], rows, [4.6, 4.6, 5.6, 2.4], {})]

    blocks += [("h2", t("investor")), ("bullets", [rich(x, True) for x in rep.get("investor", [])])]

    # rinka ir aplinka: analitiko punktai (kodėl keitėsi rodikliai, aplinkos veiksniai) + automatinė ECB palūkanų lentelė
    mk = C.market_table()
    if rep.get("market") or mk:
        blocks.append(("h2", t("market")))
        if rep.get("market"):
            blocks.append(("bullets", [rich(x, True) for x in rep["market"]]))
        if mk:
            pct = lambda v: "–" if v is None else C.fmt_num(f"{v:.2f}", "%", lang)  # noqa: E731
            rows = []
            for sid, last, now, m3, m12, src in mk:
                ch = None if now is None or m12 is None else now - m12
                chs = "–" if ch is None else f"{'▲' if ch > 0.005 else '▼' if ch < -0.005 else '•'} {C.fmt_num(f'{abs(ch):.2f}', '', lang)} p. p."
                rows.append([[("t", t("mk_" + sid), False)], f"{pct(now)} ({last})", pct(m3), pct(m12), chs])
            blocks += [("table", ["", t("market_now"), t("market_3m"), t("market_12m"), t("market_chg")], rows, [6.2, 3.2, 2.4, 2.4, 2.8], {"num": [1, 2, 3, 4]}),
                       ("small", t("market_note"))]

    # spauda: analitiko santrauka (jei yra) + automatinė naujienų suvestinė iš data/news.csv
    dig, start, end = C.news_digest(rep["issuer"], days=90, lang=lang)
    blocks.append(("h2", t("press")))
    if rep.get("press"):
        blocks.append(("bullets", [rich(x, True) for x in rep["press"]]))
    langs = ", ".join(LANG_NAMES.get(x, x) for x in sorted({d["lang"] for d in dig})) or "–"
    n = sum(1 + d["more"] for d in dig)
    blocks.append(("lead", t("press_intro").format(langs=langs, start=start, end=end, n=n) if dig else t("press_none")))
    if dig:
        rows = [(d["date"], d["topic"], t("topic_" + d["topic"]), d["title"], d["url"], d["source"],
                 C.tr(L.get("press_more"), lang).format(n=d["more"]) if d["more"] else "") for d in dig[:10]]
        blocks.append(("news", rows))

    # ESG - tiksliai pagal šabloną (klausimai anglų k., nekeičiami), be lentelės
    tpl = C.seb()["esg_template"]
    esg = rep.get("esg", {})
    items, comments = C.esg_items(rep), esg.get("comments", {})
    blocks.append(("h2", tpl["title"]))
    for sec in tpl["sections"]:
        blocks.append(("h3", sec["heading"]))
        for q in sec["items"]:
            it = items.get(q["id"], {"status": "review"})
            st = it.get("status", "review")
            pts = [(lab, [rich(x) for x in items]) for lab, items in C.points(C.tr(it.get("evidence", ""), lang))]
            if st == "review":
                pts.insert(0, (t("needs_review").strip().rstrip("."), []))
            blocks.append(("check", st == "yes", q["q"], pts))
        if sec.get("note"):
            blocks.append(("note", sec["note"]))
        com = C.tr(comments.get(sec.get("comments", ""), ""), lang)
        blocks.append(("comment", tpl["comments_label"], rich(com) if com else []))
    blocks.append(("comment", tpl["conclusion"], rich(esg.get("conclusion", ""))))
    pol = [C.seb()["policies"][k]["name"] for k in esg.get("policies", []) if k in C.seb()["policies"]]
    if pol:
        blocks.append(("small", t("seb_applied") + "; ".join(pol)))

    blocks += [("h2", t("sources")), ("notes", cites.notes()), ("small", t("disclaimer"))]
    return blocks


# ---------- Word ----------
def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), str(v))
    return e


TC_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"]
TBL_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing",
             "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]
P_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
           "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
           "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
           "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]


def put(parent, el, order):
    """Įdeda elementą į schemos nurodytą vietą (Word griežtai tikrina elementų tvarką), pakeisdamas tokį patį."""
    name = el.tag.split("}")[1]
    old = parent.find(qn("w:" + name))
    if old is not None:
        parent.remove(old)
    after = order[order.index(name) + 1:]
    for i, child in enumerate(parent):
        if child.tag.split("}")[1] in after:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def shade(cell, hexcolor):
    put(cell._tc.get_or_add_tcPr(), _el("w:shd", val="clear", color="auto", fill=hexcolor), TC_ORDER)


def hyperlink(par, url, text, size=None, sup=False, color=SERIES):
    rid = par.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), rid)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color)
    rpr.append(c)
    if sup:
        v = OxmlElement("w:vertAlign")
        v.set(qn("w:val"), "superscript")
        rpr.append(v)
    if size:
        s = OxmlElement("w:sz")
        s.set(qn("w:val"), str(int(size * 2)))
        rpr.append(s)
    r.append(rpr)
    tx = OxmlElement("w:t")
    tx.text = text
    tx.set(qn("xml:space"), "preserve")
    r.append(tx)
    h.append(r)
    par._p.append(h)


def add_segs(par, segs, size=None, color=None):
    prev_c = False
    for kind, a, b in segs:
        if kind == "t":
            run = par.add_run(a)
            run.bold = b
            if size:
                run.font.size = Pt(size)
            if color:
                run.font.color.rgb = RGBColor.from_string(color)
            prev_c = False
        else:
            if prev_c:
                r = par.add_run(",")
                r.font.superscript = True
                r.font.color.rgb = RGBColor.from_string(SERIES)
            if b:
                hyperlink(par, b, str(a), size=(size or 10) * 0.85, sup=True)
            else:
                r = par.add_run(str(a))
                r.font.superscript = True
            prev_c = True


def cell_segs(v):
    if isinstance(v, list):
        return v
    return [("t", str(v), False)]


def cell_borders(cell, **edges):
    """cell_borders(c, left=(sz, "1F4E79"), bottom=(4, "C9D3DE"), top=None ...) - None = be linijos."""
    b = _el("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        v = edges.get(edge)
        b.append(_el(f"w:{edge}", val="single", sz=v[0], color=v[1], space=0) if v else _el(f"w:{edge}", val="nil"))
    put(cell._tc.get_or_add_tcPr(), b, TC_ORDER)


def table_setup(table, widths_cm, margins=(60, 100, 60, 100), lines=None):
    """Lentelė be standartinių rėmelių, fiksuoti stulpelių pločiai, langelių paraštės (twips: viršus, kairė, apačia, dešinė)."""
    tblPr = table._tbl.tblPr
    put(tblPr, _el("w:tblW", w=int(sum(widths_cm) * 567), type="dxa"), TBL_ORDER)
    put(tblPr, _el("w:tblLayout", type="fixed"), TBL_ORDER)
    bd = _el("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if lines and edge in ("insideH", "bottom"):
            bd.append(_el(f"w:{edge}", val="single", sz=4, color=lines, space=0))
        else:
            bd.append(_el(f"w:{edge}", val="nil"))
    put(tblPr, bd, TBL_ORDER)
    mar = _el("w:tblCellMar")
    for edge, v in zip(("top", "left", "bottom", "right"), margins):
        mar.append(_el(f"w:{edge}", w=v, type="dxa"))
    put(tblPr, mar, TBL_ORDER)
    grid = table._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths_cm):
        gc.set(qn("w:w"), str(int(w * 567)))
    for row in table.rows:
        for c, w in zip(row.cells, widths_cm):
            c.width = Cm(w)


def para_border(p, edge="bottom", sz=6, color=LINE, space=3):
    bd = _el("w:pBdr")
    bd.append(_el(f"w:{edge}", val="single", sz=sz, color=color, space=space))
    put(p._p.get_or_add_pPr(), bd, P_ORDER)


def para_shade(p, fill):
    put(p._p.get_or_add_pPr(), _el("w:shd", val="clear", color="auto", fill=fill), P_ORDER)


def run(p, text, size=None, bold=None, color=None, italic=None, caps=False, font=None):
    r = p.add_run(text)
    if size:
        r.font.size = Pt(size)
    if bold is not None:
        r.bold = bold
    if italic is not None:
        r.italic = italic
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    if caps:
        r.font.all_caps = True
    if font:
        r.font.name = font
    return r


def tight(p, before=0, after=0, line=None):
    pf = p.paragraph_format
    pf.space_before, pf.space_after = Pt(before), Pt(after)
    if line:
        pf.line_spacing = line
    return p


def page_field(p, size=8, color=ACCENT):
    r = p.add_run()
    r.font.size, r.bold, r.font.color.rgb = Pt(size), True, RGBColor.from_string(color)
    r._r.append(_el("w:fldChar", fldCharType="begin"))
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = " PAGE "
    r._r.append(it)
    r._r.append(_el("w:fldChar", fldCharType="end"))


def to_docx(blocks, path, footer_text):
    W = 17.0
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.8)
    sec.header_distance, sec.footer_distance = Cm(0.9), Cm(0.8)
    sec.different_first_page_header_footer = True
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    st.font.size = Pt(10.5)
    st.font.color.rgb = RGBColor.from_string(INK)
    st.paragraph_format.space_after = Pt(4)
    st.paragraph_format.line_spacing = 1.1
    cover = next((b for b in blocks if b[0] == "cover"), ("cover", "", "", "", ""))
    tab_right = lambda p: p.paragraph_format.tab_stops.add_tab_stop(Cm(W), alignment=2)  # noqa: E731

    # antraštė (nuo 2 psl.) ir poraštė (visuose)
    hp = sec.header.paragraphs[0]
    hp.style = doc.styles["Normal"]  # be „Header“ stiliaus tabuliacijų (centras), kad dešinė pusė būtų lygiai prie krašto
    tab_right(hp)
    run(hp, cover[1], 8, True, ACCENT)
    run(hp, "\t" + cover[2], 8, False, MUTED)
    para_border(hp, "bottom", 4, LINE, 4)
    for ft in (sec.footer, sec.first_page_footer):
        fp = ft.paragraphs[0]
        fp.style = doc.styles["Normal"]
        tab_right(fp)
        para_border(fp, "top", 4, LINE, 4)
        run(fp, footer_text + "\t", 8, False, MUTED)
        page_field(fp)

    sec_no = 0
    for kind, *rest in blocks:
        if kind == "cover":
            bar = tight(doc.add_paragraph(), 0, 10)
            para_border(bar, "top", 24, ACCENT, 0)
            bar.paragraph_format.line_spacing = 0.5
            tb = doc.add_table(rows=1, cols=2)
            table_setup(tb, [12.4, 4.6], margins=(0, 0, 0, 0))
            a, b = tb.rows[0].cells
            run(tight(a.paragraphs[0]), rest[1], 8, True, SERIES, caps=True)
            run(tight(a.add_paragraph(), 2, 0), rest[0], 26 if len(rest[0]) < 28 else 20, True, ACCENT)
            if rest[4]:
                run(tight(a.add_paragraph(), 2, 0), rest[4], 9, False, MUTED)
            bp = tight(b.paragraphs[0], 22, 0)
            bp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if rest[3]:
                fill, ink = VERDICT_FILL.get(rest[2], VERDICT_FILL["ok"])
                r = run(bp, f"  ●  {rest[3].upper()}  ", 9, True, ink)
                r._r.get_or_add_rPr().append(_el("w:shd", val="clear", color="auto", fill=fill))
            ac = tight(doc.add_paragraph(), 6, 10)
            para_border(ac, "top", 4, LINE, 0)
            ac.paragraph_format.line_spacing = 0.5
        elif kind == "sidebar":
            title, facts, held_title, held = rest
            tb = doc.add_table(rows=1, cols=2)
            table_setup(tb, [8.5, 8.5], margins=(160, 220, 160, 220))
            a, b = tb.rows[0].cells
            for c in (a, b):
                shade(c, PANEL)
                cell_borders(c, top=(12, ACCENT))
            run(tight(a.paragraphs[0], 0, 2), title, 9, True, ACCENT)
            for k, segs in facts:
                run(tight(a.add_paragraph(), 4, 0), k, 7, True, MUTED, caps=True)
                p = tight(a.add_paragraph())
                add_segs(p, segs, size=9) if segs else run(p, "–", 9)
            run(tight(b.paragraphs[0], 0, 2), held_title, 9, True, ACCENT)
            for isin, name, terms, er in held:
                run(tight(b.add_paragraph(), 4, 0), name, 8.5, True, INK)
                run(tight(b.add_paragraph()), isin + (f"  ·  {terms}" if terms else ""), 7.5, False, MUTED)
                if er:
                    run(tight(b.add_paragraph()), er, 7.5, False, ACCENT)
            tight(doc.add_paragraph(), 0, 6)
        elif kind == "h1":
            p = tight(doc.add_paragraph(), 10, 5)
            p.paragraph_format.keep_with_next = True
            run(p, rest[0], 13, True, ACCENT)
        elif kind == "endfirst":
            continue
        elif kind == "verdict":
            fill, ink = VERDICT_FILL.get(rest[0], VERDICT_FILL["ok"])
            tb = doc.add_table(rows=1, cols=1)
            table_setup(tb, [W], margins=(180, 260, 200, 260))
            c = tb.rows[0].cells[0]
            shade(c, fill)
            cell_borders(c, left=(24, ink))
            add_segs(tight(c.paragraphs[0], 0, 0, 1.15), rest[1], size=10.5, color=INK)
            tight(doc.add_paragraph(), 0, 6)
        elif kind == "tiles":
            tiles, gap = rest[0], 0.35
            tw = (W - gap * (len(tiles) - 1)) / len(tiles)
            tb = doc.add_table(rows=1, cols=len(tiles) * 2 - 1)
            table_setup(tb, [x for i in range(len(tiles)) for x in ([tw, gap] if i < len(tiles) - 1 else [tw])], margins=(150, 170, 150, 100))
            for i, (label, val, delta, per) in enumerate(tiles):
                c = tb.rows[0].cells[i * 2]
                cell_borders(c, top=(18, SERIES), left=(4, LINE), right=(4, LINE), bottom=(4, LINE))
                p = tight(c.paragraphs[0], 0, 3)
                run(p, label, 7, True, MUTED, caps=True)
                run(p, f"  ·  {per}", 7, False, MUTED)
                num, unit = tile_value(val)
                p = tight(c.add_paragraph(), 0, 1)
                run(p, num, 17, True, ACCENT)
                if unit:
                    run(p, " " + unit, 9, False, MUTED)
                run(tight(c.add_paragraph()), delta or " ", 7.5, False, MUTED)
            tight(doc.add_paragraph(), 0, 2)
        elif kind == "image":
            doc.add_picture(str(rest[0]), width=Cm(W))
        elif kind == "h2":
            sec_no += 1
            p = tight(doc.add_paragraph(), 18, 8)
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.tab_stops.add_tab_stop(Cm(0.85))
            run(p, f"{sec_no:02d}\t", 9, True, SERIES)
            run(p, rest[0], 14, True, ACCENT)
            para_border(p, "bottom", 8, ACCENT, 4)
        elif kind == "h3":
            p = tight(doc.add_paragraph(), 10, 4)
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.left_indent = Cm(0.15)
            para_shade(p, ACCENT_LIGHT)
            run(p, rest[0], 9.5, True, ACCENT, caps=True)
        elif kind in ("small", "lead"):
            p = tight(doc.add_paragraph(), 2, 3)
            run(p, rest[0], 8 if kind == "small" else 9, False, MUTED)
        elif kind == "bullets":
            for segs in rest[0]:
                p = tight(doc.add_paragraph(), 0, 4)
                pf = p.paragraph_format
                pf.left_indent, pf.first_line_indent = Cm(0.45), Cm(-0.45)
                pf.tab_stops.add_tab_stop(Cm(0.45))
                run(p, "▪\t", 10, True, SERIES)
                add_segs(p, segs)
        elif kind == "check":
            ok, q, pts = rest
            tb = doc.add_table(rows=1, cols=2)
            table_setup(tb, [0.6, W - 0.6], margins=(90, 0, 110, 0))
            m, body = tb.rows[0].cells
            for c in (m, body):
                cell_borders(c, bottom=(4, "E3E8EE"))
            run(tight(m.paragraphs[0]), "☑" if ok else "☐", 12, False, "1B7F3B" if ok else "9AA5B1", font="Segoe UI Symbol")
            run(tight(body.paragraphs[0], 0, 2), q, 9.5, True, INK)
            for lab, items in pts:
                if lab:
                    run(tight(body.add_paragraph(), 2, 0), lab, 9, True, "B42318" if not items else "333333")
                for segs in items:
                    p = tight(body.add_paragraph(), 0, 1)
                    pf = p.paragraph_format
                    pf.left_indent, pf.first_line_indent = Cm(0.35), Cm(-0.35)
                    pf.tab_stops.add_tab_stop(Cm(0.35))
                    run(p, "•\t", 9, False, MUTED)
                    add_segs(p, segs, size=9, color="333333")
        elif kind == "note":
            run(tight(doc.add_paragraph(), 6, 2), rest[0], 9, italic=True, color="333333")
        elif kind == "comment":
            p = tight(doc.add_paragraph(), 6, 2)
            run(p, rest[0] + " ", 9.5, True)
            add_segs(p, rest[1], size=9.5)
        elif kind == "news":
            tb = doc.add_table(rows=0, cols=3)
            for date, topic, tlabel, title, url, source, more in rest[0]:
                cells = tb.add_row().cells
                for c in cells:
                    cell_borders(c, bottom=(4, "E3E8EE"))
                run(tight(cells[0].paragraphs[0]), date, 8.5, False, INK)
                fill, ink = TOPIC_INK.get(topic, TOPIC_INK["kita"])
                p = tight(cells[1].paragraphs[0])
                r = run(p, f" {tlabel.upper()} ", 7, True, ink)
                r._r.get_or_add_rPr().append(_el("w:shd", val="clear", color="auto", fill=fill))
                p = tight(cells[2].paragraphs[0])
                hyperlink(p, url, title, size=9, color=INK) if url else run(p, title, 9)
                run(p, f"   {source}" + (f"  ·  {more}" if more else ""), 7.5, False, MUTED)
            table_setup(tb, [1.8, 2.4, W - 4.2], margins=(80, 0, 80, 120))
            tight(doc.add_paragraph(), 0, 2)
        elif kind == "notes":
            for n, title, pg, note, url, key, date in rest[0]:
                p = tight(doc.add_paragraph(), 0, 1.5)
                pf = p.paragraph_format
                pf.left_indent, pf.first_line_indent = Cm(0.7), Cm(-0.7)
                pf.tab_stops.add_tab_stop(Cm(0.7))
                run(p, f"{n}.\t", 8.5, False, MUTED)
                desc = title + (f", p. {pg}" if pg else "") + (f" ({note})" if note else "") + (f", {date}" if date else "")
                hyperlink(p, url, desc, size=8.5) if url else run(p, desc, 8.5)
        elif kind == "table":
            header, rows, widths, opts = rest
            num = set(opts.get("num", []))
            widths = [w * W / sum(widths) for w in widths]
            tbl = doc.add_table(rows=1, cols=len(header))
            for i, h in enumerate(header):
                c = tbl.rows[0].cells[i]
                cell_borders(c, bottom=(8, ACCENT))
                p = tight(c.paragraphs[0])
                run(p, str(h), 7, True, MUTED, caps=True)
                if i in num:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for row in rows:
                cells = tbl.add_row().cells
                for i, v in enumerate(row):
                    cell_borders(cells[i], bottom=(4, HAIR))
                    par = tight(cells[i].paragraphs[0])
                    if i in num:
                        par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    if isinstance(v, tuple) and v[0] == "status":
                        shade(cells[i], STATUS_FILL.get(v[1], "FFFFFF"))
                        run(par, v[2], 8.5, True)
                    elif isinstance(v, tuple) and v[0] == "styled":
                        if v[2]:
                            shade(cells[i], v[2])
                        run(par, v[1], 8.5, v[4], v[3])
                    else:
                        add_segs(par, cell_segs(v), size=8.5)
            table_setup(tbl, widths, margins=(70, 100, 70, 100))
            tight(doc.add_paragraph(), 0, 2)
    doc.save(path)


# ---------- PDF ----------
def _cmap(path):
    from fontTools.ttLib import TTFont as _FT
    return set(_FT(str(path)).getBestCmap())


BODY_GLYPHS = _cmap(FONT_DIR / "Inter-Regular.ttf")


def esc(s):
    """XML tekstas; simboliai, kurių nėra Inter šrifte (▲ ☑ ⚠ ...), rašomi DejaVu šriftu."""
    out = []
    for ch in str(s):
        e = {"&": "&amp;", "<": "&lt;", ">": "&gt;"}.get(ch, ch)
        out.append(e if ord(ch) in BODY_GLYPHS or ch in "\n\t" else f'<font name="DejaVu">{e}</font>')
    return "".join(out)


def segs_html(segs):
    out, prev_c = [], False
    for kind, a, b in segs:
        if kind == "t":
            out.append(f"<b>{esc(a)}</b>" if b else esc(a))
            prev_c = False
        else:
            sep = "," if prev_c else ""
            link = f'<a href="{esc(b)}" color="#{SERIES}">{a}</a>' if b else str(a)
            out.append(f'<super><font color="#{SERIES}">{sep}</font>{link}</super>')
            prev_c = True
    return "".join(out)


def tile_value(v):
    """„1 925,3 mln. EUR“ -> didelis skaičius ir mažas vienetas (kad tilptų vienoje eilutėje)."""
    i = 0
    while i < len(v) and (v[i].isdigit() or v[i] in " ,.-−  "):
        i += 1
    num, unit = v[:i].strip(), v[i:].strip()
    if num.endswith(",") or num.endswith("."):
        num = num[:-1]
    return num or v, unit if num else ""


def to_pdf(blocks, path, footer_text):
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.platypus import BaseDocTemplate, Flowable, Frame, FrameBreak, KeepInFrame, NextPageTemplate, PageTemplate
    hexc = lambda h: colors.HexColor("#" + h)  # noqa: E731
    R, MD, SB = "Inter-Regular", "Inter-Medium", "Inter-SemiBold"
    S = lambda name, **kw: ParagraphStyle(name, fontName=kw.pop("font", R), textColor=hexc(kw.pop("color", INK)), **kw)  # noqa: E731
    st = {"h1": S("h1", font="InterDisplay-SemiBold", fontSize=12.5, leading=16, color=ACCENT, spaceBefore=4, spaceAfter=6),
          "h2": S("h", font="InterDisplay-SemiBold", fontSize=13.5, leading=17, color=ACCENT),
          "h2n": S("hn", font=SB, fontSize=8, leading=17, color=SERIES),
          "h3": S("h3", font=SB, fontSize=8, leading=11, color=ACCENT),
          "lead": S("ld", fontSize=8.5, leading=12, color=MUTED, spaceAfter=4),
          "b": S("b", fontSize=9.2, leading=13.4, leftIndent=12, bulletIndent=0, spaceAfter=4.5, bulletFontName=SB, bulletFontSize=7, bulletColor=hexc(SERIES)),
          "b1": S("b1", fontSize=8.9, leading=12.8, leftIndent=11, bulletIndent=0, spaceAfter=4, bulletFontName=SB, bulletFontSize=7, bulletColor=hexc(SERIES)),
          "q": S("q", font=SB, fontSize=8.8, leading=12),
          "verdict": S("v", font=MD, fontSize=9.6, leading=14),
          "evh": S("eh", font=SB, fontSize=8.3, leading=11.5, spaceBefore=2, color="344054"),
          "evb": S("eb", fontSize=8.3, leading=11.5, leftIndent=9, bulletIndent=0, color="344054", bulletColor=hexc(MUTED)),
          "note": S("n", fontSize=8.5, leading=12, spaceBefore=6, color="344054"),
          "small": S("s", fontSize=7, leading=9.5, color=MUTED, spaceBefore=3),
          "srcn": S("srcn", fontSize=7.4, leading=10, color=MUTED, alignment=2), "src": S("src", fontSize=7.4, leading=10),
          "cell": S("cell", fontSize=7.9, leading=10.4), "celln": S("celln", fontSize=7.9, leading=10.4, alignment=2),
          "cellh": S("cellh", font=SB, fontSize=6.6, leading=9, color=MUTED),
          "cellhn": S("cellhn", font=SB, fontSize=6.6, leading=9, color=MUTED, alignment=2),
          "tl": S("tl", font=SB, fontSize=6.4, leading=8.5, color=MUTED),
          "tv": S("tv", font="InterDisplay-SemiBold", fontSize=16, leading=19, color=ACCENT),
          "td": S("td", fontSize=7, leading=9, color=MUTED),
          "chip": S("chip", font=SB, fontSize=6.3, leading=8, alignment=1),
          "sk": S("sk", font=SB, fontSize=6.4, leading=8.5, color=MUTED, spaceBefore=7),
          "sv": S("sv", fontSize=8.2, leading=11.2),
          "sh": S("sh", font=SB, fontSize=8.5, leading=11, color=ACCENT),
          "si": S("si", font=SB, fontSize=7.6, leading=10, spaceBefore=6), "ss": S("ss", fontSize=7.2, leading=9.6, color=MUTED)}
    W, PW, PH, M = 17.0 * cm, *A4, 2.0 * cm
    SIDE, GAP = 5.4 * cm, 0.7 * cm
    MAIN = W - SIDE - GAP
    HEAD = 4.3 * cm  # pirmo puslapio antraštės aukštis nuo viršaus
    cover = next((b for b in blocks if b[0] == "cover"), ("cover", "", "", "", "", ""))
    flow, side, sec_no, width = [], [], 0, MAIN

    def boxed(content, style, w=None):
        tb = Table([[content]], colWidths=[w or width])
        tb.setStyle(TableStyle(style))
        return tb

    class FirstPageEnd(Flowable):
        """Po santraukos - naujas puslapis, jei dar esame pirmame (santraukai užėjus į 2 psl., lūžio nedarome)."""
        def wrap(self, aw, ah):
            self._h = ah + 1 if self.canv.getPageNumber() == 1 else 0
            return aw, self._h

        def split(self, aw, ah):
            return []

        def draw(self):
            pass

    for kind, *rest in blocks:
        if kind == "cover":
            continue  # piešiama pirmo puslapio antraštėje
        if kind == "sidebar":
            title, facts, held_title, held = rest
            side.append(Paragraph(esc(title), st["sh"]))
            for k, segs in facts:
                side += [Paragraph(esc(k).upper(), st["sk"]), Paragraph(segs_html(segs) or "–", st["sv"])]
            if held:
                side += [Spacer(1, 12), Paragraph(esc(held_title), st["sh"])]
                for isin, name, terms, er in held:
                    side.append(Paragraph(esc(name), st["si"]))
                    side.append(Paragraph(esc(isin) + (f"  ·  {esc(terms)}" if terms else ""), st["ss"]))
                    if er:
                        side.append(Paragraph(f'<font color="#{ACCENT}">{esc(er)}</font>', st["ss"]))
        elif kind == "verdict":
            fill, ink = VERDICT_FILL.get(rest[0], VERDICT_FILL["ok"])
            flow += [boxed(Paragraph(segs_html(rest[1]), st["verdict"]),
                           [("BACKGROUND", (0, 0), (-1, -1), hexc(fill)), ("LINEBEFORE", (0, 0), (0, 0), 3, hexc(ink)),
                            ("TOPPADDING", (0, 0), (-1, -1), 10), ("BOTTOMPADDING", (0, 0), (-1, -1), 11),
                            ("LEFTPADDING", (0, 0), (-1, -1), 13), ("RIGHTPADDING", (0, 0), (-1, -1), 12)]), Spacer(1, 12)]
        elif kind == "tiles":
            tiles, gap = rest[0], 0.3 * cm
            per_row = 2 if width < 14 * cm else len(tiles)
            tw = (width - gap * (per_row - 1)) / per_row
            grid, style = [], [("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 10),
                               ("RIGHTPADDING", (0, 0), (-1, -1), 6), ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 9)]
            for i, (lb, v, d, per) in enumerate(tiles):
                num, unit = tile_value(v)
                cell = [Paragraph(esc(lb).upper() + f'<font color="#{LINE}">  ·  </font>{esc(per)}', st["tl"]), Spacer(1, 4),
                        Paragraph(esc(num) + (f' <font name="{R}" size="8.5" color="#{MUTED}">{esc(unit)}</font>' if unit else ""), st["tv"]), Spacer(1, 2),
                        Paragraph(esc(d) or "&nbsp;", st["td"])]
                r, c = divmod(i, per_row)
                if c == 0:
                    if grid:
                        grid.append([""] * (per_row * 2 - 1))  # tarpas tarp eilučių
                    grid.append([""] * (per_row * 2 - 1))
                ri, ci = len(grid) - 1, c * 2
                grid[ri][ci] = cell
                style += [("BACKGROUND", (ci, ri), (ci, ri), colors.white), ("BOX", (ci, ri), (ci, ri), 0.5, hexc(LINE)),
                          ("LINEABOVE", (ci, ri), (ci, ri), 2, hexc(SERIES))]
            widths = [tw if j % 2 == 0 else gap for j in range(per_row * 2 - 1)]
            heights = [None if j % 2 == 0 else gap for j in range(len(grid))]
            for j in range(len(grid)):
                if j % 2:
                    style += [("TOPPADDING", (0, j), (-1, j), 0), ("BOTTOMPADDING", (0, j), (-1, j), 0)]
            tb = Table(grid, colWidths=widths, rowHeights=heights)
            tb.setStyle(TableStyle(style))
            flow += [tb, Spacer(1, 8)]
        elif kind == "h1":
            flow.append(Paragraph(esc(rest[0]), st["h1"]))
        elif kind == "endfirst":
            width = W
            flow += [NextPageTemplate("later"), FirstPageEnd()]
        elif kind == "h2":
            sec_no += 1
            tb = Table([[Paragraph(f"{sec_no:02d}", st["h2n"]), Paragraph(esc(rest[0]), st["h2"])]], colWidths=[0.8 * cm, width - 0.8 * cm])
            tb.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
                                    ("LINEBELOW", (0, 0), (-1, -1), 0.6, hexc(LINE)), ("LINEBELOW", (0, 0), (0, 0), 1.6, hexc(SERIES))]))
            gap = Spacer(1, 9)
            tb.keepWithNext = gap.keepWithNext = True
            flow += [Spacer(1, 18), tb, gap]
        elif kind == "h3":
            p = Paragraph(esc(rest[0]).upper(), st["h3"])
            tb = boxed(p, [("LINEBEFORE", (0, 0), (0, 0), 2, hexc(ACCENT)), ("BACKGROUND", (0, 0), (-1, -1), hexc(PANEL)),
                           ("LEFTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 4.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5)])
            gap = Spacer(1, 3)
            tb.keepWithNext = gap.keepWithNext = True
            flow += [Spacer(1, 9), tb, gap]
        elif kind == "image":
            img = Image(str(rest[0]))
            img.drawWidth, img.drawHeight = width, width * img.imageHeight / img.imageWidth
            flow += [Spacer(1, 10), img]
        elif kind == "small":
            flow.append(Paragraph(esc(rest[0]), st["small"]))
        elif kind == "lead":
            flow.append(Paragraph(esc(rest[0]), st["lead"]))
        elif kind == "bullets":
            for segs in rest[0]:
                flow.append(Paragraph(segs_html(segs), st["b1" if width < W else "b"], bulletText="●"))
        elif kind == "check":
            ok, q, pts = rest
            fill, ink, sym = ("E3F3E8", "1B7F3B", "✓") if ok else ("F0F2F5", "98A2B3", "–")
            mark = Table([[Paragraph(f'<font color="#{ink}">{sym}</font>', S("mk", font=SB, fontSize=8, leading=10, alignment=1))]],
                         colWidths=[0.42 * cm], rowHeights=[0.42 * cm])
            mark.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), hexc(fill)), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                      ("TOPPADDING", (0, 0), (-1, -1), 0.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 0), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
            body = [Paragraph(esc(q), st["q"])]
            for lab, items in pts:
                if lab:
                    body.append(Paragraph(f'<font color="#{"B42318" if not items else "344054"}">{esc(lab)}</font>', st["evh"]))
                for segs in items:
                    body.append(Paragraph(segs_html(segs), st["evb"], bulletText="•"))
            tb = Table([[mark, body]], colWidths=[0.7 * cm, width - 0.7 * cm])
            tb.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                    ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                                    ("LINEBELOW", (0, 0), (-1, -1), 0.4, hexc(HAIR))]))
            flow.append(tb)
        elif kind == "note":
            flow.append(Paragraph(f"<i>{esc(rest[0])}</i>", st["note"]))
        elif kind == "comment":
            flow.append(Paragraph(f"<b>{esc(rest[0])}</b> {segs_html(rest[1])}", st["note"]))
        elif kind == "news":
            data, style = [], [("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                               ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("LINEBELOW", (0, 0), (-1, -1), 0.4, hexc(HAIR))]
            for date, topic, tlabel, title, url, source, more in rest[0]:
                fill, ink = TOPIC_INK.get(topic, TOPIC_INK["kita"])
                chip = Table([[Paragraph(f'<font color="#{ink}">{esc(tlabel).upper()}</font>', st["chip"])]], colWidths=[2.1 * cm])
                chip.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), hexc(fill)), ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
                                          ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2)]))
                head = f'<a href="{esc(url)}" color="#{INK}">{esc(title)}</a>'
                tail = f'<font color="#{MUTED}" size="7">  {esc(source)}{"  ·  " + esc(more) if more else ""}</font>'
                data.append([Paragraph(esc(date), S("nd", fontSize=7.6, leading=10.4, color=MUTED)), chip, Paragraph(head + tail, st["cell"])])
            tb = Table(data, colWidths=[1.8 * cm, 2.4 * cm, width - 4.2 * cm])
            tb.setStyle(TableStyle(style))
            flow.append(tb)
        elif kind == "notes":
            data = []
            for n, title, pg, note, url, key, date in rest[0]:
                desc = esc(title + (f", p. {pg}" if pg else "") + (f" ({note})" if note else "") + (f", {date}" if date else ""))
                data.append([Paragraph(f"{n}.", st["srcn"]), Paragraph(f'<a href="{esc(url)}" color="#{SERIES}">{desc}</a>' if url else desc, st["src"])])
            tb = Table(data, colWidths=[0.7 * cm, width - 0.7 * cm])
            tb.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (0, -1), 5),
                                    ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]))
            flow += [tb, Spacer(1, 6)]
        elif kind == "table":
            header, rows, widths, opts = rest
            num = set(opts.get("num", []))
            widths = [w * width / sum(widths) for w in widths]
            data = [[Paragraph(esc(h).upper(), st["cellhn" if i in num else "cellh"]) for i, h in enumerate(header)]]
            style = [("LINEBELOW", (0, 0), (-1, 0), 0.9, hexc(ACCENT)), ("LINEBELOW", (0, 1), (-1, -1), 0.4, hexc(HAIR)),
                     ("VALIGN", (0, 0), (-1, -1), "TOP"), ("VALIGN", (0, 0), (-1, 0), "BOTTOM"),
                     ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5), ("LEFTPADDING", (0, 0), (0, -1), 0),
                     ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, 0), 4)]
            for ri, row in enumerate(rows, start=1):
                line = []
                for i, v in enumerate(row):
                    if isinstance(v, tuple) and v[0] == "status":
                        style.append(("BACKGROUND", (i, ri), (i, ri), hexc(STATUS_FILL.get(v[1], "FFFFFF"))))
                        style.append(("LEFTPADDING", (i, ri), (i, ri), 5))
                        line.append(Paragraph(f"<b>{esc(v[2])}</b>", st["cell"]))
                    elif isinstance(v, tuple) and v[0] == "styled":
                        if v[2]:
                            style.append(("BACKGROUND", (i, ri), (i, ri), hexc(v[2])))
                        txt = esc(v[1])
                        txt = f'<font color="#{v[3]}">{txt}</font>' if v[3] else txt
                        line.append(Paragraph(f"<b>{txt}</b>" if v[4] else txt, st["celln" if i in num else "cell"]))
                    else:
                        line.append(Paragraph(segs_html(cell_segs(v)), st["celln" if i in num else "cell"]))
                data.append(line)
            tb = Table(data, colWidths=widths, repeatRows=1)
            tb.setStyle(TableStyle(style))
            flow += [tb, Spacer(1, 4)]

    def pill(canvas, x_right, y, label, status, size=8):
        fill, ink = VERDICT_FILL.get(status, VERDICT_FILL["ok"])
        canvas.setFont(SB, size)
        w = canvas.stringWidth(label, SB, size) + 0.85 * cm
        x = x_right - w
        canvas.setFillColor(hexc(fill))
        canvas.roundRect(x, y, w, 0.66 * cm, 0.33 * cm, stroke=0, fill=1)
        canvas.setFillColor(hexc(ink))
        canvas.circle(x + 0.36 * cm, y + 0.33 * cm, 0.085 * cm, stroke=0, fill=1)
        canvas.drawString(x + 0.58 * cm, y + 0.235 * cm, label)

    def first(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(hexc(ACCENT))
        canvas.rect(0, PH - 0.32 * cm, PW, 0.32 * cm, stroke=0, fill=1)
        canvas.setFillColor(hexc(SERIES))
        canvas.rect(M, PH - 0.32 * cm - 0.08 * cm, 2.2 * cm, 0.08 * cm, stroke=0, fill=1)
        canvas.setFillColor(hexc(SERIES))
        canvas.setFont(SB, 7.5)
        canvas.drawString(M, PH - 1.45 * cm, cover[2].upper())
        canvas.setFillColor(hexc(ACCENT))
        size = 28
        while size > 18 and canvas.stringWidth(cover[1], "InterDisplay-Bold", size) > 12.2 * cm:
            size -= 1
        canvas.setFont("InterDisplay-Bold", size)
        canvas.drawString(M, PH - 2.55 * cm, cover[1])
        if cover[5]:
            canvas.setFillColor(hexc(MUTED))
            canvas.setFont(R, 8.5)
            canvas.drawString(M, PH - 3.2 * cm, cover[5][:110])
        if cover[4]:
            pill(canvas, PW - M, PH - 2.6 * cm, cover[4].upper(), cover[3])
        canvas.setStrokeColor(hexc(LINE))
        canvas.setLineWidth(0.6)
        canvas.line(M, PH - 3.65 * cm, PW - M, PH - 3.65 * cm)
        # šoninės juostos fonas
        canvas.setFillColor(hexc(PANEL))
        canvas.rect(PW - M - SIDE, M, SIDE, PH - HEAD - M, stroke=0, fill=1)
        canvas.setFillColor(hexc(ACCENT))
        canvas.rect(PW - M - SIDE, PH - HEAD - 0.06 * cm, SIDE, 0.06 * cm, stroke=0, fill=1)
        canvas.restoreState()
        footer(canvas)

    def later(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(hexc(ACCENT))
        canvas.rect(0, PH - 0.16 * cm, PW, 0.16 * cm, stroke=0, fill=1)
        canvas.setFont(SB, 7.5)
        canvas.drawString(M, PH - 1.15 * cm, cover[1])
        canvas.setFont(R, 7.5)
        canvas.setFillColor(hexc(MUTED))
        canvas.drawRightString(PW - M, PH - 1.15 * cm, cover[2])
        canvas.setStrokeColor(hexc(LINE))
        canvas.setLineWidth(0.5)
        canvas.line(M, PH - 1.38 * cm, PW - M, PH - 1.38 * cm)
        canvas.restoreState()
        footer(canvas)

    def footer(canvas):
        canvas.saveState()
        canvas.setStrokeColor(hexc(LINE))
        canvas.setLineWidth(0.5)
        canvas.line(M, 1.45 * cm, PW - M, 1.45 * cm)
        canvas.setFont(R, 6.8)
        canvas.setFillColor(hexc(MUTED))
        canvas.drawString(M, 1.0 * cm, footer_text)
        canvas.restoreState()

    class NumberedCanvas(Canvas):
        """Puslapio numeris „X / Y“ (bendras puslapių skaičius žinomas tik pabaigoje)."""
        def __init__(self, *a, **kw):
            super().__init__(*a, **kw)
            self._pages = []

        def showPage(self):
            self._pages.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            n = len(self._pages)
            for state in self._pages:
                self.__dict__.update(state)
                self.setFont(SB, 7.2)
                self.setFillColor(hexc(ACCENT))
                self.drawRightString(PW - M, 1.0 * cm, f"{self._pageNumber} / {n}")
                super().showPage()
            super().save()

    doc = BaseDocTemplate(str(path), pagesize=A4, leftMargin=M, rightMargin=M, topMargin=M, bottomMargin=M,
                          title=f"{cover[1]} · {cover[2]}", author="issuers-review")
    top = PH - HEAD - 0.45 * cm
    fs = Frame(PW - M - SIDE, M, SIDE, top - M, id="side", leftPadding=12, rightPadding=10, topPadding=4, bottomPadding=8)
    fm = Frame(M, M, MAIN, PH - HEAD - M, id="main", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    f2 = Frame(M, M, W, PH - 2 * M, id="f2", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate("first", [fs, fm], onPage=first, autoNextPageTemplate="later"), PageTemplate("later", [f2], onPage=later)])
    side_box = KeepInFrame(SIDE - 22, top - M - 12, side or [Spacer(1, 1)], mode="shrink")
    doc.build([side_box, FrameBreak()] + flow, canvasmaker=NumberedCanvas)


# ---------- paleidimas ----------
def inputs_hash(issuer, rep):
    """Failai perkuriami tik pasikeitus jų šaltiniams (kitaip kiekvienas paleidimas keistų dvejetainius failus)."""
    h = hashlib.sha1()
    for p in [rep["_file"], C.DATA / "history" / f"{issuer}.csv", C.DATA / "sanctions.json", C.DATA / "market.csv", C.DATA / "news.csv", Path(__file__), Path(C.__file__), *sorted(C.CONFIG.glob("*"))]:
        if Path(p).exists():
            h.update(Path(p).read_bytes())
    return h.hexdigest()


def build(only=None, force=False):
    cfg = C.settings()
    built = []
    state_path = C.DOCS / "files" / "build_state.json"
    state = C.load_json(state_path)
    with tempfile.TemporaryDirectory() as tmp:
        for issuer in C.issuers():
            if only and issuer != only:
                continue
            for rep in C.reports(issuer):
                out = C.DOCS / "files" / issuer
                out.mkdir(parents=True, exist_ok=True)
                for lang in cfg.get("languages", ["lt"]):
                    base = out / f"{issuer}_{rep['period']}_{lang}"
                    hkey, hval = base.name, inputs_hash(issuer, rep)
                    if not force and state.get(hkey) == hval and all(base.with_suffix("." + e).exists() for e in cfg["formats"]):
                        continue
                    state[hkey] = hval
                    blocks = model(rep, lang, tmp)
                    footer = f"{C.issuers()[issuer]['name']} · {rep['period']} · {rep['analysis_date']}"
                    if "docx" in cfg["formats"]:
                        to_docx(blocks, base.with_suffix(".docx"), footer)
                    if "pdf" in cfg["formats"]:
                        to_pdf(blocks, base.with_suffix(".pdf"), footer)
                    built.append(base.name)
    C.save_json(state_path, state)
    print("Sugeneruota:", ", ".join(built) or "nieko (šaltiniai nepasikeitė)")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    build(args[0] if args else None, force="--force" in sys.argv)
