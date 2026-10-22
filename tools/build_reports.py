"""Generuoja ataskaitų failus (Word, PDF, Excel) iš reports/<emitentas>/<laikotarpis>.yaml ir data/history/*.csv.

Šaltiniai pateikiami kaip išnašos: tekste mažas numeris (nuoroda tiesiai į šaltinį, PDF - į puslapį),
pabaigoje - numeruotas šaltinių sąrašas su nuorodomis. ESG klausimynas - tiksliai pagal config/seb_policies.yaml esg_template.
Kalbos ir formatai - config/settings.yaml. Rezultatas: docs/files/<emitentas>/<emitentas>_<laikotarpis>_<kalba>.<docx|pdf|xlsx>
Naudojimas: python tools/build_reports.py            - tik pasikeitusios ataskaitos
            python tools/build_reports.py ignitis    - tik vienas emitentas
            python tools/build_reports.py --force    - visos iš naujo
"""
import hashlib
import sys
import tempfile
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
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle  # noqa: E402
from reportlab.lib.units import cm  # noqa: E402
from reportlab.pdfbase import pdfmetrics  # noqa: E402
from reportlab.pdfbase.ttfonts import TTFont  # noqa: E402
from reportlab.platypus import Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

FONT_DIR = C.ROOT / "assets" / "fonts"
pdfmetrics.registerFont(TTFont("DejaVu", str(FONT_DIR / "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", str(FONT_DIR / "DejaVuSans-Bold.ttf")))
pdfmetrics.registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold", italic="DejaVu", boldItalic="DejaVu-Bold")
ACCENT, ACCENT_LIGHT, INK, MUTED, LINE, SERIES = "1F4E79", "EAF1F8", "1A1A1A", "5F6B7A", "C9D3DE", "2A78D6"
STATUS_KEYS = {"met": "st_met", "partly": "st_partly", "not_met": "st_not_met", "pending": "st_pending", "changed": "st_changed"}
STATUS_FILL = {"met": "E3F3E8", "partly": "FDF1D8", "not_met": "FDE7E5", "pending": "EAF1F8", "changed": "FDF1D8"}
LEVEL_FILL = {"good": "E3F3E8", "mid": "FDF1D8", "weak": "FDE7E5"}
TREND_INK = {"good": "1B7F3B", "bad": "B42318", "flat": "5F6B7A", "none": "5F6B7A"}


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
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 8})
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
        ax.bar(xs, ys, width=0.62, color="#" + SERIES, zorder=2)
        for xi, yi in zip(xs, ys):
            ax.text(xi, yi, C.fmt_num(vals[ps[xs.index(xi)]]["value"], unit, lang), ha="center", va="bottom", fontsize=6.5, color="#" + INK)
        ax.set_xticks(xs, ps, fontsize=6.5, color="#" + MUTED)
        ul = C.unit_label(unit, lang)
        ax.set_title(C.tr(M.get(m, m), lang) + (f", {ul}" if ul else ""), fontsize=7.5, color="#" + INK, loc="left")
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

    blocks = [("title", iss["name"]), ("subtitle", f"{t('title')} · {rep['period']}"),
              ("meta", [(t("analysis_date"), [("t", str(rep["analysis_date"]), False)]),
                        (t("rating"), rich(rep.get("rating", ""))), (t("next_report"), rich(rep.get("next_report", "")))])]
    hl = highlights(rep, lang)
    if hl:
        blocks.append(("tiles", hl))

    blocks += [("h2", t("quick_summary")), ("bullets", [rich(x, True) for x in rep.get("quick_summary", [])])]

    # priemonės
    rows = []
    for s in secs:
        mat = s["maturity_date"] or (t("perpetual") if s["type"] == "bond" else t("share"))
        cpn = C.fmt_num(s["coupon_pct"], "%", lang) if s["coupon_pct"] else "–"
        note = ("⚠ " if s.get("verified") == "review" else "") + s.get("status_note", "")
        rows.append([s["isin"], s["name"], cpn, mat, note])
    blocks += [("h2", t("instruments")), ("table", [t("isin"), t("instrument"), t("coupon"), t("maturity"), t("note")], rows, [3.0, 4.6, 1.7, 2.3, 5.4], {"num": [2]})]

    # rodikliai: šaltiniai - išnašos prie rodiklio pavadinimo
    periods = rep.get("periods", [])
    rows = []
    for m in rep.get("ratios", []):
        vals = hist.get(m, {})
        if not vals:
            continue
        unit = next(iter(vals.values()))["unit"]
        ul = C.unit_label(unit, lang)
        label = [("t", C.tr(M.get(m, m), lang) + (f", {ul}" if ul else ""), False)]
        seen = []
        for p in periods:
            r = vals.get(p)
            if r and r["source"] in cites.sources and (r["source"], r.get("page", "")) not in seen:
                seen.append((r["source"], r.get("page", "")))
        for src, pg in seen[:3]:
            label.append(("c", cites.add(src, pg), cites.url(src, pg)))
        cells = [label]
        last, prev, ch, trend, level = C.ratio_change(m, vals, periods, lang)
        for p in periods:
            r = vals.get(p)
            v = "–" if not r else C.fmt_num(r["value"], unit, lang) + ("*" if r.get("verified") == "check" else "ᶜ" if r.get("verified") == "calc" else "")
            if p == last and level:
                v = ("styled", v, LEVEL_FILL[level], None, True)
            cells.append(v)
        cells.append(("styled", f"{ch} vs {prev}" if ch else "–", None, TREND_INK[trend], trend in ("good", "bad")))
        rows.append(cells)
    blocks += [("h2", t("key_ratios")), ("table", [t("metric")] + periods + [t("change")], rows, [5.2] + [2.1] * len(periods) + [3.9],
                                         {"num": list(range(1, len(periods) + 2))}),
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

    # ESG - tiksliai pagal šabloną (klausimai anglų k., nekeičiami), be lentelės
    tpl = C.seb()["esg_template"]
    esg = rep.get("esg", {})
    items, comments = esg.get("items", {}), esg.get("comments", {})
    blocks.append(("h2", tpl["title"]))
    for sec in tpl["sections"]:
        blocks.append(("h3", sec["heading"]))
        for q in sec["items"]:
            it = items.get(q["id"], {"status": "review"})
            st = it.get("status", "review")
            blocks.append(("check", "☑" if st == "yes" else "☐", q["q"]))
            pts = [(lab, [rich(x) for x in items]) for lab, items in C.points(C.tr(it.get("evidence", ""), lang))]
            if st == "review":
                pts.insert(0, (t("needs_review").strip().rstrip("."), []))
            if pts:
                blocks.append(("evlist", pts))
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
def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def borders(table, color=LINE, inside=True):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        none = edge in ("left", "right", "insideV") or (not inside and edge == "insideH")
        e.set(qn("w:val"), "nil" if none else "single")
        if not none:
            e.set(qn("w:sz"), "4")
            e.set(qn("w:color"), color)
        b.append(e)
    tblPr.append(b)


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


def to_docx(blocks, path, footer_text):
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin, sec.bottom_margin = Cm(1.6), Cm(1.6)
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(10.5)
    st.font.color.rgb = RGBColor.from_string(INK)
    st.paragraph_format.space_after = Pt(4)
    fp = sec.footer.paragraphs[0]
    fr = fp.add_run(footer_text + "   ·   ")
    fr.font.size = Pt(8)
    fr.font.color.rgb = RGBColor.from_string(MUTED)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    fp._p.append(fld)
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    for kind, *rest in blocks:
        if kind == "title":
            p = doc.add_paragraph()
            r = p.add_run(rest[0])
            r.bold, r.font.size, r.font.color.rgb = True, Pt(24), RGBColor.from_string(ACCENT)
            p.paragraph_format.space_after = Pt(0)
        elif kind == "subtitle":
            p = doc.add_paragraph()
            r = p.add_run(rest[0])
            r.font.size, r.font.color.rgb = Pt(12), RGBColor.from_string(MUTED)
        elif kind == "meta":
            for k, segs in rest[0]:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(f"{k}: ")
                r.bold = True
                r.font.size = Pt(9.5)
                add_segs(p, segs, size=9.5)
        elif kind == "tiles":
            tiles = rest[0]
            doc.add_paragraph().paragraph_format.space_after = Pt(0)
            tb = doc.add_table(rows=1, cols=len(tiles))
            tb.alignment = WD_TABLE_ALIGNMENT.CENTER
            for c, (label, val, delta, per) in zip(tb.rows[0].cells, tiles):
                shade(c, ACCENT_LIGHT)
                c.text = ""
                p = c.paragraphs[0]
                r = p.add_run(label + f" ({per})")
                r.font.size, r.font.color.rgb = Pt(8), RGBColor.from_string(MUTED)
                p = c.add_paragraph()
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(val)
                r.bold, r.font.size, r.font.color.rgb = True, Pt(13), RGBColor.from_string(ACCENT)
                p = c.add_paragraph()
                r = p.add_run(delta)
                r.font.size, r.font.color.rgb = Pt(8), RGBColor.from_string(MUTED)
            borders(tb, color="FFFFFF")
        elif kind == "image":
            doc.add_picture(str(rest[0]), width=Cm(17))
        elif kind == "h2":
            p = doc.add_paragraph()
            p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(14), Pt(4)
            r = p.add_run(rest[0])
            r.bold, r.font.size, r.font.color.rgb = True, Pt(14), RGBColor.from_string(ACCENT)
            pPr = p._p.get_or_add_pPr()
            bd = OxmlElement("w:pBdr")
            b = OxmlElement("w:bottom")
            for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:color", LINE), ("w:space", "2")):
                b.set(qn(k), v)
            bd.append(b)
            pPr.append(bd)
        elif kind == "h3":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            r = p.add_run(rest[0])
            r.bold, r.font.size = True, Pt(11)
        elif kind == "small":
            p = doc.add_paragraph()
            r = p.add_run(rest[0])
            r.font.size, r.italic, r.font.color.rgb = Pt(8), True, RGBColor.from_string(MUTED)
        elif kind == "bullets":
            for segs in rest[0]:
                p = doc.add_paragraph(style="List Bullet")
                p.paragraph_format.space_after = Pt(3)
                add_segs(p, segs)
        elif kind == "check":
            p = doc.add_paragraph()
            p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(4), Pt(1)
            r = p.add_run(rest[0] + " ")
            r.font.size = Pt(12)
            p.add_run(rest[1])
        elif kind == "evlist":
            for lab, items in rest[0]:
                if lab:
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent, p.paragraph_format.space_after = Cm(0.7), Pt(0)
                    r = p.add_run(lab)
                    r.bold, r.font.size = True, Pt(9.5)
                    r.font.color.rgb = RGBColor.from_string("B42318" if not items else "333333")
                for segs in items:
                    p = doc.add_paragraph()
                    ind = Cm(1.3 if lab else 0.9)
                    p.paragraph_format.left_indent, p.paragraph_format.first_line_indent = ind, Cm(-0.35)
                    p.paragraph_format.space_after = Pt(0)
                    p.add_run("•  ").font.size = Pt(9.5)
                    add_segs(p, segs, size=9.5, color="333333")
            doc.paragraphs[-1].paragraph_format.space_after = Pt(6)
        elif kind == "evidence":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.7)
            add_segs(p, rest[0], size=9.5, color="333333")
        elif kind == "note":
            p = doc.add_paragraph()
            r = p.add_run(rest[0])
            r.italic, r.font.size = True, Pt(9.5)
        elif kind == "comment":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            r = p.add_run(rest[0] + " ")
            r.bold = True
            add_segs(p, rest[1])
        elif kind == "notes":
            for n, title, pg, note, url, key, date in rest[0]:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.left_indent, p.paragraph_format.first_line_indent = Cm(0.6), Cm(-0.6)
                r = p.add_run(f"{n}. ")
                r.font.size = Pt(8.5)
                desc = title + (f", p. {pg}" if pg else "") + (f" ({note})" if note else "") + (f", {date}" if date else "")
                if url:
                    hyperlink(p, url, desc, size=8.5)
                else:
                    p.add_run(desc).font.size = Pt(8.5)
        elif kind == "table":
            header, rows, widths, opts = rest
            num = set(opts.get("num", []))
            tbl = doc.add_table(rows=1, cols=len(header))
            borders(tbl)
            for i, h in enumerate(header):
                c = tbl.rows[0].cells[i]
                c.text = ""
                r = c.paragraphs[0].add_run(str(h))
                r.bold, r.font.size, r.font.color.rgb = True, Pt(8.5), RGBColor(255, 255, 255)
                shade(c, ACCENT)
                if i in num:
                    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for ri, row in enumerate(rows):
                cells = tbl.add_row().cells
                for i, v in enumerate(row):
                    cells[i].text = ""
                    par = cells[i].paragraphs[0]
                    par.paragraph_format.space_after = Pt(0)
                    if isinstance(v, tuple) and v[0] == "status":
                        shade(cells[i], STATUS_FILL.get(v[1], "FFFFFF"))
                        r = par.add_run(v[2])
                        r.font.size, r.bold = Pt(8.5), True
                        continue
                    if isinstance(v, tuple) and v[0] == "styled":
                        if v[2]:
                            shade(cells[i], v[2])
                        r = par.add_run(v[1])
                        r.font.size, r.bold = Pt(8.5), v[4]
                        if v[3]:
                            r.font.color.rgb = RGBColor.from_string(v[3])
                        if i in num:
                            par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        continue
                    add_segs(par, cell_segs(v), size=8.5)
                    if i in num:
                        par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                if ri % 2 == 1:
                    for c in cells:
                        if c._tc.tcPr is None or c._tc.tcPr.find(qn("w:shd")) is None:
                            shade(c, "F6F8FA")
            for row in tbl.rows:
                for i, wcm in enumerate(widths[:len(row.cells)]):
                    row.cells[i].width = Cm(wcm)
    doc.save(path)


# ---------- PDF ----------
def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


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


def to_pdf(blocks, path, footer_text):
    S = lambda name, **kw: ParagraphStyle(name, fontName=kw.pop("font", "DejaVu"), textColor=kw.pop("color", colors.HexColor("#" + INK)), **kw)  # noqa: E731
    st = {"title": S("t", font="DejaVu-Bold", fontSize=22, leading=26, color=colors.HexColor("#" + ACCENT)),
          "subtitle": S("st", fontSize=11, leading=15, color=colors.HexColor("#" + MUTED), spaceAfter=8),
          "h2": S("h", font="DejaVu-Bold", fontSize=12.5, leading=16, color=colors.HexColor("#" + ACCENT), spaceBefore=14, spaceAfter=6),
          "h3": S("h3", font="DejaVu-Bold", fontSize=10, leading=13, spaceBefore=8, spaceAfter=3),
          "p": S("p", fontSize=9, leading=13),
          "meta": S("m", fontSize=8.5, leading=12),
          "b": S("b", fontSize=9, leading=13, leftIndent=11, bulletIndent=2, spaceAfter=3),
          "check": S("c", fontSize=9, leading=13, spaceBefore=5, spaceAfter=1),
          "ev": S("e", fontSize=8.5, leading=12, leftIndent=16, color=colors.HexColor("#333333"), spaceAfter=2),
          "evh": S("eh", fontSize=8.5, leading=12, leftIndent=16, spaceBefore=1),
          "evb": S("eb", fontSize=8.5, leading=11.5, leftIndent=28, bulletIndent=18, color=colors.HexColor("#333333")),
          "evb2": S("eb2", fontSize=8.5, leading=11.5, leftIndent=40, bulletIndent=30, color=colors.HexColor("#333333")),
          "note": S("n", fontSize=8.5, leading=12, spaceBefore=4),
          "small": S("s", fontSize=7, leading=9, color=colors.HexColor("#" + MUTED), spaceBefore=3),
          "src": S("src", fontSize=7.5, leading=10, leftIndent=12, firstLineIndent=-12, spaceAfter=1),
          "cell": S("cell", fontSize=7.8, leading=10), "celln": S("celln", fontSize=7.8, leading=10, alignment=2),
          "cellh": S("cellh", font="DejaVu-Bold", fontSize=7.8, leading=10, color=colors.white),
          "cellhn": S("cellhn", font="DejaVu-Bold", fontSize=7.8, leading=10, color=colors.white, alignment=2),
          "tl": S("tl", fontSize=7, leading=9, color=colors.HexColor("#" + MUTED)),
          "tv": S("tv", font="DejaVu-Bold", fontSize=12.5, leading=16, color=colors.HexColor("#" + ACCENT)),
          "td": S("td", fontSize=7, leading=9, color=colors.HexColor("#" + MUTED))}
    W = 17.0 * cm
    flow = []
    for kind, *rest in blocks:
        if kind in ("title", "subtitle", "h2", "h3"):
            flow.append(Paragraph(esc(rest[0]), st[kind]))
        elif kind == "meta":
            for k, segs in rest[0]:
                flow.append(Paragraph(f"<b>{esc(k)}:</b> {segs_html(segs)}", st["meta"]))
        elif kind == "tiles":
            tiles = rest[0]
            cells = [[Paragraph(esc(f"{lb} ({per})"), st["tl"]), Paragraph(esc(v), st["tv"]), Paragraph(esc(d), st["td"])] for lb, v, d, per in tiles]
            tb = Table([cells], colWidths=[W / len(tiles)] * len(tiles))
            tb.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#" + ACCENT_LIGHT)), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                    ("LINEAFTER", (0, 0), (-2, -1), 3, colors.white), ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
            flow += [Spacer(1, 8), tb]
        elif kind == "image":
            img = Image(str(rest[0]))
            ratio = img.imageHeight / img.imageWidth
            img.drawWidth, img.drawHeight = W, W * ratio
            flow += [Spacer(1, 6), img]
        elif kind == "small":
            flow.append(Paragraph(esc(rest[0]), st["small"]))
        elif kind == "bullets":
            for segs in rest[0]:
                flow.append(Paragraph(segs_html(segs), st["b"], bulletText="•"))
        elif kind == "check":
            flow.append(Paragraph(f'<font size="11">{rest[0]}</font>  {esc(rest[1])}', st["check"]))
        elif kind == "evlist":
            for lab, items in rest[0]:
                if lab:
                    flow.append(Paragraph(f'<b><font color="#{"B42318" if not items else "333333"}">{esc(lab)}</font></b>', st["evh"]))
                for segs in items:
                    flow.append(Paragraph(segs_html(segs), st["evb2" if lab else "evb"], bulletText="•"))
            flow.append(Spacer(1, 4))
        elif kind == "evidence":
            flow.append(Paragraph(segs_html(rest[0]), st["ev"]))
        elif kind == "note":
            flow.append(Paragraph(f"<i>{esc(rest[0])}</i>", st["note"]))
        elif kind == "comment":
            flow.append(Paragraph(f"<b>{esc(rest[0])}</b> {segs_html(rest[1])}", st["note"]))
        elif kind == "notes":
            for n, title, pg, note, url, key, date in rest[0]:
                desc = esc(title + (f", p. {pg}" if pg else "") + (f" ({note})" if note else "") + (f", {date}" if date else ""))
                link = f'<a href="{esc(url)}" color="#{SERIES}">{desc}</a>' if url else desc
                flow.append(Paragraph(f"{n}. {link}", st["src"]))
        elif kind == "table":
            header, rows, widths, opts = rest
            num = set(opts.get("num", []))
            widths = [w * W / sum(widths) for w in widths]
            data = [[Paragraph(esc(h), st["cellhn" if i in num else "cellh"]) for i, h in enumerate(header)]]
            style = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + ACCENT)), ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor("#" + LINE)),
                     ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4)]
            for ri, row in enumerate(rows, start=1):
                line = []
                if ri % 2 == 0:  # zebra pirmiau, kad langelių spalvos būtų viršuje
                    style.append(("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#F6F8FA")))
                for i, v in enumerate(row):
                    if isinstance(v, tuple) and v[0] == "status":
                        style.append(("BACKGROUND", (i, ri), (i, ri), colors.HexColor("#" + STATUS_FILL.get(v[1], "FFFFFF"))))
                        line.append(Paragraph(f"<b>{esc(v[2])}</b>", st["cell"]))
                    elif isinstance(v, tuple) and v[0] == "styled":
                        if v[2]:
                            style.append(("BACKGROUND", (i, ri), (i, ri), colors.HexColor("#" + v[2])))
                        txt = esc(v[1])
                        txt = f'<font color="#{v[3]}">{txt}</font>' if v[3] else txt
                        line.append(Paragraph(f"<b>{txt}</b>" if v[4] else txt, st["celln" if i in num else "cell"]))
                    else:
                        line.append(Paragraph(segs_html(cell_segs(v)), st["celln" if i in num else "cell"]))
                data.append(line)
            tb = Table(data, colWidths=widths, repeatRows=1)
            tb.setStyle(TableStyle(style))
            flow += [tb, Spacer(1, 4)]

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("DejaVu", 7)
        canvas.setFillColor(colors.HexColor("#" + MUTED))
        canvas.drawRightString(A4[0] - 2 * cm, 1.0 * cm, f"{footer_text}   ·   {doc.page}")
        canvas.restoreState()
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=1.6 * cm, bottomMargin=1.6 * cm,
                            title=footer_text, author="issuers-review")
    doc.build(flow, onFirstPage=footer, onLaterPages=footer)


# ---------- Excel ----------
def to_xlsx(rep, lang, path):
    L, M = C.i18n(), C.metrics()
    t = lambda k: C.tr(L.get(k, k), lang)  # noqa: E731
    wb = Workbook()
    head, fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=ACCENT)

    def sheet(ws, header, rows, widths):
        ws.append(header)
        for c in ws[1]:
            c.font, c.fill = head, fill
        for r in rows:
            ws.append(r)
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "A2"

    hist = C.history(rep["issuer"])
    periods = sorted({p for v in hist.values() for p in v}, key=period_key)
    ws = wb.active
    ws.title = t("key_ratios")[:31]
    rows = []
    for m, vals in hist.items():
        unit = next(iter(vals.values()))["unit"]
        row = [C.tr(M.get(m, m), lang), unit]
        for p in periods:
            r = vals.get(p)
            row.append(float(r["value"]) if r and r["value"] else None)
        srcs = sorted({f"{r['source']}" + (f" p.{r['page']}" if r.get("page") else "") for r in vals.values()})
        rows.append(row + ["; ".join(srcs)])
    sheet(ws, [t("metric"), "Unit"] + periods + [t("sources")], rows, [34, 8] + [11] * len(periods) + [30])

    ws = wb.create_sheet(t("instruments")[:31])
    secs = [s for s in C.securities() if s["issuer"] == rep["issuer"]]
    sheet(ws, ["ISIN", t("instrument"), "Held", t("coupon"), "Issue", t("maturity"), "Call", t("status"), t("note"), "Source"],
          [[s["isin"], s["name"], s["held"], float(s["coupon_pct"]) if s["coupon_pct"] else None, s["issue_date"], s["maturity_date"],
            s["call_date"], s["status"], s["status_note"], s["source_url"]] for s in secs], [15, 34, 6, 8, 11, 11, 11, 9, 50, 40])

    ws = wb.create_sheet(t("plans")[:31])
    sheet(ws, [t("plan"), t("outcome"), t("status"), t("sources")],
          [[C.plain(C.tr(p["plan"], lang)), C.plain(C.tr(p["outcome"], lang)), t(STATUS_KEYS.get(p["status"], p["status"])), p.get("src", "")]
           for p in rep.get("plans", [])], [55, 55, 14, 16])

    ws = wb.create_sheet("ESG")
    items = rep.get("esg", {}).get("items", {})
    rows = []
    for sec in C.seb()["esg_template"]["sections"]:
        for q in sec["items"]:
            it = items.get(q["id"], {})
            rows.append([sec["heading"], q["q"], it.get("status", "review"), C.tr(it.get("evidence"), lang)])
    sheet(ws, ["Section", "Question", "Status", "Evidence"], rows, [30, 60, 10, 90])

    ws = wb.create_sheet(t("sources")[:31])
    sheet(ws, ["#", "Title", "Date", "URL"], [[k, v["title"], str(v.get("date", "")), v["url"]] for k, v in rep.get("sources", {}).items()], [6, 60, 12, 90])
    wb.save(path)


# ---------- paleidimas ----------
def inputs_hash(issuer, rep):
    """Failai perkuriami tik pasikeitus jų šaltiniams (kitaip kiekvienas paleidimas keistų dvejetainius failus)."""
    h = hashlib.sha1()
    for p in [rep["_file"], C.DATA / "history" / f"{issuer}.csv", Path(__file__), Path(C.__file__), *sorted(C.CONFIG.glob("*"))]:
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
                    if "xlsx" in cfg["formats"]:
                        to_xlsx(rep, lang, base.with_suffix(".xlsx"))
                    built.append(base.name)
    C.save_json(state_path, state)
    print("Sugeneruota:", ", ".join(built) or "nieko (šaltiniai nepasikeitė)")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    build(args[0] if args else None, force="--force" in sys.argv)
