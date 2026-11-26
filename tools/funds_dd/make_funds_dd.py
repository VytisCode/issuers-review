"""SEB „Liquid Funds DD Template“ atkūrimas (13 skaidrių, .pptx) pagal Copilot specifikaciją (2026-10-28).

  python make_funds_dd.py                      # tuščias šablonas -> Liquid_Funds_DD_Template.pptx
  python make_funds_dd.py DUOMENYS.json        # užpildytas DD (vėliau; struktūra - FILL žemiau)

Šriftai: SEB SansSerif nėra -> temos šriftai Arial (antraštės ir tekstas). Visi tekstai naudoja TEMOS šriftus
(+mj-lt / +mn-lt), todėl pritaikius originalų SEB šabloną (Design -> Browse themes) šriftai pasikeis patys.
Logotipas: SEB_LOGO failas (PNG/EMF su skaidriu fonu) dedamas į Slide Master; jei failo nėra - vieta tuščia
(specifikacija draudžia rašyti „SEB“ tekstu).
"""
import copy
import json
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

HERE = Path(__file__).resolve().parent
SEB_LOGO = HERE / "seb_logo.png"
DARK = RGBColor(0x19, 0x18, 0x24)
GRID = "AFAFAF"
FONT = "Arial"
THEME = {"dk1": "191824", "lt1": "FFFFFF", "dk2": "AFAFAF", "lt2": "FFFFFF", "accent1": "87C081", "accent2": "DCB241",
         "accent3": "4561A7", "accent4": "D86D64", "accent5": "AFAFAF", "accent6": "3E3E3E", "hlink": "191824", "folHlink": "191824"}
TITLE_BOX = (0.602, 0.846, 12.129, 1.150)
COL = (2.016, 9.978)

# (antraštė, lentelė [(1 stulpelis, 2 stulpelis)], vieta (x, y, w, h), šrifto dydis, paryškintos eilutės {idx: (bold1, bold2)}, eilučių aukščiai)
EX = "In case any red flags, risks, unclarities were identified then additional checks (sanctions screening and adverse media checks) on the fund’s Ultimate Beneficial Owners and/or senior management shall be performed."
UNGC = "Red flag/signal related to controversies based on UN Global Compact, OECD Guidelines for Multinational Enterprises or any other relevant principles/standards"
TABLES = {
    3: ("Fund overview  (index funds)", [("ISIN", ""), ("Ticker", ""), ("Fees", ""), ("AUM", ""), ("Bid-ask spread", ""),
        ("Liquidity (e.g. 90D Avg Volume)", ""), ("Replication Method", ""), ("Tracking difference / tracking error", ""),
        ("Securities lending / CP risk management", ""), ("Other relevant information (taxes, leverage, etc)", "")],
        (0.602, 1.396, 11.993, 5.099), 12, {}, None),
    4: ("Fund overview (mutual funds)", [("ISIN", ""), ("Ticker", ""), ("Fees", ""), ("Retrocessions", ""), ("Swing pricing", ""), ("AUM", ""),
        ("Trading cycle", ""), ("Valuation cycle", ""), ("Notice period", ""), ("Currency", ""), ("", "")],
        (0.602, 1.396, 11.993, 5.604), 12, {}, None),
    5: ("Index overview (if index is different than used internally)", [("Index", ""), ("Index weighting method", ""),
        ("Country / sector / company / factor actual exposures", "Reference to the fact sheet"), ("Actual credit ratings", "Reference to the fact sheet"),
        ("Interest / spread duration", ""), ("FX exposure (principle)", "")],
        (0.602, 1.609, 11.993, 3.076), 12, {}, None),
    6: ("ESG overview (excluding climate future index fund)", [("SFDR article", "Article 6 / 8 / 9"),
        ("Exclusionary screening", "The fund excludes companies from the investment universe involved in the industries/reasons below"),
        ("Yes/No", "Controversial and nuclear weapons"), ("Yes/No", "Tobacco"), ("Yes/No", "Gambling"),
        ("Yes/No", "Fossil fuel related screens (not fully excluded but sectors like thermal coal, oil sands, etc)"), ("Yes/No", UNGC),
        ("Other important exclusions:", ""), ("Positive screening", ""), ("Yes/No", "The fund manager is UN-PRI signatory"),
        ("Yes/No", "The fund implements positive screening"), ("Additional comments:", "")],
        (0.602, 1.396, 11.993, 6.110), 12, {1: (True, True), 8: (True, False)}, None),
    7: ("ESG overview (excluding climate future index fund)", [("Engagement", "(voluntary field for passively managed funds)"),
        ("Yes/No", "The fund manager/company engages with the companies"), ("Additional comments", ""), ("Indicators", ""),
        ("Yes/No", "The indicators support the targets of sustainability and PAI indicators followed by the Fund"), ("Additional comments:", ""),
        ("Sanctions", ""), ("Yes/No", "Compliance with the list of sanctions (fund management company) (add a screenshot)"),
        ("Yes/No", "No issued identified from adverse media checks (add a screenshot)"), ("No issues identified/checks performed", EX)],
        (0.602, 1.396, 11.993, 5.076), 12, {0: (True, False), 3: (True, False), 6: (True, False)}, None),
    8: ("ESG overview (climate future index fund)", [("Yes / No", "Article 8 and 9 funds are only allowed"),
        ("Yes/No", "Majority of the assets are invested into ETFs following EU Climate Transition Benchmark and EU Paris-aligned Benchmark or other Article 9 funds fulfilling objective of the fund and contributing to lower emissions"),
        ("Exclusionary screening", "The fund shall have no involvement in the following industries below"),
        ("Yes/No", "Controversial and nuclear weapons"), ("Yes/No", "Thermal coal and oil sands"), ("Yes/No", "Tobacco"),
        ("Exclusionary screening", "The fund minimizes exposure to the following industries below"),
        ("Yes/No", "Fossil fuel, gambling, civilian arms and other controversial sectors"), ("Yes/No", UNGC), ("Indicators", ""),
        ("Yes/No", "The indicators support the targets of sustainability and PAI indicators followed by the Fund (Going forward decarbonization trend should be downward and the fund should have also lower carbon intensity compared to broad investable equity market (3 year rolling average))"),
        ("Sanctions", ""), ("Yes/No", "Compliance with the list of sanctions (fund management company) (add a screenshot)"),
        ("Yes/No", "No issued identified from adverse media checks (add a screenshot)"), ("No issues identified/checks performed", EX)],
        (0.535, 0.964, 11.993, 6.510), 11, {2: (True, True), 6: (True, True), 9: (True, False), 11: (True, False)},
        [0.372, 0.391, 0.391, 0.374, 0.374, 0.374, 0.374, 0.400, 0.478, 0.417, 0.409, 0.506, 0.506, 0.506, 0.506]),
}
EMPTY = {2: "Conclusion", 9: "Performance", 10: "Investment team", 11: "Investment strategy & process"}


def set_theme(prs):
    """Temos spalvos ir šriftai (Slide Master tema)."""
    master = prs.slide_master
    rel = next(r for r in master.part.rels.values() if r.reltype.endswith("/theme"))
    part = rel.target_part
    root = etree.fromstring(part.blob)
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    cs = root.find(f".//{{{a}}}clrScheme")
    cs.set("name", "SEB")
    for k, v in THEME.items():
        el = cs.find(f"{{{a}}}{k}")
        for c in list(el):
            el.remove(c)
        etree.SubElement(el, f"{{{a}}}srgbClr", val=v)
    fs = root.find(f".//{{{a}}}fontScheme")
    fs.set("name", "SEB fallback")
    for tag in ("majorFont", "minorFont"):
        fs.find(f"{{{a}}}{tag}/{{{a}}}latin").set("typeface", FONT)
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def run_fmt(r, size, bold=False, major=False):
    f = r.font
    f.size, f.bold, f.name = Pt(size), bold, "+mj-lt" if major else "+mn-lt"
    f.color.rgb = DARK


def no_space(p):
    p.space_before = p.space_after = Pt(0)
    p.alignment = PP_ALIGN.LEFT


def textbox(slide, box, paras, inset=0.0):
    """paras: [(tekstas, dydis, bold, major)]"""
    tb = slide.shapes.add_textbox(*(Inches(v) for v in box))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(inset)
    tf.margin_top = tf.margin_bottom = Inches(0.05)
    for i, (txt, size, bold, major) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        no_space(p)
        r = p.add_run()
        r.text = txt
        run_fmt(r, size, bold, major)
    return tb


def title(slide, txt, box=TITLE_BOX, sub=None):
    t = slide.shapes.title
    t.left, t.top, t.width, t.height = (Inches(v) for v in box)
    tf = t.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]
    no_space(p)
    r = p.add_run()
    r.text = txt
    run_fmt(r, 35, major=True)
    if sub:
        p2 = tf.add_paragraph()
        no_space(p2)
        r2 = p2.add_run()
        r2.text = sub
        run_fmt(r2, 16)


def cell_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(tag), w="12700", cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=GRID)
        etree.SubElement(ln, qn("a:prstDash"), val="solid")
    # baltas fonas (be užpildo stiliaus)
    for f in tcPr.findall(qn("a:noFill")) + tcPr.findall(qn("a:solidFill")):
        tcPr.remove(f)
    sf = etree.SubElement(tcPr, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr"), val="FFFFFF")


def table(slide, rows, box, size, bold, heights):
    x, y, w, h = box
    gs = slide.shapes.add_table(len(rows), 2, *(Inches(v) for v in box))
    tbl = gs.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
    sid.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # No Style, Table Grid
    tbl.columns[0].width, tbl.columns[1].width = Inches(COL[0]), Inches(COL[1])
    hs = heights or [h / len(rows)] * len(rows)
    for i, (c1, c2) in enumerate(rows):
        tbl.rows[i].height = Inches(hs[i])
        b = bold.get(i, (False, False))
        for j, txt in enumerate((c1, c2)):
            cell = tbl.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell_border(cell)
            lines = txt if isinstance(txt, list) else [txt]
            tf = cell.text_frame
            tf.word_wrap = True
            for k, line in enumerate(lines):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                no_space(p)
                r = p.add_run()
                r.text = line
                run_fmt(r, size, b[j])
            if not txt:  # tuščio langelio dydis
                tf.paragraphs[0].font.size = Pt(size)
    return tbl


def master_setup(prs):
    """Skaidrės numeris ir logotipas Slide Master'yje; antraštės vieta - maketuose."""
    m = prs.slide_master
    for ph in m.placeholders:
        t = ph.placeholder_format.type
        name = str(t)
        if "SLIDE_NUMBER" in name:
            ph.left, ph.top, ph.width, ph.height = Inches(0.602), Inches(6.936), Inches(0.713), Inches(0.399)
            for p in ph.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                for r in p.runs:
                    run_fmt(r, 8)
                p.font.size = Pt(8)
        elif "DATE" in name or "FOOTER" in name:
            ph._element.getparent().remove(ph._element)
        elif "TITLE" in name:
            ph.left, ph.top, ph.width, ph.height = (Inches(v) for v in TITLE_BOX)
    if SEB_LOGO.exists():
        w, hh = 0.441, 0.194
        top = 6.936 + (0.399 - hh) / 2
        m.shapes.add_picture(str(SEB_LOGO), Inches(13.333 - 0.45 - w), Inches(top), Inches(w), Inches(hh))


def slide_number(slide, layout):
    """Kopijuoja maketo skaidrės numerio vietą į skaidrę (taip PowerPoint rodo numerį)."""
    src = next((ph for ph in slide.slide_layout.slide_master.placeholders if "SLIDE_NUMBER" in str(ph.placeholder_format.type)), None)
    if src is None:
        return
    el = copy.deepcopy(src._element)
    slide.shapes._spTree.append(el)


def build(data=None, out=HERE / "Liquid_Funds_DD_Template.pptx"):
    data = data or {}
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    set_theme(prs)
    master_setup(prs)
    lay = prs.slide_layouts[5]  # Title Only
    for ph in lay.placeholders:
        n = str(ph.placeholder_format.type)
        if "TITLE" in n:
            ph.left, ph.top, ph.width, ph.height = (Inches(v) for v in TITLE_BOX)
        elif "DATE" in n or "FOOTER" in n:
            ph._element.getparent().remove(ph._element)
        elif "SLIDE_NUMBER" in n:
            ph.left, ph.top, ph.width, ph.height = Inches(0.602), Inches(6.936), Inches(0.713), Inches(0.399)
    blank = prs.slide_layouts[6]
    for ph in list(blank.placeholders):
        n = str(ph.placeholder_format.type)
        if "DATE" in n or "FOOTER" in n:
            ph._element.getparent().remove(ph._element)
        elif "SLIDE_NUMBER" in n:
            ph.left, ph.top, ph.width, ph.height = Inches(0.602), Inches(6.936), Inches(0.713), Inches(0.399)
    f = data.get("front", {})
    for n in range(1, 14):
        if n == 1:
            s = prs.slides.add_slide(blank)
            textbox(s, (0.602, 3.330, 12.129, 2.008), [
                (f.get("name", "Fund name"), 35, False, True),
                ("Fund manager:" + (" " + f["manager"] if f.get("manager") else ""), 20, False, False),
                ("Fund manager contact person:" + (" " + f["contact"] if f.get("contact") else ""), 20, False, False),
                ("Due diligence report prepared by:" + (" " + f["prepared"] if f.get("prepared") else ""), 20, False, False),
                ("Date:" + (" " + f["date"] if f.get("date") else ""), 20, False, False)])
        else:
            s = prs.slides.add_slide(lay)
            if n in TABLES:
                ttl, rows, box, size, bold, hs = TABLES[n]
                title(s, ttl, (0.535, 0.30, 12.129, 0.664) if n == 8 else TITLE_BOX)
                table(s, rows, box, size, bold, hs)
            elif n == 12:
                title(s, "Current portfolio characteristics", sub="including impact on current portfolio")
            elif n == 13:
                title(s, "Other")
                textbox(s, (0.602, 1.652, 12.129, 2.008), [
                    ("Legal analysis: usually standard investments do not need legal analysis but the need is decided by the fund manager", 20, False, False),
                    ("", 20, False, False), ("", 20, False, False),
                    ("Operational risk analysis: usually standard investments do not need operational risk analysis but the need is decided by the fund manager", 20, False, False)])
            else:
                title(s, EMPTY[n])
        slide_number(s, lay)
    prs.save(out)
    return out


if __name__ == "__main__":
    d = json.loads(Path(sys.argv[1]).read_text()) if len(sys.argv) > 1 else None
    print(build(d))
