"""Funds DD (SEB „Liquid Funds DD Template“) užpildymo įrankiai: skaidrės, lentelės, žymės, tinkamumo patikra.

Teksto žymės langeliuose ir pastraipose:
  **tekstas**  – paryškinta
  [S3] / [S3 p.12] – šaltinio žymė, nuoroda į šaltinio URL (SOURCES)
  !!tekstas!!  – geltonas paryškinimas (NOT FOUND / patikrinti viduje)
  ^^tekstas^^  – raudonas tekstas (problema, rizika)
  ++tekstas++  – žalias tekstas (pasirinkta reikšmė apskaitos lentelėje)
  ~~tekstas~~  – pilkas tekstas (nepasirinkta reikšmė)
Langelis gali būti eilutė arba eilučių sąrašas (kiekviena – atskira pastraipa).
"""
import json
import re
from pathlib import Path

from lxml import etree
from PIL import ImageFont
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

import make_funds_dd as M

RED, GREEN, GREY = RGBColor(0xC0, 0x1E, 0x1E), RGBColor(0x3E, 0x8E, 0x41), RGBColor(0x9A, 0x9A, 0x9A)
RISK = RGBColor(0xB4, 0x5F, 0x06)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TOKEN = re.compile(r"(\*\*.+?\*\*|!!.+?!!|\^\^.+?\^\^|\+\+.+?\+\+|~~.+?~~|\[[A-Z]\d+[^\]]*\])")
FONT_R = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
FONT_B = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
SOURCES = {}       # "S1": (pavadinimas, url)
MISSING = []       # (skaidrė, kas trūksta)
FINAL = False      # galutinė versija (savininkas 2026-11-23): !!…!! rodoma pilkai kaip n/a, be raudonos ir be MISSING
PAGE_H, BOTTOM = 7.5, 6.85   # lentelė neturi užlįsti ant numerio / logotipo eilutės


def highlight(r, color="FFFF00"):
    rPr = r._r.get_or_add_rPr()
    h = etree.SubElement(rPr, qn("a:highlight"))
    etree.SubElement(h, qn("a:srgbClr"), val=color)


def add_runs(p, text, size, bold=False, slide_no=None, color_override=None):
    for part in TOKEN.split(text or ""):
        if not part:
            continue
        r = p.add_run()
        kind = None
        if part.startswith("**"):
            txt, kind = part[2:-2], "b"
        elif part.startswith("!!"):
            txt, kind = part[2:-2], "hl"
        elif part.startswith("^^"):
            txt, kind = part[2:-2], "red"
        elif part.startswith("++"):
            txt, kind = part[2:-2], "green"
        elif part.startswith("~~"):
            txt, kind = part[2:-2], "grey"
        elif re.fullmatch(r"\[[A-Z]\d+[^\]]*\]", part):
            txt, kind = part, "src"
        else:
            txt = part
        r.text = txt
        if kind == "hl" and FINAL:
            kind = "grey"
        M.run_fmt(r, size, bold or kind in ("b", "green", "hl", "red"))
        if color_override is not None and kind is None:
            r.font.color.rgb = color_override
        if kind == "hl":
            r.font.color.rgb = RED
            if slide_no:
                MISSING.append((slide_no, txt))
        elif kind == "red":
            r.font.color.rgb = RISK
        elif kind == "green":
            r.font.color.rgb = GREEN
        elif kind == "grey":
            r.font.color.rgb = GREY
        elif kind == "src":
            key = re.match(r"\[([A-Z]\d+)", txt).group(1)
            if key in SOURCES:
                r.hyperlink.address = SOURCES[key][1]
                r.font.color.rgb = M.DARK
            r.font.size = Pt(max(size - 2, 8))


def plain(text):
    return TOKEN.sub(lambda m: m.group(0).strip("*!^+~"), text or "")


def text_lines(text, width_in, size, bold=False):
    """Kiek eilučių užims tekstas (Arial metrika – Liberation Sans)."""
    f = ImageFont.truetype(FONT_B if bold else FONT_R, int(size * 10))
    max_w = width_in * 72 * 10
    n = 0
    for para in (text if isinstance(text, list) else [text]):
        words, line = plain(para).split(" "), ""
        n_para = 1
        for w in words:
            t = (line + " " + w).strip()
            if f.getlength(t) <= max_w:
                line = t
            else:
                n_para += 1
                line = w
        n += n_para
    return n


def bullet(p, size):
    """„- “ eilutės pradžioje -> tikras ženklelis su įtrauka."""
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Pt(size * 0.95))))
    pPr.set("indent", str(-int(Pt(size * 0.95))))
    for t in ("a:buNone", "a:buChar", "a:buFont"):
        for el in pPr.findall(qn(t)):
            pPr.remove(el)
    etree.SubElement(pPr, qn("a:buFont"), typeface="Arial")
    etree.SubElement(pPr, qn("a:buChar"), char="•")


def para_fill(tf, content, size, bold, slide_no, color=None):
    paras = content if isinstance(content, list) else [content]
    for k, line in enumerate(paras):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        M.no_space(p)
        if isinstance(line, str) and line.startswith("- "):
            bullet(p, size)
            line = line[2:]
        add_runs(p, line, size, bold, slide_no, color)
    if not content:
        tf.paragraphs[0].font.size = Pt(size)


def table(slide, rows, box, widths, size=12, bold_cols=(), bold_rows=(), bold_cells=(), min_h=None, slide_no=None, max_bottom=BOTTOM, fit=True,
          head=False, shade_col=None):
    """Bendro stiliaus lentelė (šablono): balta, plonos pilkos linijos, be užpildų. Sumažina šriftą, kol telpa."""
    x, y, w, h = box
    pad = 0.16
    while True:
        hs = []
        for i, row in enumerate(rows):
            lines = max(text_lines(c, widths[j] - pad, size, j in bold_cols or i in bold_rows) for j, c in enumerate(row))
            hs.append(max(lines * size * 1.2 / 72 + 0.08, min_h or 0))
        if not fit or y + sum(hs) <= max_bottom or size <= 9:
            break
        size -= 0.5
    if y + sum(hs) > max_bottom:
        print(f"!! skaidrė {slide_no}: lentelė netelpa ({y + sum(hs):.2f} > {max_bottom})")
    gs = slide.shapes.add_table(len(rows), len(widths), Inches(x), Inches(y), Inches(sum(widths)), Inches(sum(hs)))
    tbl = gs.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
    sid.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"
    for j, cw in enumerate(widths):
        tbl.columns[j].width = Inches(cw)
    for i, row in enumerate(rows):
        tbl.rows[i].height = Inches(hs[i])
        for j, c in enumerate(row):
            cell = tbl.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            M.cell_border(cell)
            cell.text_frame.word_wrap = True
            dark = head and i == 0
            if dark or (shade_col is not None and j == shade_col):
                tcPr = cell._tc.get_or_add_tcPr()
                for f in tcPr.findall(qn("a:solidFill")):
                    tcPr.remove(f)
                sf = etree.SubElement(tcPr, qn("a:solidFill"))
                etree.SubElement(sf, qn("a:srgbClr"), val="191824" if dark else "EEF5ED")
                # užpildas turi būti po linijų elementų
                tcPr.remove(sf)
                tcPr.append(sf)
            para_fill(cell.text_frame, c, size, dark or j in bold_cols or i in bold_rows or (i, j) in bold_cells, slide_no, WHITE if dark else None)
    return y + sum(hs)


def body(slide, box, paras, size=14, slide_no=None, fit=True):
    """Teksto blokas (be ženklelių); paras – sąrašas eilučių. Sumažina šriftą, kol telpa."""
    x, y, w, h = box
    while fit and size > 10:
        need = sum(text_lines(p, w - 0.1, size) for p in paras) * size * 1.25 / 72 + len(paras) * size * 0.35 / 72
        if need <= h:
            break
        size -= 0.5
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    for k, line in enumerate(paras):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        M.no_space(p)
        p.space_after = Pt(size * 0.35)
        if line.startswith("- "):
            bullet(p, size)
            line = line[2:]
        add_runs(p, line, size, False, slide_no)
    return tb


def card(slide, box, label, value, note=None, accent="191824", slide_no=None, vsize=22):
    """Kortelė svarbiausiam skaičiui: plona rėmelio linija, spalvota viršutinė juosta."""
    from pptx.enum.shapes import MSO_SHAPE
    x, y, w, h = box
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    sh.line.color.rgb = RGBColor(0xD0, 0xD0, 0xD0); sh.line.width = Pt(0.75)
    sh.shadow.inherit = False
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.06))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor.from_string(accent); bar.line.fill.background()
    bar.shadow.inherit = False
    tb = slide.shapes.add_textbox(Inches(x + 0.12), Inches(y + 0.12), Inches(w - 0.24), Inches(h - 0.18))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]; M.no_space(p); add_runs(p, label, 10, slide_no=slide_no)
    for r in p.runs:
        r.font.color.rgb = RGBColor(0x6B, 0x6B, 0x6B)
    p = tf.add_paragraph(); M.no_space(p); add_runs(p, value, vsize, True, slide_no)
    if note:
        p = tf.add_paragraph(); M.no_space(p); add_runs(p, note, 9.5, slide_no=slide_no)
    return sh


def heading(slide, box, text, size=13, color="191824"):
    """Bloko antraštė su plona linija apačioje."""
    x, y, w, h = box
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.32))
    tf = tb.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]; M.no_space(p); add_runs(p, text, size, True)
    for r in p.runs:
        r.font.color.rgb = RGBColor.from_string(color)
    ln = slide.shapes.add_connector(1, Inches(x), Inches(y + 0.33), Inches(x + w), Inches(y + 0.33))
    ln.line.color.rgb = RGBColor.from_string(color); ln.line.width = Pt(1)


def column_chart(slide, box, cats, series, title=None, size=10, number_format='0.0"%"', labels=True, palette=None, axis_format='0"%"'):
    cd = CategoryChartData()
    cd.categories = cats
    for name, vals in series:
        cd.add_series(name, vals)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, *(Inches(v) for v in box), cd)
    ch = gf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(size)
    ch.font.size = Pt(size)
    ch.font.color.rgb = M.DARK
    if title:
        ch.has_title = True
        ch.chart_title.text_frame.text = title
        ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(size + 1)
        ch.chart_title.text_frame.paragraphs[0].runs[0].font.bold = False
    else:
        ch.has_title = False
    palette = palette or ["191824", "87C081", "4561A7", "DCB241", "AFAFAF", "D86D64"]
    plot = ch.plots[0]
    plot.gap_width = 80
    labels = True  # reikšmės virš stulpelių privalomos visiems DD (Vytis 2026-11-01); labels=False ignoruojamas
    plot.has_data_labels = labels
    if labels:  # reikšmės virš stulpelių – visada (Vytis 2026-11-01)
        from pptx.enum.chart import XL_LABEL_POSITION
        plot.data_labels.font.size = Pt(max(size - 2, 7))
        plot.data_labels.number_format = number_format
        plot.data_labels.number_format_is_linked = False
        plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    for k, s in enumerate(plot.series):
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = RGBColor.from_string(palette[k % len(palette)])
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = RGBColor(0xE6, 0xE6, 0xE6)
    va.tick_labels.font.size = Pt(size)
    va.tick_labels.number_format = axis_format
    va.tick_labels.number_format_is_linked = False
    va.format.line.fill.background()
    from pptx.enum.chart import XL_TICK_LABEL_POSITION
    ch.category_axis.tick_labels.font.size = Pt(size)
    ch.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    return ch


def line_chart(slide, box, cats, series, title=None, size=10, axis_format='0', palette=None, widths=None):
    """Linijinė diagrama (kaupiamoji grąža, slenkantys rodikliai); be duomenų žymių."""
    cd = CategoryChartData()
    cd.categories = cats
    for name, vals in series:
        cd.add_series(name, vals)
    ch = slide.shapes.add_chart(XL_CHART_TYPE.LINE, *(Inches(v) for v in box), cd).chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(size)
    ch.font.size = Pt(size)
    ch.font.color.rgb = M.DARK
    ch.has_title = bool(title)
    if title:
        ch.chart_title.text_frame.text = title
        ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(size + 1)
        ch.chart_title.text_frame.paragraphs[0].runs[0].font.bold = False
    palette = palette or ["191824", "87C081", "4561A7", "DCB241", "AFAFAF", "D86D64"]
    for k, s in enumerate(ch.plots[0].series):
        s.smooth = False
        s.marker.style = None
        from pptx.enum.chart import XL_MARKER_STYLE
        s.marker.style = XL_MARKER_STYLE.NONE
        s.format.line.color.rgb = RGBColor.from_string(palette[k % len(palette)])
        s.format.line.width = Pt((widths or [2.25] * 9)[k])
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = RGBColor(0xE6, 0xE6, 0xE6)
    va.tick_labels.font.size = Pt(size)
    va.tick_labels.number_format = axis_format
    va.tick_labels.number_format_is_linked = False
    va.format.line.fill.background()
    from pptx.enum.chart import XL_TICK_LABEL_POSITION
    ch.category_axis.tick_labels.font.size = Pt(size)
    ch.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    return ch


def signed_bars(slide, box, cats, vals, title=None, size=10, number_format='+0.0;-0.0;0.0', axis_format='0', pos="87C081", neg="C01E1E"):
    """Viena serija, stulpelis žalias (> 0) arba raudonas (< 0), reikšmė virš/po stulpeliu, be legendos.
    Skirtumui (fondas − indeksas, pp) – aiškiau nei dvi serijos šalia (Vytis 2026-11-12: „grafikas netvarkingas ir neaiškus“)."""
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series(title or "value", vals)
    ch = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, *(Inches(v) for v in box), cd).chart
    ch.has_legend = False
    ch.font.size = Pt(size)
    ch.font.color.rgb = M.DARK
    ch.has_title = bool(title)
    if title:
        ch.chart_title.text_frame.text = title
        ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(size + 1)
        ch.chart_title.text_frame.paragraphs[0].runs[0].font.bold = False
    from pptx.enum.chart import XL_LABEL_POSITION, XL_TICK_LABEL_POSITION
    plot = ch.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(max(size - 1, 7))
    plot.data_labels.number_format = number_format
    plot.data_labels.number_format_is_linked = False
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    ser = plot.series[0]
    for i, v in enumerate(vals):
        pt = ser.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = RGBColor.from_string(pos if (v or 0) >= 0 else neg)
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = RGBColor(0xE6, 0xE6, 0xE6)
    va.tick_labels.font.size = Pt(size)
    va.tick_labels.number_format = axis_format
    va.tick_labels.number_format_is_linked = False
    va.format.line.fill.background()
    ch.category_axis.tick_labels.font.size = Pt(size)
    ch.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    return ch


def tick_skip(chart, n):
    """Kas n-ta kategorijos etiketė (dienos duomenims linijinėje diagramoje)."""
    from lxml import etree
    from pptx.oxml.ns import qn
    ax = chart._chartSpace.find(".//" + qn("c:catAx"))
    nm = ax.find(qn("c:noMultiLvlLbl"))  # schemos tvarka: … lblOffset, tickLblSkip, tickMarkSkip, noMultiLvlLbl
    for tag in ("c:tickLblSkip", "c:tickMarkSkip"):
        el = ax.find(qn(tag))
        if el is None:
            el = etree.Element(qn(tag))
            if nm is not None:
                nm.addprevious(el)
            else:
                ax.append(el)
        el.set("val", str(n))
    return chart


def picture(slide, path, box):
    x, y, w, h = box
    return slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h) if h else None)


# ------------------------------------------------------------------ 2026-11-01 (AQR DD peržiūra): bendros taisyklės visiems Funds DD
# Šablono skaidrės (SEB Liquid Funds DD Template) – griežtai šia tvarka, po vieną; nenaudojamos – nerodomos (Vytis 2026-11-12).
TEMPLATE = ["Fund name", "Conclusion", "Fund overview  (index funds)", "Fund overview (mutual funds)",
            "Index overview (if index is different than used internally)", "ESG overview (excluding climate future index fund)",
            "ESG overview (excluding climate future index fund)", "ESG overview (climate future index fund)", "Performance",
            "Investment team", "Investment strategy & process", "Current portfolio characteristics", "Other"]
DIVIDER = "Additional analysis"
MUTED = RGBColor(0x5A, 0x5A, 0x5A)


def data_note(slide, box, text, size=9):
    """Po kiekviena lentele / diagrama: „Data:“ – duomenų rinkinys (šaltinis, klasė, laikotarpis, paėmimo data),
    skaičiavimas žodžiais ir Excel lapas; šaltinių žymės paspaudžiamos."""
    x, y, w, h = box
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h or 0.4))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    for k, line in enumerate(text if isinstance(text, list) else [text]):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        M.no_space(p)
        add_runs(p, line, size, color_override=MUTED)
        if k == 0 and p.runs and not line.startswith("Data:"):
            pass
    return tb


def heat_color(v, scale=4.0):
    """Žalia (teigiama) – balta – raudona (neigiama); scale – reikšmė, ties kuria spalva sočiausia."""
    if v is None:
        return "F2F2F2"
    t = max(-1.0, min(1.0, v / scale))
    if t >= 0:
        r, g, b = 255 - int(t * (255 - 0x87)), 255 - int(t * (255 - 0xC0)), 255 - int(t * (255 - 0x81))
    else:
        t = -t
        r, g, b = 255 - int(t * (255 - 0xE0)), 255 - int(t * (255 - 0x8C)), 255 - int(t * (255 - 0x86))
    return f"{r:02X}{g:02X}{b:02X}"


def heatmap(slide, box, row_labels, col_labels, values, size=9, scale=4.0, fmt="{:+.1f}", label_w=1.3, slide_no=None):
    """Mėnesių šilumos žemėlapis (metai × mėnesiai): langelio spalva pagal reikšmę, skaičius langelyje."""
    x, y, w, h = box
    cw = (w - label_w) / len(col_labels)
    rh = min(0.32, h / (len(row_labels) + 1))
    gs = slide.shapes.add_table(len(row_labels) + 1, len(col_labels) + 1, Inches(x), Inches(y), Inches(w), Inches(rh * (len(row_labels) + 1)))
    tbl = gs.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0"); tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
    sid.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"
    tbl.columns[0].width = Inches(label_w)
    for j in range(len(col_labels)):
        tbl.columns[j + 1].width = Inches(cw)
    for i in range(len(row_labels) + 1):
        tbl.rows[i].height = Inches(rh)
        for j in range(len(col_labels) + 1):
            cell = tbl.cell(i, j)
            cell.margin_left = cell.margin_right = Inches(0.03)
            cell.margin_top = cell.margin_bottom = Inches(0.01)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            M.cell_border(cell)
            if i == 0:
                txt, fill, bold = ("" if j == 0 else col_labels[j - 1]), "191824", True
            elif j == 0:
                txt, fill, bold = row_labels[i - 1], "FFFFFF", True
            else:
                v = values[i - 1][j - 1]
                txt, fill, bold = ("" if v is None else fmt.format(v)), heat_color(v, scale), False
            tcPr = cell._tc.get_or_add_tcPr()
            for f in tcPr.findall(qn("a:noFill")) + tcPr.findall(qn("a:solidFill")):
                tcPr.remove(f)  # cell_border jau įdėjo baltą užpildą – PowerPoint skaito pirmą
            sf = etree.SubElement(tcPr, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr"), val=fill)
            p = cell.text_frame.paragraphs[0]
            from pptx.enum.text import PP_ALIGN
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            add_runs(p, txt, size, bold, slide_no, WHITE if i == 0 else None)
    return y + rh * (len(row_labels) + 1)


def pdf_crop(pdf, page, find, out_png, pad=(30, 12, 30, 40), dpi=150, width_pt=None):
    """Oficialaus dokumento iškarpa vietoj ekrano kopijos: randa tekstą `find` PDF puslapyje `page` (1-based)
    ir išsaugo PNG su aplinka. pad = (kairė, viršus, dešinė, apačia) taškais. Grąžina (png, (w_pt, h_pt))."""
    import html
    import subprocess
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        xml = Path(td) / "b.html"
        subprocess.run(["pdftotext", "-f", str(page), "-l", str(page), "-bbox-layout", str(pdf), str(xml)], check=True)
        doc = xml.read_text(errors="ignore")
    pw, ph = (float(v) for v in re.search(r'<page width="([\d.]+)" height="([\d.]+)"', doc).groups())
    words = [(float(a), float(b), float(c), float(d), html.unescape(t)) for a, b, c, d, t in
             re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', doc)]
    toks = find.split()
    for i in range(len(words) - len(toks) + 1):
        if all(words[i + k][4].strip().lower() == toks[k].lower() for k in range(len(toks))):
            sel = words[i:i + len(toks)]
            break
    else:
        raise ValueError(f"'{find}' nerasta {pdf} p.{page}")
    # pasuktas puslapis (pvz. metinės ataskaitos lentelė 270°): pdftotext žodžių koordinatės jau pasuktos (kaip rodoma),
    # bet <page width/height> – nepasukti; be sukeitimo iškarpa nukerpama ties ~597 pt (iShares Poland DD 2026-11-23)
    info = subprocess.run(["pdfinfo", "-f", str(page), "-l", str(page), str(pdf)], capture_output=True, text=True).stdout
    m_ = re.search(r"Page\s+%d rot:\s+(\d+)" % page, info)
    if m_ and int(m_.group(1)) % 180 == 90 and pw < ph:
        pw, ph = ph, pw
    x0 = max(0, min(w[0] for w in sel) - pad[0]); y0 = max(0, min(w[1] for w in sel) - pad[1])
    x1 = min(pw, (x0 + width_pt) if width_pt else max(w[2] for w in sel) + pad[2]); y1 = min(ph, max(w[3] for w in sel) + pad[3])
    k = dpi / 72
    out_png = Path(out_png)
    subprocess.run(["pdftoppm", "-f", str(page), "-l", str(page), "-r", str(dpi), "-png", "-singlefile",
                    "-x", str(int(x0 * k)), "-y", str(int(y0 * k)), "-W", str(int((x1 - x0) * k)), "-H", str(int((y1 - y0) * k)),
                    str(pdf), str(out_png.with_suffix(""))], check=True)
    return out_png, (x1 - x0, y1 - y0)


def fund_pages(pdf, *phrases, first=1, last=None):
    """Skėčio dokumentuose (pvz. iShares V plc metinė ataskaita ~750 psl.) – puslapiai, kuriuose yra VISOS frazės
    (be didžiųjų/mažųjų raidžių skirtumo). Pirma surask fondo puslapius, tada skaityk tik juos. Grąžina [psl. nr. (1-based)]."""
    import subprocess
    txt = subprocess.run(["pdftotext", "-layout", "-f", str(first)] + (["-l", str(last)] if last else []) + [str(pdf), "-"],
                         capture_output=True, text=True).stdout
    low = [" ".join(p.lower().split()) for p in txt.split("\f")]
    return [first + i for i, t in enumerate(low) if all(" ".join(ph.lower().split()) in t for ph in phrases)]


def evidence_picture(slide, png, size_pt, box, caption, slide_no=None):
    """Dokumento iškarpa su rėmeliu ir parašu „Copy of the official document: … [S#]“."""
    x, y, w, h = box
    wpt, hpt = size_pt
    sc = min(w / (wpt / 72), h / (hpt / 72))
    pic = slide.shapes.add_picture(str(png), Inches(x), Inches(y), Inches(wpt / 72 * sc), Inches(hpt / 72 * sc))
    pic.line.color.rgb = RGBColor(0xAF, 0xAF, 0xAF); pic.line.width = Pt(0.75)
    data_note(slide, (x, y + hpt / 72 * sc + 0.05, w, 0.35), caption, 9)
    return y + hpt / 72 * sc + 0.4


def ucits_bbg_line(isin, proven=None):
    """Bloomberg FUND_EURO_DIRECT_UCITS patikra reikalinga tik tada, kai UCITS statusas neįrodytas dokumentuose (Vytis 2026-11-01)."""
    if proven:
        return f"UCITS statusas įrodytas fondo dokumentuose ({proven}) – Bloomberg patikra =@BDP(\"{isin} isin\";\"FUND_EURO_DIRECT_UCITS\") nereikalinga."
    return f"Bloomberg patikra (PSPKĮ 49.1): =@BDP(\"{isin} isin\";\"FUND_EURO_DIRECT_UCITS\") – !!įrašyti rezultatą (UCITS statusas dokumentuose nerastas)!!"


MEDIA_HEAD = ["Date", "Published by", "What happened", "Who is involved", "Status / outcome", "Relevance for the fund", "Source"]
MEDIA_W = (0.85, 1.55, 3.4, 1.55, 2.0, 1.75, 1.03)


class Deck:
    """Pradinis pristatymas pagal šabloną (tema, master, maketai) ir skaidrių pridėjimas."""

    def __init__(self):
        from pptx import Presentation
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        M.set_theme(prs)
        M.master_setup(prs)
        self.lay = prs.slide_layouts[5]
        self.blank = prs.slide_layouts[6]
        for lay in (self.lay, self.blank):
            for ph in list(lay.placeholders):
                n = str(ph.placeholder_format.type)
                if "TITLE" in n:
                    ph.left, ph.top, ph.width, ph.height = (Inches(v) for v in M.TITLE_BOX)
                elif "DATE" in n or "FOOTER" in n:
                    ph._element.getparent().remove(ph._element)
                elif "SLIDE_NUMBER" in n:
                    ph.left, ph.top, ph.width, ph.height = Inches(0.602), Inches(6.936), Inches(0.713), Inches(0.399)
        self.prs = prs

    @property
    def n(self):
        return len(self.prs.slides)

    def slide(self, ttl=None, sub=None, box=M.TITLE_BOX, blank=False):
        s = self.prs.slides.add_slide(self.blank if blank else self.lay)
        if ttl is not None and not blank:
            M.title(s, ttl, box, sub)
        M.slide_number(s, self.lay)
        return s

    def na_slide(self, ttl, reason):
        """PANAIKINTA (Vytis 2026-11-12, kolegos pastaba AQR DD): nenaudojamų šablono skaidrių nerodome visai.
        Netinkama šablono skaidrė (pvz. „Fund overview (index funds)“ aktyviam fondui) tiesiog praleidžiama."""
        raise RuntimeError(f"na_slide('{ttl}'): unused template slides are not shown – leave the slide out (rule 2026-11-12)")

    def divider(self, text=DIVIDER, sub="Slides below supplement the template slides above; each title names the template slide it belongs to."):
        """Skyriklis tarp šablono dalies ir papildomų skaidrių."""
        s = self.slide(blank=True)
        tb = s.shapes.add_textbox(Inches(0.602), Inches(2.9), Inches(12.129), Inches(1.6))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; M.no_space(p)
        r = p.add_run(); r.text = text; M.run_fmt(r, 32, major=True)
        p = tf.add_paragraph(); M.no_space(p); add_runs(p, sub, 16, color_override=MUTED)
        return s

    def save(self, out):
        self.prs.save(out)
        return out


# ------------------------------------------------------------------ šaltiniai ir dokumentų ZIP (2026-11-01)
# Taisyklė (savininkas 2026-11-01, AQR DD): šaltinių sąrašas ir „Dokumentai (ZIP)“ sutampa 1:1.
# - Šaltiniu dedamas tik dokumentas, kurio turinys perskaitytas ir kurį savininkas gali atsidaryti: jokių antraščių,
#   mokamo turinio (paywall), prisijungimo ar svetainės pradinių puslapių. Neperskaitytas -> ne šaltinis.
# - URL turi vesti tiesiai į dokumentą; jei tiesioginės nuorodos nėra, bet failą pateikė savininkas -> `local` (į ZIP failas),
#   o ne svetainės pradinis puslapis.
# - Savininko pateikti failai (D#, ir KIID / prospektas be viešos nuorodos) - į ZIP su ta pačia numeracija (`local`),
#   užšifruotai (dd_vault `_job` -> kind received; dashboard'as sujungia su atsisiųstais šaltiniais).
# - Morningstar jungties duomenys (M#) ir savi skaičiavimai (C#) - įrodymų ZIP; sankcijų įrankiai (X#) - tik nuorodos.
# - update.yml (tools/dd_sources.py) tikrina atsisiųstą turinį ir atmeta netinkamus; `expect` - tekstas, kuris turi būti.
PAYWALL = ("bloomberg.com", "ft.com", "wsj.com", "investmentofficer.com", "reuters.com/subscribe")


HEADLINE_NOTE = ("Headline only – the full text is behind a login / paywall and was not read. Listed for reference; "
                 "not used as evidence (facts are supported by the other sources cited next to it).")


def where(key, url=None, local=None, headline=False):
    """Kur šaltinis yra – rašoma Sources skaidrėje, kad sąrašą būtų galima sutikrinti su ZIP."""
    if headline:
        return "headline only, full text not read (login / paywall)"
    if key.startswith("X"):
        return "link only (search tool)"
    if key.startswith(("M", "C")):
        return "Evidence ZIP"
    return "Documents ZIP" + (" (received copy)" if local else "")


def sources_json(path, sources, docs=None, local=None, expect=None, skip=("X", "M", "C"), headline=()):
    """dd_vault.py sources/add --sources failas: [{key,title,url}|{key,title,local,url?}] + `expect`.
    Neleidžia mokamo turinio / pradinių puslapių be vietinės kopijos (AssertionError – geriau sustoti nei įdėti šlamštą).
    `headline` – raktai, kurių perskaityta tik antraštė (savininkas 2026-11-01: „antraštė ir nuoroda geriau negu nieko“):
    į ZIP – tekstinis failas su antrašte, nuoroda ir HEADLINE_NOTE, ne prisijungimo langas; įrodymu nelaikoma."""
    from urllib.parse import urlparse
    local, expect, out = local or {}, expect or {}, []
    for k, (t, u) in list(sources.items()) + [(k, (t, None)) for k, t in (docs or {}).items()]:
        if k.startswith(skip):
            continue
        e = {"key": k, "title": t}
        if k in headline:
            e.update(url=u, note=f"{HEADLINE_NOTE}\nLink: {u}")
        elif k in local:
            e["local"] = str(local[k])
            if u:
                e["url"] = u
        else:
            assert u, f"{k}: nėra nei URL, nei vietinio failo"
            assert urlparse(u).path.strip("/"), f"{k}: svetainės pradinis puslapis ({u}) – reikia dokumento nuorodos arba local failo"
            assert not any(d in u for d in PAYWALL), f"{k}: mokamas turinys ({u}) – šaltiniu netinka, jei neperskaitytas visas"
            e["url"] = u
        if k in expect:
            e["expect"] = list(expect[k])
        out.append(e)
    Path(path).write_text(json.dumps(out, ensure_ascii=False, indent=1))
    return out


# ------------------------------------------------------------------ PDF iš HTML / teksto (v4.1, 2026-11-12: ZIP – tik PDF ir Excel)
def html_pdf(html, out):
    """HTML → PDF (Chromium „print to PDF“, A4; @page paraštės – HTML'e)."""
    import glob
    import subprocess
    import tempfile
    chrome = (sorted(glob.glob("/opt/pw-browsers/chromium*/chrome-linux*/chrome")) or ["chromium"])[-1]
    with tempfile.TemporaryDirectory() as td:
        f = Path(td) / "page.html"
        f.write_text(html.replace("<style>", "<style>@page{size:A4;margin:14mm}", 1), encoding="utf-8")
        subprocess.run([chrome, "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={out}", f.as_uri()],
                       check=True, capture_output=True, timeout=300)
    return out


def txt_pages_pdf(txt, out, title):
    """Teksto kopija su spausdintais puslapių numeriais (eilutė „N“ pradeda N puslapį) → PDF, kuriame PDF puslapis = spausdintas
    puslapis (citata [P1 p.N] atsiverčiama N puslapyje). Kiekvieno puslapio viršuje – dokumento pavadinimas ir puslapio numeris."""
    import html as H
    import re as _re
    pages, cur = {}, None
    for line in Path(txt).read_text(encoding="utf-8").splitlines():
        m = _re.fullmatch(r"\s*(\d{1,3})\s*", line)
        if m and (cur is None or int(m.group(1)) == cur + 1):
            cur = int(m.group(1)); pages[cur] = []; continue
        if cur:
            pages[cur].append(line)
    assert list(pages) == list(range(1, len(pages) + 1)), "puslapių numeriai ne iš eilės"
    def fs(ls):  # šrifto dydis, kad vienas spausdintas puslapis tilptų į vieną PDF puslapį
        n = sum(max(1, -(-len(x) // 125)) for x in ls)
        return min(7.6, 7.6 * 70 / max(n, 1))
    body = "".join(f'<section><div class="h">{H.escape(title)} – page {n} of {len(pages)} (page number as printed)</div>'
                   f'<pre style="font-size:{fs(ls):.2f}pt">{H.escape(chr(10).join(ls).strip())}</pre></section>' for n, ls in pages.items())
    css = ("body{font-family:Arial,sans-serif;margin:0} section{page-break-after:always;height:262mm;overflow:hidden} section:last-child{page-break-after:auto}"
           ".h{font-size:8pt;color:#777;border-bottom:1px solid #ccc;margin-bottom:6px;padding-bottom:3px}"
           "pre{font-family:Arial,sans-serif;font-size:7.6pt;white-space:pre-wrap;line-height:1.28;margin:0}")
    html_pdf(f"<html><head><meta charset='utf-8'><style>{css}</style></head><body>{body}</body></html>", out)
    return len(pages)


# ------------------------------------------------------------------ sankcijų patikra (savininkas 2026-11-13)
# Taisyklė: kiekvienam tikrinamam vardui – atskiras ekrano nuotraukos langelis KIEKVIENOJE sistemoje
# (vardai × sistemos, pvz. 3 × 3 = 9). Pagrindinė skaidrė – žingsniai + rezultatų lentelė (vardas × sistema);
# po jos – po vieną skaidrę kiekvienam vardui su 3 dideliais langeliais (nuotrauka turi būti įskaitoma).
# Tikrinamos tik įmonės (savininkas 2026-11-13): valdymo įmonė + investicijų valdytojas. Pats fondas (FCP / skėtis)
# netikrinamas – sankcijos taikomos įmonei, ne fondui (šablonas: „fund management company“).
SANCTION_SYSTEMS = (("SEB eNameChecker", "X1"), ("OFAC Sanctions List Search", "X2"), ("EU Sanctions Map", "X3"))


def sanction_keys(names):
    """ORDER raktai: 'sanctions' + po vieną kiekvienam vardui ('sanctions_1' …)."""
    return ["sanctions"] + [f"sanctions_{i}" for i in range(1, len(names) + 1)]


def sanctions_slides(d, names, systems=SANCTION_SYSTEMS, slide_no=None, accent="AFAFAF"):
    """names = [(vardas, vaidmuo su šaltiniu), …], pvz. ("FundRock Management Company S.A.", "management company, Luxembourg [P1 p.7]").
    Sukuria 1 + len(names) skaidres; grąžina jų skaičių."""
    n0 = d.n
    s = d.slide("Compliance with the list of sanctions")
    first = slide_no or (n0 + 1)
    table(s, [
        ["Step", "What to do"],
        ["1. Search these names", [f"- **{a}** – {b}" for a, b in names]],
        ["2. In these lists", [f"- {sy} [{k}]" for sy, k in systems]],
        ["3. Paste results", "!!Analyst: one screenshot per name and list – " + (f"slide {first + 1} " if len(names) == 1 else f"slides {first + 1}–{first + len(names)} ") +
                             f"({len(names)} × {len(systems)} = {len(names) * len(systems)} boxes); mark the result below!!"],
    ], (0.602, 1.40, 0, 0), (2.6, 9.393), 12, bold_cols=(0,), head=True, slide_no=first)
    w0 = 3.4
    wk = (12.0 - w0) / len(systems)
    rows = [["Name / list"] + [sy for sy, _ in systems]]
    for i, (a, _) in enumerate(names, 1):
        rows.append([f"**{a}** (slide {first + i})"] + ["!!Match: Yes / No!!"] * len(systems))
    table(s, rows, (0.602, 4.35, 0, 0), (w0, *[wk] * len(systems)), 11, head=True, slide_no=first)
    for i, (a, b) in enumerate(names, 1):
        s = d.slide(f"Sanctions screenshots – name {i} of {len(names)}")
        body(s, (0.602, 1.30, 12.129, 0.4), [f"**{a}** – {b}"], 16, slide_no=first + i)
        w = (12.129 - 0.13 * (len(systems) - 1)) / len(systems)
        for k, (sy, key) in enumerate(systems):
            card(s, (0.602 + k * (w + 0.13), 1.85, w, 4.85), f"Screenshot {i}.{k + 1} – [{key}]", sy,
                 f"Paste the search result for “{a}” here (search date visible).", accent, slide_no=first + i, vsize=16)
    return d.n - n0


def bare_refs(pptx_path, keys=("C", "D")):
    """Patikra prieš pateikiant (savininkas 2026-11-13): [C1] / [D3] be puslapio, lapo ar langelio.
    Grąžina [(skaidrė, žymė, teksto ištrauka)]; turi būti tuščia. Žymė be „p.“, lapo pavadinimo ar langelio = klaida."""
    from pptx import Presentation
    bad, pat = [], re.compile(r"\[(%s)\d+\]" % "|".join(keys))
    for i, sl in enumerate(Presentation(pptx_path).slides, 1):
        for sh in sl.shapes:
            txts = []
            if sh.has_text_frame:
                txts.append(sh.text_frame.text)
            if getattr(sh, "has_table", False) and sh.has_table:
                txts += [c.text for r in sh.table.rows for c in r.cells]
            for t in txts:
                for m in pat.finditer(t):
                    if t[m.end():m.end() + 7].strip().startswith("sheet"):   # „Excel [C1] sheet “S07 …”“ po lentele – lapas nurodytas
                        continue
                    bad.append((i, m.group(0), t[max(0, m.start() - 70):m.end()]))
    return bad
