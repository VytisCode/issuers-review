"""Vienas skaičiavimų Excel failas DD / peržiūrai (savininkas 2026-11-12: kiekvienai skaidrės lentelei ar diagramai – lygiai
tokia pati lentelė / diagrama Excel'yje, su visais skaičiavimais ir aiškia duomenų kilme).

Taisyklės:
- lapas „Start here“ – kaip skaityti, spalvos, lapų sąrašas (skaidrė → lapas);
- lapai „S07 Performance“ … – po vieną kiekvienai skaidrei su lentele / diagrama; viršuje skaidrės pavadinimas,
  lentelė tokia pati kaip skaidrėje, skaičiai – formulės į duomenų lapus; diagrama – iš tos lentelės;
- duomenų lapai „Data …“ – įvestys (mėlynas šriftas), prie KIEKVIENOS eilutės ar stulpelio – šaltinis
  (dokumentas, puslapis / lapas ir eilutė, Morningstar datapoint, SecId, nuoroda, paėmimo data);
- formulės – juodu šriftu; po perskaičiavimo (LibreOffice) Excel reikšmės lyginamos su skaidrės skaičiais (`check`).
"""
import shutil
import subprocess
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as CL

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HF = Font(bold=True, color="FFFFFF")
HEAD = PatternFill("solid", fgColor="191824")
TITLE = Font(bold=True, size=14)
NOTE = Font(italic=True, color="555555")
WRAP = Alignment(wrap_text=True, vertical="top")


def q(name):
    return "'" + name.replace("'", "''") + "'"


class Sheet:
    def __init__(self, book, ws):
        self.book, self.ws, self.r = book, ws, 1

    def text(self, t, font=None, merge=0):
        c = self.ws.cell(self.r, 1, t)
        if font:
            c.font = font
        self.r += 1
        return c

    def gap(self, n=1):
        self.r += n

    def header(self, cols, row=None):
        row = row or self.r
        for j, c in enumerate(cols, 1):
            x = self.ws.cell(row, j, c)
            x.font, x.fill, x.alignment = HF, HEAD, Alignment(wrap_text=True, vertical="center")
        if row == self.r:
            self.r += 1
        return row

    def row(self, vals, fmt=None, expect=None, tol=None, inputs=None):
        """vals: str su '=' – formulė; kiti – reikšmės. inputs – stulpelių indeksai (0-based), kurie yra įvestys
        (mėlynai); None = visi ne formulės skaičiai. expect: {stulp. indeksas: skaidrės reikšmė} – patikrai."""
        r = self.r
        for j, v in enumerate(vals, 1):
            if v is None:
                continue
            c = self.ws.cell(r, j, v)
            isf = isinstance(v, str) and v.startswith("=")
            if (inputs is None and not isf and isinstance(v, (int, float))) or (inputs is not None and (j - 1) in inputs):
                c.font = BLUE
            f = (fmt[j - 1] if j - 1 < len(fmt) else None) if isinstance(fmt, (list, tuple)) else fmt
            if f and (isf or isinstance(v, (int, float))):
                c.number_format = f
        for j, (val, t) in (expect or {}).items():
            self.book.checks.append((self.ws.title, f"{CL(j + 1)}{r}", val, t if t is not None else (tol or 1e-6)))
        self.r += 1
        return r

    def tag(self, row, col, val=None, tol=1e-6, key="C1"):
        """Tiksli nuoroda tekstui (savininkas 2026-11-13: „[C1] – nežinau, kur ieškoti“):
        grąžina „[C1 “S07 Performance” B26]“; jei duota skaidrės reikšmė – įtraukia į automatinę patikrą."""
        if val is not None:
            self.book.checks.append((self.ws.title, f"{CL(col)}{row}", val, tol))
        return f"[{key} “{self.ws.title}” {CL(col)}{row}]"

    def cell(self, row, col):
        """Absoliuti nuoroda kitam lapui: 'Lapas'!$C$5"""
        return f"{q(self.ws.title)}!${CL(col)}${row}"

    def ref(self, row, col):
        return f"{q(self.ws.title)}!{CL(col)}{row}"

    def rng(self, r0, r1, col):
        return f"{q(self.ws.title)}!${CL(col)}${r0}:${CL(col)}${r1}"

    def widths(self, *w):
        for j, x in enumerate(w, 1):
            self.ws.column_dimensions[CL(j)].width = x

    def chart(self, kind, cats, series, anchor, title, ytitle=None, w=22, h=9, stacked=False, colors=None, fmt=None):
        """cats = (min_row, max_row, col); series = [(header_row, min_row, max_row, col)]; kind bar|line."""
        ch = BarChart() if kind == "bar" else LineChart()
        if kind == "bar":
            ch.type = "col"
            if stacked:
                ch.grouping, ch.overlap = "stacked", 100
        ch.title, ch.height, ch.width = title, h, w
        if ytitle:
            ch.y_axis.title = ytitle
        for hr, r0, r1, col in series:
            ch.add_data(Reference(self.ws, min_col=col, min_row=hr, max_row=r1), titles_from_data=True)
        ch.set_categories(Reference(self.ws, min_col=cats[2], min_row=cats[0], max_row=cats[1]))
        for k, s in enumerate(ch.series):
            if colors and k < len(colors) and colors[k]:
                if kind == "bar":
                    s.graphicalProperties.solidFill = colors[k]
                    s.graphicalProperties.line.solidFill = colors[k]
                else:
                    s.graphicalProperties.line.solidFill = colors[k]
                    s.smooth = False
            elif kind == "line":
                s.smooth = False
        if fmt:
            ch.y_axis.number_format = fmt
        ch.y_axis.delete = False
        ch.x_axis.delete = False
        self.ws.add_chart(ch, anchor)
        return ch

    def heat(self, rng, mid=0):
        self.ws.conditional_formatting.add(rng, ColorScaleRule(start_type="min", start_color="C01E1E", mid_type="num", mid_value=mid,
                                                               mid_color="FFFFFF", end_type="max", end_color="87C081"))


class Book:
    def __init__(self, title, intro):
        self.wb = Workbook()
        self.start = self.wb.active
        self.start.title = "Start here"
        self.title, self.intro = title, intro
        self.index = []      # (lapas, skaidrė, aprašas)
        self.checks = []     # (lapas, langelis, skaidrės reikšmė, tolerancija)

    def sheet(self, name, desc, slide=None, title=None, note=None):
        assert len(name) <= 31, name
        ws = self.wb.create_sheet(name)
        s = Sheet(self, ws)
        self.index.append((name, slide, desc))
        if title:
            s.text(title, TITLE)
            if note:
                s.text(note, NOTE)
            s.gap()
        return s

    def save(self, path, check=True):
        st = self.start
        rows = [(self.title, TITLE), *[(t, None) for t in self.intro], ("", None),
                ("How to read this workbook", BOLD),
                ("Blue numbers = inputs copied from a source; the source is written on the same row (or at the top of the column).", BLUE),
                ("Black numbers = Excel formulas. Click a cell to see how it is calculated; formulas point to the 'Data' sheets.", None),
                ("Sheets 'S## …' = one sheet per slide with a table or chart (## = slide number): the same table / chart as on the slide.", None),
                ("Sheets 'Data …' = all inputs, each with its source (document and page / sheet and row, or Morningstar datapoint, link and date).", None),
                ("Source keys ([P1], [D3], [M1] …) are the same as on the slide 'Sources'; the documents are in the same ZIP, file names start with the key.", None),
                ("", None), ("Sheets", BOLD)]
        for i, (t, f) in enumerate(rows, 1):
            c = st.cell(i, 1, t)
            if f:
                c.font = f
        r = len(rows) + 1
        for j, h in enumerate(("Sheet", "Slide", "What it shows"), 1):
            x = st.cell(r, j, h); x.font, x.fill = HF, HEAD
        for name, slide, desc in self.index:
            r += 1
            c = st.cell(r, 1, name); c.hyperlink = f"#{q(name)}!A1"; c.font = Font(color="0563C1", underline="single")
            st.cell(r, 2, slide or "")
            st.cell(r, 3, desc)
        st.column_dimensions["A"].width = 34; st.column_dimensions["B"].width = 8; st.column_dimensions["C"].width = 120
        path = Path(path)
        self.wb.save(path)
        # formulių reikšmės įrašomos (kad matytųsi ir peržiūros programose) – LibreOffice perskaičiavimas
        tmp = path.parent / "_recalc"
        shutil.rmtree(tmp, ignore_errors=True); tmp.mkdir()
        subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(tmp), str(path)],
                       check=True, capture_output=True, timeout=300)
        out = tmp / path.name
        bad = []
        if check:
            wv = load_workbook(out, data_only=True)
            for sh, addr, val, tol in self.checks:
                got = wv[sh][addr].value
                if isinstance(val, str):
                    ok = str(got) == val
                else:
                    ok = isinstance(got, (int, float)) and abs(got - val) <= tol
                if not ok:
                    bad.append((sh, addr, val, got))
        # galutinis – perskaičiuota LibreOffice kopija (formulės + išsaugotos reikšmės: skaičiai matomi ir Excel „Protected View“ režime)
        shutil.move(str(out), str(path))
        shutil.rmtree(tmp, ignore_errors=True)
        return len(self.checks), bad
