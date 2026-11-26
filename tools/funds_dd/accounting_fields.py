"""Apskaitos formos „Field Name / Data“ laukai (Vytis 2026-11-11) – bendri obligacijų (docx) ir fondų (pptx) DD.

Tvarka ir pavadinimai – pažodžiui kaip apskaitos formoje (įskaitant „Securitie“), tvarkos nekeisti.
OPTIONS – leistinų reikšmių sąrašai iš formos; dokumente rodomi visi, pasirinkta paryškinta žaliai.
Exchange sąrašas formoje neišsamus: jei biržos nėra, reikšmė įrašoma ir pažymima „added – not in the list“.
Taisyklės:
- Ticker: fondas – pilnas Bloomberg tickeris („SNAW GY Equity“); įmonės obligacija – „<ISIN> Corp“; vyriausybės – „<ISIN> Gov“.
- Settlements T+: nauja emisija – nurodyti, kada atsiskaitoma / pradedama platinti (data).
- Investment Guidelines Class: Credit – įmonių obligacijos; Fixed Income – vyriausybių obligacijos.
- Kiekviename lauke su sąrašu VISADA pasirenkama reikšmė iš sąrašo (Vytis 2026-11-11), įskaitant
  Collective investment undertaking types ir obligacijoms (pvz. Fixed Income). Išimtis – tik Exchange (neišsamus sąrašas).
"""

FIELDS = [
    ("Type of issue", ["Government Agency", "Government", "Investment Fund", "Derivative Instrument", "Municipality", "Corporate",
                       "Term Deposits", "Supranational"]),
    ("Trading Venue", ["Regulated Market", "Other Market"]),
    ("Exchange", ["Luxembourg LX", "Ireland ID", "Stockholm SS", "Germany GY"]),
    ("Ticker", None),
    ("ISIN code", None),
    ("Currency", None),
    ("Securitie class", ["Real Estate", "Alternative Investment Funds", "Equity", "Fixed Income", "Investment Funds", "Derivatives",
                         "Term Deposits"]),
    ("Securitie type", ["Common Shares", "Preferred Shares", "Treasury Bills", "Bonds", "Term Deposit", "Mortgage-Backed Securities",
                        "Equity Funds", "Options", "Futures", "Fixed Income Funds", "Money Market Funds", "Liquid Alternative Funds",
                        "Real Estate Fund with unit NAV", "Real Estate - direct", "Private Equity and Venture Capital Funds",
                        "Infrastructure Funds", "Real Estate Funds without unit NAV", "Forest Funds"]),
    ("Asset class affiliation", ["Real Estate", "Alternative Investment Funds", "Equity", "Fixed Income", "Investment Funds", "Derivatives",
                                 "Term Deposits"]),
    ("Fund decimal units", None),
    ("Settlements T+", None),
    ("Investment Guidelines Class (High level classification)", ["Cash", "Fixed Income", "Credit", "Equity", "Property"]),
    ("Investment Guidelines Subclass (Low level classification)", ["Cash", "Government bonds", "Investment grade corporate bonds",
                                                                   "HY corporate bonds", "EMD HC", "EMD LC", "Loans, direct lending & Mezzanine",
                                                                   "Global Equities", "Private Equity & VC", "Property & Infrastructure"]),
    ("Collective investment undertaking types (investment funds)", ["Real Estate", "Alternative Investment Funds", "Equity", "Fixed Income",
                                                                    "Investment Funds", "Derivatives", "Term Deposits"]),
]
FUND_ONLY = {"Collective investment undertaking types (investment funds)"}
LIST_INCOMPLETE = {"Exchange"}

# Trumpi KIS tipų paaiškinimai (pagal Vyčio pateiktą Lietuvos banko standartą, apskaita_fondams.md)
KIS_SHORT = {
    "lt": [("KIS 1", "skolos VP (obligacijų) fondai"), ("KIS 2", "mišrūs (subalansuoti) fondai"), ("KIS 3", "akcijų fondai"),
           ("KIS 4", "pinigų rinkos fondai"), ("KIS 5", "nekilnojamojo turto fondai"), ("KIS 6", "rizikos draudimo (hedge) fondai"),
           ("KIS 7", "kiti KIS")],
    "en": [("KIS 1", "bond (debt securities) funds"), ("KIS 2", "mixed (balanced) funds"), ("KIS 3", "equity funds"),
           ("KIS 4", "money market funds"), ("KIS 5", "real estate funds"), ("KIS 6", "hedge funds"), ("KIS 7", "other collective investment undertakings")],
}


def rows(values):
    """values: {lauko pavadinimas: (reikšmė, pastaba)} arba {lauko pavadinimas: reikšmė}.
    Grąžina [(field, value, note, options|None, selected_in_list: bool)] – formos tvarka; trūkstamas laukas – klaida."""
    out = []
    for name, opts in FIELDS:
        if name not in values:
            raise KeyError(f"Apskaitos laukas neužpildytas: {name}")
        v = values[name]
        val, note = v if isinstance(v, tuple) else (v, "")
        sel = bool(opts) and val in opts
        if opts and not sel and name not in LIST_INCOMPLETE:
            raise ValueError(f"{name}: '{val}' nėra leistinų reikšmių sąraše – būtina pasirinkti iš sąrašo")
        out.append((name, val, note, opts, sel))
    return out


TXT = {
    "en": {"head": ["Field Name", "Data", "Allowed options (selected in green)"], "free": "free text – no list",
           "added": "not in the form list – add manually", "fund_only": "funds only",
           "lead": "Fields exactly in the order of the accounting form. The right-hand column shows the list each value was chosen from.",
           "incomplete": "The Exchange list in the form is not complete: if the exchange is missing, it is added manually."},
    "lt": {"head": ["Field Name", "Data", "Leistinos reikšmės (pasirinkta – žaliai)"], "free": "laisvas tekstas – sąrašo nėra",
           "added": "formos sąraše nėra – pridėti ranka", "fund_only": "tik fondams",
           "lead": "Laukai – tiksliai apskaitos formos tvarka. Dešiniajame stulpelyje – sąrašas, iš kurio parinkta reikšmė.",
           "incomplete": "Formos biržų (Exchange) sąrašas neišsamus: jei biržos nėra, ji pridedama ranka."},
}


def docx_table(g, values, lang="en", ws=(4.6, 5.4, 7.2)):
    """Lentelė Word DD (python-docx); g – generatoriaus globals() (d, cell_text, shade, margins, borders, widths, keep, rich, Pt, rgb, spalvos)."""
    T = TXT[lang]
    R = rows(values)
    t = g["d"].add_table(rows=len(R) + 1, cols=3); g["no_borders"](t)
    for j, h in enumerate(T["head"]):
        c = t.cell(0, j); g["cell_text"](c, h, 8.5, True, "FFFFFF"); g["shade"](c, g["NAVY"]); g["margins"](c, 50, 50, 90, 90)
    for i, (name, val, note, opts, sel) in enumerate(R, 1):
        c0, c1, c2 = t.cell(i, 0), t.cell(i, 1), t.cell(i, 2)
        g["cell_text"](c0, name, 8.5, True, g["NAVY"])
        g["shade"](c0, "DDEBF7" if name in FUND_ONLY else "DFF2DF")  # spalvos kaip apskaitos formoje
        g["cell_text"](c1, f"**{val}**" + (f"\n{note}" if note else ""), 8.5)
        if len(c1.paragraphs) > 1:
            for r in c1.paragraphs[1].runs:
                r.font.size = g["Pt"](7.5); r.font.color.rgb = g["rgb"](g["GREY"])
        c2.text = ""; p = c2.paragraphs[0]
        if opts:
            for k, o in enumerate(opts):
                r = p.add_run(("✓ " + o) if o == val else o)
                r.font.size = g["Pt"](7.5)
                if o == val:
                    r.bold = True; r.font.color.rgb = g["rgb"](g["GREEN"])
                else:
                    r.font.color.rgb = g["rgb"]("8A94A3")
                if k < len(opts) - 1:
                    s = p.add_run("  ·  "); s.font.size = g["Pt"](7.5); s.font.color.rgb = g["rgb"]("C0C7D0")
            if not sel and name in LIST_INCOMPLETE and not val.startswith(("Not applicable", "Netaikoma")):
                r = p.add_run(f"  ·  + {val} ({T['added']})"); r.font.size = g["Pt"](7.5); r.bold = True; r.font.color.rgb = g["rgb"](g["AMBER"])
        else:
            r = p.add_run(T["free"]); r.italic = True; r.font.size = g["Pt"](7.5); r.font.color.rgb = g["rgb"]("8A94A3")
        for c in (c0, c1, c2):
            g["margins"](c, 45, 45, 90, 90); g["borders"](c, bottom=(4, g["LINE"]))
            c.vertical_alignment = 1  # center
    g["widths"](t, list(ws)); g["keep"](t)
    return t


def pptx_rows(values, lang="en"):
    """Eilutės fondų DD skaidrei (funds_dd_lib.table žymėjimas: ++žalia++, ~~pilka~~, ^^ruda^^)."""
    T = TXT[lang]
    out = [T["head"]]
    for name, val, note, opts, sel in rows(values):
        data = f"**{val}**" + (f" – {note}" if note else "")
        if opts:
            parts = [("++✓ " + o + "++") if o == val else ("~~" + o + "~~") for o in opts]
            if not sel and name in LIST_INCOMPLETE and not val.startswith(("Not applicable", "Netaikoma", "NOT FOUND")):
                parts.append(f"^^+ {val} ({T['added']})^^")
            o = " · ".join(parts)
        else:
            o = "~~" + T["free"] + "~~"
        out.append([name, data, o])
    return out


def kis_lines(selected=None, lang="lt"):
    """KIS 1–7 trumpi paaiškinimai fondų DD klasifikavimo skaidrei; pasirinktas – žaliai (++…++)."""
    return [("++" if k == selected else "") + f"{k} – {t}" + ("++" if k == selected else "") for k, t in KIS_SHORT[lang]]


# Pilni KIS apibrėžimai – pažodžiui, kaip pateikė Vytis 2026-11-11 (Lietuvos banko standartas); tekstą nekeisti
KIS_FULL = [
    ("KIS 1", "KIS, kurių investavimo strategijoje numatyta iki 100 proc. grynųjų aktyvų investuoti į skolos vertybinius popierius ir (arba) į skolos vertybinius popierius investuojančių KIS išleistus vienetus (akcijas)"),
    ("KIS 2", "KIS, kurių investavimo strategijoje numatytas mišrus (subalansuotas) investavimas ir (arba) investavimas į mišraus (subalansuoto) investavimo KIS išleistus vienetus (akcijas)"),
    ("KIS 3", "KIS, kurių investavimo strategijoje numatyta iki 100 proc. grynųjų aktyvų investuoti į nuosavybės vertybinius popierius ir (arba) į nuosavybės vertybinius popierius investuojančių KIS išleistus vienetus (akcijas)"),
    ("KIS 4", "KIS, kurių investavimo strategijoje numatyta iki 100 proc. grynųjų aktyvų investuoti į pinigų rinkos priemones"),
    ("KIS 5", "KIS, kurių investavimo strategijoje numatyta iki 100 proc. grynųjų aktyvų investuoti į nekilnojamą turtą ir (arba) į nekilnojamą turtą investuojančių KIS išleistus vienetus (akcijas)"),
    ("KIS 6", "KIS, kurių investavimo stratega atitinka rizikos draudimo fondų (angl., hedge funds) taikomą strategiją ir (arba) kurie investuoja į rizikos draudimo fondų išleistus vienetus (akcijas)"),
    ("KIS 7", "kiti KIS"),
]


def kis_full_rows(selected=None):
    """Suflerio lentelė: [KIS tipas, apibrėžimas]; pasirinktas – žaliai, kiti – įprastai."""
    out = [["KIS tipas", "Apibrėžimas (Lietuvos banko standartas, pažodžiui)"]]
    for k, t in KIS_FULL:
        out.append([f"++{k}++" if k == selected else f"**{k}**", f"++{t}++" if k == selected else t])
    return out


def kis_full_lines(selected=None):
    return [(f"++{k} – {t}++" if k == selected else f"**{k}** – {t}") for k, t in KIS_FULL]
