"""Funds DD: iShares MSCI Poland UCITS ETF, USD (Acc) (IE00B4M7GH52), XETRA line IBCJ GY.
Struktūra – struktura_etf_dd.md v4 / v4.1 / v4.2: šablono skaidrės šablono tvarka (indekso fondas – „Fund overview  (index funds)“),
nenaudojamos nerodomos, skyriklis „Additional analysis“; vienas skaičiavimų Excel [C1] su lapu kiekvienai skaidrei su lentele /
diagrama ir patikra Excel = skaidrė; kiekviena [C1] žymė su lapu + langeliu; ZIP – tik PDF ir Excel.
Duomenys: oficialūs iShares / BlackRock ir MSCI dokumentai (fund-docs, 23.11.2026), Morningstar jungtis (23.11.2026),
savininko pateiktas Bloomberg tickeris (IBCJ GY Equity, 23.11.2026).
Paleisti: python3 tools/funds_dd/make_dd_ishares_poland.py  (duomenys ir rezultatas – $FUNDS_DD_WORK, numatyta /mnt/project-files/funds_dd; žr. README.md)
"""
import datetime as dt
import json
import math
import os
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from pptx.util import Inches  # noqa: E402

import calc_book as CB  # noqa: E402
import funds_dd_lib as L  # noqa: E402
import accounting_fields as AF  # noqa: E402
from funds_dd_lib import SOURCES  # noqa: E402

REV = "2026-11-23"
FUND = "iShares MSCI Poland UCITS ETF"
ISIN = "IE00B4M7GH52"
WORK = Path(os.environ.get("FUNDS_DD_WORK", "/mnt/project-files/funds_dd"))  # vietinis darbo aplankas (ne repo)
EVD = WORK / "dd_evidence" / "iShares_MSCI_Poland_IE00B4M7GH52"
DOC = EVD / "docs"
PKG = EVD / "package"; PKG.mkdir(parents=True, exist_ok=True)
OUT = WORK / f"DD_Funds_iShares_MSCI_Poland_{ISIN}_{REV}.pptx"
XLSX = PKG / f"C1_Calculations_iShares_MSCI_Poland_{REV}.xlsx"
TMPI = Path(tempfile.gettempdir()) / "poland_dd_img"; TMPI.mkdir(exist_ok=True)
MSU = "https://www.morningstar.com/integrations/mcp-redirects/security/"
OWN = "Bloomberg, 2026-11-23"
OWNI = "internal assessment, 2026-11-23"
L.FINAL = True  # galutinė versija (savininkas 2026-11-23): be raudonų žymų, NOT FOUND → pilka n/a
SPREAD, AVGVOL = 0.065, 252_900  # savininko Bloomberg duomenys 2026-11-23: „0,065“ (Avg Bid Ask Spread %), „252,9k“ (90 d. vid. apyvarta, vnt.)

# ------------------------------------------------------------------ Morningstar duomenys (paimta 23.11.2026)
_m = json.loads((EVD / "ms_monthly_2026-11-23.json").read_text())["result"]
SER = {k: {t["date"][:7]: float(t["value"]) for x in v["values"] if x["datapointId"] == "HP010" for t in x["timeSeriesData"]} for k, v in _m.items()}
SNAP = json.loads((EVD / "ms_snapshot_2026-11-23.json").read_text())["funds"]
_DN = json.loads((EVD / "ms_snapshot_2026-11-23.json").read_text())["datapoints"]
_d = json.loads((EVD / "ms_daily_2026-11-23.json").read_text())["result"]
DLY = {k: {t["date"]: float(t["value"]) for t in v["values"][0]["timeSeriesData"]} for k, v in _d.items()}
_E = json.loads((EVD / "ms_esg_ts_2026-11-23.json").read_text())
ESG = _E["data"]
HOLD = json.loads((EVD / "ms_holdings_2026-11-23.json").read_text())

SOURCES.update({
    "M1": ("Morningstar – iShares MSCI Poland UCITS ETF USD (Acc), fund NAV, SecId F00000LYBU, read 23.11.2026", MSU + "F00000LYBU/quote"),
    "M2": ("Morningstar – iShares MSCI Poland UCITS ETF, XETRA line IBCJ (EUR), SecId 0P0000TJXS, read 23.11.2026", MSU + "0P0000TJXS/quote"),
    "M3": ("Morningstar – Expat Poland WIG20 UCITS ETF (BGPLWIG04173), SecId F00001T1YT", MSU + "F00001T1YT/quote"),
    "M4": ("Morningstar – Amundi MSCI Eastern Europe ex Russia UCITS ETF Acc (LU1900066462), SecId F000013D6B", MSU + "F000013D6B/quote"),
    "M5": ("Morningstar – iShares Core MSCI EM IMI UCITS ETF USD Acc (IE00BKM4GZ66), SecId F00000TJCR", MSU + "F00000TJCR/quote"),
    "M6": ("Morningstar – iShares Core MSCI World UCITS ETF (IE00B4L5Y983), XETRA EUNL – Article 6 comparator, SecId 0P0000MEHZ", MSU + "0P0000MEHZ/quote"),
    "W1": ("iShares – product page iShares MSCI Poland UCITS ETF (key facts, sectors, securities lending, listings), read 23.11.2026",
           "https://www.ishares.com/uk/individual/en/products/251875/ishares-msci-poland-ucits-etf"),
    "J1": ("justETF – profile iShares MSCI Poland UCITS ETF USD (Acc), IE00B4M7GH52, read 23.11.2026", "https://www.justetf.com/en/etf-profile.html?isin=IE00B4M7GH52"),
    "J2": ("justETF – profile Expat Poland WIG20 UCITS ETF, BGPLWIG04173, read 23.11.2026", "https://www.justetf.com/en/etf-profile.html?isin=BGPLWIG04173"),
    "J3": ("justETF – profile Amundi MSCI Eastern Europe ex Russia UCITS ETF Acc, LU1900066462, read 23.11.2026", "https://www.justetf.com/en/etf-profile.html?isin=LU1900066462"),
    "N1": ("Central Bank of Ireland – warning notice on unauthorised firms cloning “BlackRock Asset Management Ireland Limited”, 26.09.2022, updated 20.10.2026",
           "https://www.centralbank.ie/news/article/warning-notice-central-bank-of-ireland-issues-warning-on-unauthorised-firm-blackrock-asset-management-ireland-limited-28-apr-2023"),
    "N2": ("Texas Attorney General – press release “Trump Administration Supports Attorney General Ken Paxton’s Lawsuit Against BlackRock, State Street …”, 23.05.2025",
           "https://www.texasattorneygeneral.gov/news/releases/trump-administration-supports-attorney-general-ken-paxtons-lawsuit-against-blackrock-state-street"),
    "N3": ("Climate Litigation Database – Texas v. BlackRock, Inc., No. 6:24-cv-00437 (E.D. Tex.), case page read 23.11.2026",
           "https://www.climatecasechart.com/document/texas-v-blackrock-inc_f224"),
    "X1": ("SEB eNameChecker (internal)", os.environ.get("ENAMECHECKER_URL", "")),
    "X2": ("OFAC Sanctions List Search", "https://sanctionssearch.ofac.treas.gov/"),
    "X3": ("EU Sanctions Map", "https://www.sanctionsmap.eu/#/main"),
})
IG = "https://www.ishares.com/gls-download/literature/"
URL = {"P1": IG + "prospectus/ishares-v-plc-prospectus-en.pdf",
       "K1": IG + "kiid/eu-priips-ishares-msci-poland-ucits-etf-usd-acc-ie00b4m7gh52-en.pdf",
       "F1": IG + "fact-sheet/spol-ishares-msci-poland-ucits-etf-fund-fact-sheet-en-gb.pdf",
       "A1": IG + "annual-report/ishares-v-plc-annual-report-en-30-nov-25.pdf",
       "I1": "https://www.msci.com/documents/10199/255599/msci-poland-index-usd-net.pdf"}
DOCS = {
    "P1": "iShares V plc – Prospectus, 3 September 2026 (fund section p.89; investment restrictions Schedule III pp.229–231); PDF page = printed page",
    "K1": "PRIIPs Key Information Document, iShares MSCI Poland UCITS ETF USD (Acc) (IE00B4M7GH52), 09.04.2026",
    "F1": "Factsheet, iShares MSCI Poland UCITS ETF USD (Acc), October 2026 (performance as of 31.10.2026)",
    "A1": "iShares V plc – Annual report and audited financial statements, year ended 30.11.2025; page = PDF page (printed page number is one lower)",
    "I1": "MSCI – Index factsheet MSCI Poland Index (USD), 30.10.2026",
    "C1": f"Own calculation – workbook {XLSX.name}: one sheet per slide with a table or chart (S##), all inputs with sources (Data sheets); formulas",
}
LOCAL = {k: DOC / n for k, n in (("P1", "prospectus_b9876949eb.pdf"), ("K1", "kid_7a36854ba2.pdf"), ("F1", "factsheet_92f39be76c.pdf"),
                                 ("A1", "annual_report_2025_7d4a8706fd.pdf"), ("I1", "index_factsheet_4da864c975.pdf"))}
EXPECT = {"W1": ["MSCI Poland"], "J1": ["IE00B4M7GH52"], "J2": ["BGPLWIG04173"], "J3": ["LU1900066462"], "N1": ["BlackRock Asset Management Ireland"],
          "N2": ["BlackRock"], "N3": ["BlackRock"]}

pc = lambda v, d=1: f"{v * 100:+.{d}f}%"  # noqa: E731
p0 = lambda v, d=1: f"{v * 100:.{d}f}%"  # noqa: E731
pp = lambda v, d=1: f"{v * 100:+.{d}f} pp"  # noqa: E731
GREENC, DARK, AMBER, REDC = "87C081", "191824", "B45F06", "C01E1E"
fm = lambda x: x[8:10] + "." + x[5:7] + "." + x[:4]  # noqa: E731


def comp(s, ks):
    return math.prod(1 + s[k] / 100 for k in ks) - 1


def ann(s, ks):
    return math.prod(1 + s[k] / 100 for k in ks) ** (12 / len(ks)) - 1


def vol(s, ks):
    x = [s[k] / 100 for k in ks]
    m = sum(x) / len(x)
    return math.sqrt(sum((v - m) ** 2 for v in x) / (len(x) - 1)) * math.sqrt(12)


def months(m0, m1):
    y, m, out = int(m0[:4]), int(m0[5:]), []
    while f"{y}-{m:02d}" <= m1:
        out.append(f"{y}-{m:02d}"); m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


# ------------------------------------------------------------------ serijos
FUNDM, XET, EXPAT, CEC, EM, WLD = (SER[k] for k in ("F00000LYBU", "0P0000TJXS", "F00001T1YT", "F000013D6B", "F00000TJCR", "0P0000MEHZ"))
LM = "2026-10"
FK = months("2011-02", LM)                      # fondas nuo pirmo pilno mėnesio
EK = months("2014-07", LM)                      # EIMI nuo pirmo pilno mėnesio
assert all(k in FUNDM and k in WLD for k in FK) and all(k in EM for k in EK)
K10, K5, K3, K1Y = months("2016-11", LM), months("2021-11", LM), months("2023-11", LM), months("2025-11", LM)


def index_of(s, ks, start):
    v, out = 1.0, {start: 1.0}
    for k in ks:
        v *= 1 + s[k] / 100; out[k] = v
    return out


VF, VW = index_of(FUNDM, FK, "2011-01"), index_of(WLD, FK, "2011-01")
VFE, VEM = index_of(FUNDM, EK, "2014-06"), index_of(EM, EK, "2014-06")


def falls(v, ks):
    pk, out = None, {}
    for k in ks:
        pk = v[k] if pk is None else max(pk, v[k])
        out[k] = v[k] / pk - 1
    return out


FF, FW = falls(VF, ["2011-01"] + FK), falls(VW, ["2011-01"] + FK)
MDD_F, MDD_W = min(FF.values()), min(FW.values())
MDD_F_AT = min(FF, key=FF.get)
FFE, FEM = falls(VFE, ["2014-06"] + EK), falls(VEM, ["2014-06"] + EK)
MDD_FE, MDD_EM = min(FFE.values()), min(FEM.values())


def episodes(v, ks, thr):
    """Kritimai nuo piko iki dugno ≥ thr: [(pikas, dugnas, kritimas)]."""
    out, pk, cur = [], ks[0], None
    for k in ks + [None]:
        if k is None or v[k] >= v[pk]:
            if cur and cur[2] <= -thr:
                out.append(cur)
            if k is None:
                break
            pk, cur = k, None
        else:
            f = v[k] / v[pk] - 1
            if cur is None or f < cur[2]:
                cur = (pk, k, f)
    return out


EPM = episodes(VF, ["2011-01"] + FK, 0.15)     # fondo (MSCI Poland) kritimai ≥ 15 % mėnesio pabaigomis
WORST = sorted(EK, key=lambda k: FUNDM[k])[:6]


def capture(f, b, ks, down=True):
    ix = [k for k in ks if (b[k] < 0) == down]
    g = lambda s: math.prod(1 + s[k] / 100 for k in ix) ** (1 / len(ix)) - 1  # noqa: E731
    return g(f) / g(b), len(ix)


CAPD, CAPU = capture(FUNDM, EM, EK), capture(FUNDM, EM, EK, False)
CORR = (lambda a, b: sum((x - sum(a) / len(a)) * (y - sum(b) / len(b)) for x, y in zip(a, b)) /
        math.sqrt(sum((x - sum(a) / len(a)) ** 2 for x in a) * sum((y - sum(b) / len(b)) ** 2 for y in b)))([FUNDM[k] for k in EK], [EM[k] for k in EK])
# dienos (darbo dienos)
DK = [x for x in DLY["F00000LYBU"] if dt.date.fromisoformat(x).weekday() < 5]
DF, DX, DCEC, DEM, DW = (DLY[k] for k in ("F00000LYBU", "0P0000TJXS", "F000013D6B", "F00000TJCR", "0P0000MEHZ"))
EPD = episodes(DF, DK, 0.08)
FD = {n: falls(s, DK) for n, s in (("fund", DF), ("em", DEM), ("cec", DCEC), ("w", DW))}
MDDD = {n: min(f.values()) for n, f in FD.items()}
win = lambda s, a, z: s[z] / s[a] - 1  # noqa: E731
NOV20 = DF["2026-11-20"] / DF["2026-11-19"] - 1

# ------------------------------------------------------------------ dokumentų skaičiai
TDY = [2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025]
FS_CY = {2016: (0.02, 0.13), 2017: (54.33, 54.72), 2018: (-13.14, -12.87), 2019: (-6.03, -5.87), 2020: (-11.91, -11.39), 2021: (8.16, 8.46),
         2022: (-27.36, -27.24), 2023: (48.25, 48.60), 2024: (-6.47, -6.65), 2025: (74.88, 74.61)}
TD = {y: (a - b) / 100 for y, (a, b) in FS_CY.items()}
TD_AVG = sum(TD.values()) / len(TD)
TD_POS = sum(1 for v in TD.values() if v > 0)
IDX_CY = {2012: (39.24, 18.22, 16.38), 2013: (2.59, -2.60, 23.55), 2014: (-14.21, -2.19, 3.84), 2015: (-25.35, -14.92, -2.19), 2016: (0.13, 11.19, 8.36),
          2017: (54.72, 37.28, 23.95), 2018: (-12.87, -14.57, -10.08), 2019: (-5.87, 18.42, 26.35), 2020: (-11.39, 18.31, 16.25), 2021: (8.46, -2.54, 18.22),
          2022: (-27.24, -20.09, -18.40), 2023: (48.60, 9.83, 21.58), 2024: (-6.65, 7.50, 16.37), 2025: (74.61, 33.57, 22.06)}
IDX_UP = sum(1 for v in IDX_CY.values() if v[0] > v[1])
SL = [("2016/17", 0.17, 33.19, 57.50, 112.19), ("2017/18", 0.03, 10.78, 20.02, 110.12), ("2018/19", 0.01, 7.11, 22.79, 112.10), ("2019/20", 0.02, 7.67, 20.13, 111.81),
      ("2020/21", 0.05, 10.41, 21.47, 110.24), ("2021/22", 0.05, 18.61, 30.95, 108.69), ("2022/23", 0.08, 24.23, 37.86, 109.64), ("2023/24", 0.05, 21.39, 35.96, 109.34),
      ("2024/25", 0.05, 21.93, 38.22, 108.46), ("2025/26", 0.04, 10.74, 18.94, 109.06)]
SECT = [("Financials", 46.99, 46.22), ("Consumer Discretionary", 17.38, 15.65), ("Energy", 13.31, 15.48), ("Materials", 6.52, 7.18), ("Consumer Staples", 5.67, 5.71),
        ("Communication", 3.22, 3.05), ("Utilities", 2.66, 2.67), ("Information Technology", 2.17, 2.40), ("Industrials", 1.73, 1.64), ("Cash and/or derivatives", 0.34, None)]
TOP = [h for h in HOLD["holdings"] if h[1] > 1][:16]
TOP3 = sum(h[1] for h in TOP[:3]) / 100
TOP10 = sum(h[1] for h in TOP[:10]) / 100

# ------------------------------------------------------------------ skaidrių tvarka
SANC = [("BlackRock Asset Management Ireland Limited", "management company (the “Manager”), regulated by the Central Bank of Ireland [K1 p.1]"),
        ("BlackRock Advisors (UK) Limited", "investment manager – the Manager delegated investment management to it [P1 p.199][P1 p.12]")]
ORDER = (["title", "conclusion", "overview", "index", "esg1", "esg2", "perf", "team", "strategy", "portfolio", "other", "divider",
          "decision", "questions", "fees", "track", "index_hist", "esg_ms", "esg_trend"] + L.sanction_keys(SANC) +
         ["media", "perf_heat", "perf_risk", "perf_dd_live", "perf_dd_hist", "perf_check", "extracts",
          "alt1", "alt2", "alt_heat", "alt3", "acc_law", "acc_law2", "acc_evid", "acc_bbg", "acc_form", "acc_class", "method", "sources"])
NO = {k: i + 1 for i, k in enumerate(ORDER)}


def sl(*keys):
    nums = [NO[k] for k in keys]
    return ("Slide: " if len(nums) == 1 else "Slides: ") + "; ".join(str(n) for n in nums)


def SN(key, name):
    return f"S{NO[key]:02d} {name}"


# ======================================================================= EXCEL [C1]
BK = CB.Book(f"DD {FUND} USD (Acc) ({ISIN}) – calculations [C1], {REV}", [
    "Prepared by: Analyst. Every table and chart of the DD presentation with numbers has its own sheet here (S## = slide number).",
    "Morningstar data (monthly and daily returns in EUR, ESG, costs, fund sizes) is in the 'Data Morningstar …' sheets with SecId, datapoint code, link and retrieval date (23.11.2026).",
    "iShares, BlackRock and MSCI document figures are in 'Data documents' with document key and page (P1 prospectus, K1 KID, F1 factsheet, A1 annual report, I1 MSCI index factsheet, W1 product page).",
    "Fund vs index (tracking difference) uses the factsheet figures in USD, the fund's base currency. All own return statistics are in EUR (Morningstar, fund NAV converted to EUR).",
    "Emerging markets are represented by the iShares Core MSCI EM IMI UCITS ETF (EUR) and world equities by the iShares Core MSCI World UCITS ETF (XETRA, EUR) – investable proxies."])
P, P2 = "0.0%", "0.00%"


def ms_src(i, dp):
    return f"Morningstar {dp}, SecId {i}, retrieved 23.11.2026 – {MSU}{i}/quote"


# ---------------------------------------------------------------- Data documents
dd_ = BK.sheet("Data documents", "Figures copied from iShares / BlackRock / MSCI documents, each with document key and page",
               title="Documents – figures copied as published", note="Blue = input copied from the document named in column E (key and page). Percentages stored as decimals (0.74% = 0.0074).")
dd_.header(["Item", "Fund / value", "Benchmark / second value", "Unit", "Source"])
DR = {}


def dput(key, label, a, b=None, unit="%", src=""):
    conv = (lambda v: v / 100 if isinstance(v, (int, float)) else v) if unit == "%" else (lambda v: v)
    DR[key] = dd_.row([label, conv(a), conv(b) if b is not None else None, unit, src], fmt=[None, P2 if unit == "%" else "#,##0.00", P2 if unit == "%" else "#,##0.00"], inputs=(1, 2))
    return DR[key]


for y in TDY:
    dput(f"cy_{y}", f"Factsheet calendar-year return, share class vs benchmark (MSCI Poland NR USD), USD: {y}", *FS_CY[y], src="[F1 p.2] Calendar year performance")
for k_, lb, a, b in (("fs_1y", "1 year", 37.17, 37.10), ("fs_3y", "3 years, annualised", 34.47, 34.29), ("fs_5y", "5 years, annualised", 15.63, 15.63),
                     ("fs_si", "Since inception, annualised", 3.24, 3.31), ("fs_ytd", "YTD 2026", 25.44, 25.23)):
    dput(k_, f"Factsheet performance to 31.10.2026, share class vs benchmark, USD: {lb}", a, b, src="[F1 p.2] Cumulative & annualised performance")
dput("fs_ter", "Total expense ratio (factsheet) / KID management and other costs", 0.74, 0.74, src="[F1 p.1]; [K1 p.3]")
dput("kid_tc", "KID transaction costs a year / entry and exit fee", 0.02, 0.0, src="[K1 p.3] Composition of costs")
dput("fs_nav", "Net assets of fund, USD m, 31.10.2026 (factsheet) / 20.11.2026 (product page)", 1120.83, 1028.420282, "USD m", "[F1 p.1]; [W1] Key facts")
dput("ar_nav", "Net asset value, USD thousand: 30.11.2025 / 30.11.2024", 483714, 87734, "USD k", "[A1 p.253]")
dput("ar_nav23", "Net asset value, USD thousand: 30.11.2023", 119255, None, "USD k", "[A1 p.253]")
dput("ar_td", "Annual report FY to 30.11.2025: fund return / benchmark return, USD", 60.52, 60.33, src="[A1 p.18] Performance summary, tracking difference and tracking error")
dput("ar_td2", "Annual report FY 2025: TER / tracking difference gross of TER", 0.74, 0.93, src="[A1 p.18]")
dput("ar_te", "Annual report FY 2025: realised tracking error / anticipated tracking error (up to)", 0.41, 0.35, src="[A1 p.18]; anticipated also [P1 p.123]")
dput("sl_lim", "Securities lending: expected / maximum % of NAV (prospectus)", 39, 43, src="[P1 p.125] (expected range 0–39%)")
dput("sl_ar", "Securities on loan 30.11.2025: % of lendable assets / % of NAV", 10.76, 10.62, src="[A1 p.731]")
dput("sl_split", "Securities lending income: share kept by the fund / by BlackRock", 62.5, 37.5, src="[W1] Securities lending")
SLR = {}
for lb, r_, av, mx, co in SL:
    SLR[lb] = dd_.row([f"Securities lending year to 30 Sep {lb}: return % / average on-loan % of AUM / maximum on-loan % / collateralisation % of loan", r_ / 100, av / 100, mx / 100, co / 100,
                       "[W1] Securities lending table"], fmt=[None, P2, P2, P2, P2], inputs=(1, 2, 3, 4))
for nm, a, b in SECT:
    dput("sec_" + nm[:6], f"Sector weight: {nm}", a, b, src="Fund: [W1] Exposure breakdowns, 19.11.2026; index: [I1 p.2] Sector weights, 30.10.2026")
for k_, lb, a, b in (("pe", "Price / earnings (fund factsheet / index trailing)", 18.49, 14.49), ("pb", "Price / book", 2.26, 2.16), ("n", "Number of holdings / constituents", 16, 16),
                     ("top10", "Top 10 weight, fund 31.10.2026 / index 30.10.2026", 85.06, 85.39), ("beta", "3-year beta (factsheet) / 3-year SD % (product page)", 0.99, 17.90)):
    dput("ch_" + k_, f"Characteristic: {lb}", a, b, "x" if k_ in ("pe", "pb", "n", "beta") else "%", "[F1 p.1][F1 p.3]; [I1 p.1][I1 p.2]; [W1]")
for y, v in IDX_CY.items():
    DR[f"ix_{y}"] = dd_.row([f"MSCI index factsheet, annual net return USD: MSCI Poland / MSCI Emerging Markets / MSCI ACWI IMI, {y}", v[0] / 100, v[1] / 100, v[2] / 100,
                             "[I1 p.1] Annual performance (%)"], fmt=[None, P2, P2, P2], inputs=(1, 2, 3))
for k_, lb, a, b, c_ in (("sd3", "Annualised standard deviation 3 years", 18.02, 17.88, 12.02), ("sd5", "Annualised standard deviation 5 years", 28.45, 18.79, 14.99),
                         ("sd10", "Annualised standard deviation 10 years", 28.23, 17.68, 14.96), ("mdd", "Maximum drawdown since 29.12.2000", 77.59, 65.25, 58.59),
                         ("r10", "10 years annualised net return", 10.41, 8.15, 12.10), ("rsi", "Since 29.12.2000 annualised net return", 5.91, 8.73, 7.54)):
    DR["ir_" + k_] = dd_.row([f"MSCI index risk / return to 30.10.2026, USD: {lb} – Poland / EM / ACWI IMI", a / 100, b / 100, c_ / 100, "[I1 p.1] Index risk and return characteristics"],
                             fmt=[None, P2, P2, P2], inputs=(1, 2, 3))
dput("kid_sc", "KID scenario after 5 years, USD 10,000: stress / unfavourable", 1400, 4050, "USD", "[K1 p.2] Performance scenarios")
dput("kid_sc2", "KID scenario after 5 years, USD 10,000: moderate / favourable", 10040, 24150, "USD", "[K1 p.2] Performance scenarios")
dd_.widths(78, 16, 16, 9, 70)
DDn = CB.q(dd_.ws.title)
DB = lambda k: f"{DDn}!$B${DR[k]}"  # noqa: E731
DC = lambda k: f"{DDn}!$C${DR[k]}"  # noqa: E731
DD4 = lambda k: f"{DDn}!$D${DR[k]}"  # noqa: E731

# ---------------------------------------------------------------- Data Morningstar monthly
MSK = [("Fund NAV (this fund)", "F00000LYBU", FUNDM, "M1"), ("XETRA IBCJ (this fund)", "0P0000TJXS", XET, "M2"), ("Expat Poland WIG20", "F00001T1YT", EXPAT, "M3"),
       ("Amundi Eastern Europe ex Russia", "F000013D6B", CEC, "M4"), ("iShares Core MSCI EM IMI", "F00000TJCR", EM, "M5"), ("iShares Core MSCI World", "0P0000MEHZ", WLD, "M6")]
dm_ = BK.sheet("Data Morningstar monthly", "Morningstar monthly total returns (HP010), %, EUR, 2011-01 – 2026-10",
               title="Morningstar monthly total return, % (datapoint HP010), in EUR",
               note="Morningstar connector, retrieved 23.11.2026. Rows 4–6: SecId, source key, link. Empty = no value. Amundi: used from 04/2019 only (share class launched 14.03.2019 [M4]).")
dm_.header(["Month"] + [n for n, *_ in MSK])
dm_.row(["SecId"] + [i for _, i, _, _ in MSK], inputs=())
dm_.row(["Source key"] + [f"[{m}]" for *_, m in MSK], inputs=())
dm_.row(["Link"] + [f"{MSU}{i}/quote" for _, i, _, _ in MSK], inputs=())
MALL = months("2011-01", LM)
MM1 = dm_.r
for m in MALL:
    dm_.row([m] + [(s.get(m) if not (n.startswith("Amundi") and m < "2019-04") else None) for n, _, s, _ in MSK], fmt="0.00000")
MMr = {m: MM1 + j for j, m in enumerate(MALL)}
MMc = {n: CB.CL(j + 2) for j, (n, *_) in enumerate(MSK)}
dm_.widths(10, *(16,) * len(MSK))
DMn = CB.q(dm_.ws.title)
MMref = lambda n, m: f"{DMn}!${MMc[n]}${MMr[m]}"  # noqa: E731
MMrng = lambda n, m0, m1: f"{DMn}!${MMc[n]}${MMr[m0]}:${MMc[n]}${MMr[m1]}"  # noqa: E731
FCOMP = lambda n, m0, m1: f"=EXP(SUMPRODUCT(LN(1+{MMrng(n, m0, m1)}/100)))-1"  # noqa: E731
FANN = lambda n, m0, m1: f"=EXP(SUMPRODUCT(LN(1+{MMrng(n, m0, m1)}/100))*12/COUNT({MMrng(n, m0, m1)}))-1"  # noqa: E731
FVOL = lambda n, m0, m1: f"=STDEV({MMrng(n, m0, m1)})/100*SQRT(12)"  # noqa: E731
NF, NX, NEX, NCEC, NEM, NW = (n for n, *_ in MSK)

# ---------------------------------------------------------------- Calc: fund vs world, monthly, since 2011
cf = BK.sheet("Calc fund monthly", "Fund NAV in EUR vs iShares Core MSCI World ETF and Core MSCI EM IMI ETF: indices and falls from high, monthly 02/2011–10/2026",
              title="iShares MSCI Poland UCITS ETF (NAV, EUR) vs world and emerging-market equity ETFs, monthly",
              note="Columns B–D link to 'Data Morningstar monthly'; all other columns are formulas. Index start 31.01.2011 = 1 (EM index from 30.06.2014 = 1).")
cf.header(["Month", "Fund, %", "MSCI World ETF, %", "EM IMI ETF, %", "Fund index", "World index", "Fund fall from high", "World fall from high",
           "Fund − EM, pp", "Fund index from 06/2014", "EM index", "Fund fall from high (from 06/2014)", "EM fall from high"])
CF0 = cf.row(["2011-01 (start)", None, None, None, 1, 1, 0, 0], inputs=())
CFr = {}
for k in FK:
    i = cf.r; CFr[k] = i
    em_ok = k in EM and k >= "2014-07"
    first = k == "2014-07"
    cf.row([k, f"={MMref(NF, k)}", f"={MMref(NW, k)}", f"={MMref(NEM, k)}" if em_ok else None, f"=E{i-1}*(1+B{i}/100)", f"=F{i-1}*(1+C{i}/100)",
            f"=E{i}/MAX(E${CF0}:E{i})-1", f"=F{i}/MAX(F${CF0}:F{i})-1",
            f"=B{i}-D{i}" if em_ok else None,
            (f"=1*(1+B{i}/100)" if first else f"=J{i-1}*(1+B{i}/100)") if em_ok else None,
            (f"=1*(1+D{i}/100)" if first else f"=K{i-1}*(1+D{i}/100)") if em_ok else None,
            (f"=J{i}/MAX(1,MAX(J${CFr['2014-07'] if not first else i}:J{i}))-1") if em_ok else None,
            (f"=K{i}/MAX(1,MAX(K${CFr['2014-07'] if not first else i}:K{i}))-1") if em_ok else None],
           fmt=[None, "0.00", "0.00", "0.00", "0.0000", "0.0000", P2, P2, "0.00", "0.0000", "0.0000", P2, P2],
           expect={6: (FF[k], 1e-9), 7: (FW[k], 1e-9)} if k in (MDD_F_AT, LM) else None)
CFN = cf.r - 1
cf.widths(15, 10, 14, 12, 11, 11, 14, 14, 12, 14, 11, 16, 14)
CFn = CB.q(cf.ws.title)
CFR = lambda col, a=FK[0], z=LM: f"{CFn}!${col}${CFr[a]}:${col}${CFr[z]}"  # noqa: E731

# ---------------------------------------------------------------- Data Morningstar daily
dy = BK.sheet("Data Morningstar daily", "Morningstar daily return index (HS793), EUR, weekdays 01.11.2024–21.11.2026; falls from high",
              title="Morningstar Daily Return Index (datapoint HS793), EUR, weekdays 01.11.2024 – 20.11.2026",
              note="Morningstar connector, retrieved 23.11.2026 [M1][M2][M4][M5][M6]; weekend values (Morningstar repeats Friday) left out. Columns G–J are formulas.")
DCOLS = [("Fund NAV (F00000LYBU)", DF), ("XETRA IBCJ (0P0000TJXS)", DX), ("Amundi E. Europe ex Russia (F000013D6B)", DCEC), ("Core MSCI EM IMI (F00000TJCR)", DEM),
         ("Core MSCI World (0P0000MEHZ)", DW)]
dy.header(["Date"] + [n for n, _ in DCOLS] + ["Fund fall from high", "EM IMI fall from high", "Amundi fall from high", "World fall from high"])
DY1 = dy.r
DYr = {}
for x in DK:
    i = dy.r; DYr[x] = i
    dy.row([x] + [s.get(x) for _, s in DCOLS] + [f"=B{i}/MAX(B${DY1}:B{i})-1", f"=E{i}/MAX(E${DY1}:E{i})-1", f"=D{i}/MAX(D${DY1}:D{i})-1", f"=F{i}/MAX(F${DY1}:F{i})-1"],
           fmt=[None] + ["0.00000"] * 5 + [P2] * 4)
DYN = dy.r - 1
dy.widths(11, *(17,) * 5, 13, 13, 13, 13)
DYn = CB.q(dy.ws.title)
DYC = {"fund": "B", "xet": "C", "cec": "D", "em": "E", "w": "F"}
DYref = lambda c, x: f"{DYn}!${DYC[c]}${DYr[x]}"  # noqa: E731
assert all(x in DX and x in DCEC and x in DEM and x in DW for x in DK)

# ---------------------------------------------------------------- Data Morningstar snapshot
SNK = [("F00000LYBU", "iShares MSCI Poland (fund)", "M1"), ("0P0000TJXS", "iShares MSCI Poland, XETRA IBCJ", "M2"), ("F00001T1YT", "Expat Poland WIG20", "M3"),
       ("F000013D6B", "Amundi MSCI E. Europe ex Russia", "M4"), ("F00000TJCR", "iShares Core MSCI EM IMI", "M5"), ("0P0000MEHZ", "iShares Core MSCI World", "M6")]
DPS = ["UB422", "OS003", "OF009", "OF010", "OS00F", "EU001", "OF00L", "HS000", "HS008", "RR141", "RR158", "RR152", "RR159", "RR153", "ESG42", "ESG1W"]
sn = BK.sheet("Data Morningstar snapshot", "Morningstar snapshot datapoints (costs, size, SFDR, benchmark, portfolio date, tracking error, capture, ESG risk)",
              title="Morningstar snapshot, retrieved 23.11.2026 (fund size in EUR; risk statistics month-end 31.10.2026)",
              note="Row 4: datapoint code; row 5: datapoint name. Empty = not returned by Morningstar. Link: " + MSU + "<SecId>/quote")
sn.header(["SecId", "Fund", "Key"] + DPS)
sn.row(["", "datapoint name", ""] + [_DN[d] for d in DPS], inputs=())
SNr = {}
for i_, n_, m_ in SNK:
    SNr[i_] = sn.row([i_, n_, f"[{m_}]"] + [SNAP.get(i_, {}).get(d) for d in DPS], fmt=None)
sn.widths(12, 34, 6, *(11,) * len(DPS))
SNn = CB.q(sn.ws.title)
SNref = lambda i_, d: f"{SNn}!${CB.CL(4 + DPS.index(d))}${SNr[i_]}"  # noqa: E731

# ---------------------------------------------------------------- Data Morningstar ESG
EDATES = ["2023-12", "2024-06", "2024-12", "2025-06", "2025-12", "2026-06", "2026-09"]
EFUNDS = [("0P0000TJXS", "iShares MSCI Poland (this fund)", "M2"), ("F000013D6B", "Amundi E. Europe ex Russia", "M4"), ("F00000TJCR", "Core MSCI EM IMI", "M5"),
          ("0P0000MEHZ", "Core MSCI World (Art. 6)", "M6")]
EDP = [("ESG2H", "Carbon intensity, Scope 1+2, t CO2e per USD m revenue"), ("ESG2A", "Historical carbon risk score"), ("ESG62", "Thermal coal, % involved"),
       ("ESG63", "Tobacco, % involved"), ("ESG55", "Gambling, % involved"), ("ESG53", "Controversial weapons, % involved")]
ESV = lambda dp, i_, q: (float(ESG[i_][dp][q]) if ESG.get(i_, {}).get(dp, {}).get(q) is not None else None)  # noqa: E731
es = BK.sheet("Data Morningstar ESG", "Morningstar ESG time series (month-end values), 12/2023–09/2026",
              title="Morningstar ESG indicators, month-end values (portfolio-based), retrieved 23.11.2026",
              note="One block per datapoint; rows = funds, columns = month-end dates. Empty = no value (Amundi: swap-based, only carbon values returned). "
                   "This fund's carbon intensity jumps without a change of index (01/2024 95.9, 01/2025 51.4, 02/2026 110.4) – low reliability.")
ESr = {}
for dp, nm in EDP:
    es.text(f"{dp} – {nm}", CB.BOLD)
    es.header(["Fund", "Key"] + EDATES + ["Source"])
    for i_, n_, m_ in EFUNDS:
        ESr[(dp, i_)] = es.row([n_, f"[{m_}]"] + [ESV(dp, i_, q) for q in EDATES] + [ms_src(i_, dp)], fmt=[None, None] + ["0.000"] * len(EDATES))
    es.gap()
es.text("Single month-end values of this fund's carbon intensity (ESG2H) that jump away from the neighbouring months", CB.BOLD)
for k_, v in _E["fund_jumps"].items():
    es.row([k_, v, ms_src("0P0000TJXS", "ESG2H")], fmt=[None, "0.000"])
es.widths(30, 6, *(10,) * len(EDATES), 80)
ESn = CB.q(es.ws.title)
ESref = lambda dp, i_, q: f"{ESn}!${CB.CL(3 + EDATES.index(q))}${ESr[(dp, i_)]}"  # noqa: E731

# ================================================================= S sheets
T = {}

# ---- Performance (template)
s7 = BK.sheet(SN("perf", "Performance"), "Tracking difference by calendar year (USD, factsheet); EUR return statistics vs EM and World ETFs", slide=NO["perf"],
              title=f"Slide {NO['perf']} – Performance: fund vs MSCI Poland (USD) and own statistics in EUR",
              note="Rows 4–13 link to 'Data documents' (factsheet [F1 p.2]); difference = fund − benchmark. Statistics below: formulas on 'Data Morningstar monthly'.")
s7.header(["Year", "Fund, USD", "MSCI Poland NR USD", "Difference (tracking difference), pp"])
Y7 = {}
for y in TDY:
    Y7[y] = s7.row([str(y), f"={DB(f'cy_{y}')}", f"={DC(f'cy_{y}')}", f"=B{s7.r}-C{s7.r}"], fmt=[None, P2, P2, P2], expect={3: (TD[y], 1e-9)})
r_tda = s7.row(["Average 2016–2025", None, None, f"=AVERAGE(D{Y7[2016]}:D{Y7[2025]})"], fmt=[None, None, None, "0.000%"], expect={3: (TD_AVG, 1e-9)})
r_tdp = s7.row(["Years fund ahead of index", None, None, f"=COUNTIF(D{Y7[2016]}:D{Y7[2025]},\">0\")"], fmt=[None, None, None, "0"], expect={3: (TD_POS, 0)})
s7.chart("bar", (Y7[2016], Y7[2025], 1), [(Y7[2016] - 1, Y7[2016], Y7[2025], 4)], "F2", "Fund minus MSCI Poland, pp, each calendar year (USD)", colors=["191824"], fmt="0.00%")
s7.gap()
s7.header(["EUR, to 10/2026, a year", "This fund", "Core MSCI EM IMI ETF", "Core MSCI World ETF"])
ST7 = {}
for lb, ks in (("Since 02/2011 (fund launch)", FK), ("Since 07/2014 (EM ETF)", EK), ("10 years", K10), ("5 years", K5), ("3 years", K3), ("1 year", K1Y)):
    em_ok = all(k in EM for k in ks) and ks[0] >= "2014-07"
    ST7[lb] = s7.row([lb, FANN(NF, ks[0], LM), FANN(NEM, ks[0], LM) if em_ok else "n/a", FANN(NW, ks[0], LM)], fmt=[None, P2, P2, P2],
                     expect={1: (ann(FUNDM, ks), 1e-9), 3: (ann(WLD, ks), 1e-9), **({2: (ann(EM, ks), 1e-9)} if em_ok else {})})
for lb, ks in (("Volatility since 02/2011 (monthly, a year)", FK), ("Volatility 5 years", K5)):
    em_ok = ks[0] >= "2014-07"
    ST7[lb] = s7.row([lb, FVOL(NF, ks[0], LM), FVOL(NEM, ks[0], LM) if em_ok else "n/a", FVOL(NW, ks[0], LM)], fmt=[None, P2, P2, P2],
                     expect={1: (vol(FUNDM, ks), 1e-9), 3: (vol(WLD, ks), 1e-9)})
r_mdd7 = s7.row(["Largest fall from high, month-end, since 02/2011", f"=MIN({CFR('G')})", None, f"=MIN({CFR('H')})"], fmt=[None, P2, None, P2],
                expect={1: (MDD_F, 1e-9), 3: (MDD_W, 1e-9)})
s7.widths(46, 14, 20, 20)
T["tda"] = s7.tag(r_tda, 4); T["tdp"] = s7.tag(r_tdp, 4); T["td2020"] = s7.tag(Y7[2020], 4); T["td2025"] = s7.tag(Y7[2025], 4)
T["si"] = s7.tag(ST7["Since 02/2011 (fund launch)"], 2); T["siw"] = s7.tag(ST7["Since 02/2011 (fund launch)"], 4)
T["em14"] = s7.tag(ST7["Since 07/2014 (EM ETF)"], 2); T["em14e"] = s7.tag(ST7["Since 07/2014 (EM ETF)"], 3)
T["y5"] = s7.tag(ST7["5 years"], 2); T["y5e"] = s7.tag(ST7["5 years"], 3); T["y3"] = s7.tag(ST7["3 years"], 2); T["y10"] = s7.tag(ST7["10 years"], 2)
T["vol"] = s7.tag(ST7["Volatility since 02/2011 (monthly, a year)"], 2); T["volw"] = s7.tag(ST7["Volatility since 02/2011 (monthly, a year)"], 4)
T["vol5"] = s7.tag(ST7["Volatility 5 years"], 2); T["vol5e"] = s7.tag(ST7["Volatility 5 years"], 3)
T["mdd"] = s7.tag(r_mdd7, 2); T["mddw"] = s7.tag(r_mdd7, 4)

# ---- Portfolio
s10 = BK.sheet(SN("portfolio", "Portfolio"), "Sector weights fund vs index; holdings (Morningstar 20.11.2026)", slide=NO["portfolio"],
               title=f"Slide {NO['portfolio']} – Current portfolio characteristics", note="Sector rows link to 'Data documents'; holdings typed from the Morningstar holdings tool (portfolio 20.11.2026) [M2].")
s10.header(["Sector", "Fund 19.11.2026 [W1]", "MSCI Poland 30.10.2026 [I1 p.2]", "Difference, pp"])
S10r = {}
for nm, a, b in SECT:
    S10r[nm] = s10.row([nm, f"={DB('sec_' + nm[:6])}", f"={DC('sec_' + nm[:6])}" if b is not None else None, f"=B{s10.r}-C{s10.r}" if b is not None else None], fmt=[None, P2, P2, P2])
s10.chart("bar", (S10r[SECT[0][0]], S10r[SECT[-2][0]], 1), [(S10r[SECT[0][0]] - 1, S10r[SECT[0][0]], S10r[SECT[-2][0]], 2), (S10r[SECT[0][0]] - 1, S10r[SECT[0][0]], S10r[SECT[-2][0]], 3)],
          "F2", "Sector weights, fund vs MSCI Poland", colors=["191824", "AFAFAF"], fmt="0%")
s10.gap()
s10.header(["Holding (Morningstar, 20.11.2026)", "Weight", "Cumulative", "Source"])
H10 = {}
h0 = s10.r
for j, (nm, w) in enumerate(TOP):
    i = s10.r
    H10[j] = s10.row([nm, w / 100, f"=SUM(B${h0}:B{i})", f"[M2] {HOLD['url']}"], fmt=[None, P2, P2], inputs=(1,))
s10.widths(40, 18, 22, 14)
T["top3"] = s10.tag(H10[2], 3, TOP3); T["top10"] = s10.tag(H10[9], 3, TOP10); T["pko"] = s10.tag(H10[0], 2, TOP[0][1] / 100)

# ---- Fees
FEEF = [("This fund", "0P0000TJXS"), ("Expat Poland WIG20", "F00001T1YT"), ("Amundi MSCI E. Europe ex Russia", "F000013D6B"), ("iShares Core MSCI EM IMI", "F00000TJCR")]
s15 = BK.sheet(SN("fees", "Fees"), "Cost comparison: ongoing charge / management fee and difference to this fund", slide=NO["fees"],
               title=f"Slide {NO['fees']} – Fund overview: fees", note="Ongoing charge from Morningstar UB422 (KIID ongoing charge); Expat: management fee OS003 (no UB422 returned). This fund equals KID [K1 p.3].")
s15.header(["Fund", "Ongoing charge (or management fee)", "Difference to this fund, pp a year"])
F15 = {}
for n_, i_ in FEEF:
    dp = "UB422" if SNAP[i_].get("UB422") is not None else "OS003"
    F15[i_] = s15.row([n_ + ("" if dp == "UB422" else " (management fee)"), f"={SNref(i_, dp)}/100", f"=B{s15.r}-{SNref('0P0000TJXS', 'UB422')}/100"], fmt=[None, P2, P2],
                      expect={1: (SNAP[i_][dp] / 100, 1e-9)})
s15.gap()
r_kid = s15.row(["KID: management and other costs / transaction costs", f"={DB('fs_ter')}", f"={DB('kid_tc')}"], fmt=[None, P2, P2], expect={1: (0.0074, 1e-9), 2: (0.0002, 1e-9)})
r_tdnet = s15.row(["Average tracking difference 2016–2025 (fund − index, USD) / TER", f"={CB.q(s7.ws.title)}!$D${r_tda}", f"={DB('fs_ter')}"], fmt=[None, "0.000%", P2])
r_rec = s15.row(["Part of the TER recovered on average (TER + tracking difference)", f"=C{r_tdnet}+B{r_tdnet}"], fmt=[None, "0.000%"], expect={1: (0.0074 + TD_AVG, 1e-9)})
s15.widths(46, 24, 24)
T["rec"] = s15.tag(r_rec, 2); T["eimi_fee"] = s15.tag(F15["F00000TJCR"], 3)

# ---- Tracking & securities lending
s16 = BK.sheet(SN("track", "Tracking and lending"), "Annual report tracking figures; securities lending history (product page)", slide=NO["track"],
               title=f"Slide {NO['track']} – Fund overview: tracking difference, tracking error and securities lending",
               note="Links to 'Data documents'. Tracking difference gross of TER = fund − index + TER (annual report definition).")
s16.header(["Annual report, year to 30.11.2025 [A1 p.18]", "Value"])
r_ar = s16.row(["Fund return, USD", f"={DB('ar_td')}"], fmt=[None, P2])
r_ar2 = s16.row(["Benchmark return, USD", f"={DC('ar_td')}"], fmt=[None, P2])
r_ar3 = s16.row(["Fund − benchmark (net of TER)", f"=B{r_ar}-B{r_ar2}"], fmt=[None, P2], expect={1: (0.0019, 1e-9)})
r_ar4 = s16.row(["Tracking difference gross of TER (published)", f"={DC('ar_td2')}"], fmt=[None, P2])
r_ar5 = s16.row(["Check: net difference + TER", f"=B{r_ar3}+{DB('ar_td2')}"], fmt=[None, P2], expect={1: (0.0093, 1e-9)})
r_ar6 = s16.row(["Realised tracking error", f"={DB('ar_te')}"], fmt=[None, P2])
r_ar7 = s16.row(["Anticipated tracking error (up to)", f"={DC('ar_te')}"], fmt=[None, P2])
r_ar8 = s16.row(["Realised above anticipated by, pp", f"=B{r_ar6}-B{r_ar7}"], fmt=[None, P2], expect={1: (0.0006, 1e-9)})
s16.gap()
h16 = s16.header(["Year to 30 Sep", "Lending return", "Average on loan, % of AUM", "Maximum on loan, % of AUM", "Collateral, % of loan"])
L16 = {}
for lb, *_ in SL:
    L16[lb] = s16.row([lb] + [f"={DDn}!${c}${SLR[lb]}" for c in "BCDE"], fmt=[None, P2, P2, P2, P2])
s16.chart("bar", (L16[SL[0][0]], L16[SL[-1][0]], 1), [(h16, L16[SL[0][0]], L16[SL[-1][0]], 3), (h16, L16[SL[0][0]], L16[SL[-1][0]], 4)], "G2",
          "Securities on loan, % of AUM: average and maximum", colors=["AFAFAF", "191824"], fmt="0%")
s16.widths(42, 14, 20, 20, 18)
T["ar_net"] = s16.tag(r_ar3, 2); T["ar_te"] = s16.tag(r_ar8, 2)

# ---- Index history
s17 = BK.sheet(SN("index_hist", "Index history"), "MSCI Poland vs MSCI EM and ACWI IMI, annual 2012–2025 (USD); risk statistics", slide=NO["index_hist"],
               title=f"Slide {NO['index_hist']} – Index overview: MSCI Poland long history (USD, MSCI index factsheet [I1 p.1])",
               note="Links to 'Data documents'. Difference = MSCI Poland − MSCI Emerging Markets.")
h17 = s17.header(["Year", "MSCI Poland", "MSCI Emerging Markets", "MSCI ACWI IMI", "Poland − EM, pp"])
I17 = {}
for y in IDX_CY:
    I17[y] = s17.row([str(y), f"={DB(f'ix_{y}')}", f"={DC(f'ix_{y}')}", f"={DD4(f'ix_{y}')}", f"=B{s17.r}-C{s17.r}"], fmt=[None, P2, P2, P2, P2])
r_up = s17.row(["Years Poland ahead of EM (2012–2025)", None, None, None, f"=COUNTIF(E{I17[2012]}:E{I17[2025]},\">0\")"], fmt=[None, None, None, None, "0"], expect={4: (IDX_UP, 0)})
s17.chart("bar", (I17[2012], I17[2025], 1), [(h17, I17[2012], I17[2025], 2), (h17, I17[2012], I17[2025], 3)], "G2", "Annual net return, USD", colors=["C01E1E", "AFAFAF"], fmt="0%")
s17.gap()
s17.header(["Risk / return to 30.10.2026", "MSCI Poland", "MSCI EM", "MSCI ACWI IMI"])
R17 = {}
for k_ in ("sd3", "sd5", "sd10", "mdd", "r10", "rsi"):
    R17[k_] = s17.row([dd_.ws.cell(DR["ir_" + k_], 1).value.split(": ")[1].split(" – ")[0], f"={DB('ir_' + k_)}", f"={DC('ir_' + k_)}", f"={DD4('ir_' + k_)}"], fmt=[None, P2, P2, P2])
s17.widths(40, 14, 20, 16, 16)
T["idx_up"] = s17.tag(r_up, 5)

# ---- ESG vs alternatives
s18 = BK.sheet(SN("esg_ms", "ESG vs alternatives"), "Morningstar ESG indicators 30.09.2026, fund vs alternatives and the Article 6 MSCI World ETF", slide=NO["esg_ms"],
               title=f"Slide {NO['esg_ms']} – ESG overview: Morningstar indicators vs alternatives (30.09.2026)", note="Links to 'Data Morningstar ESG' (column 2026-09).")
s18.header(["Indicator"] + [n for _, n, _ in EFUNDS])
S18r = {}
for dp, nm in EDP:
    S18r[dp] = s18.row([nm] + [f"={ESref(dp, i_, '2026-09')}" if ESV(dp, i_, "2026-09") is not None else None for i_, _, _ in EFUNDS], fmt=[None] + ["0.00"] * len(EFUNDS),
                       expect={j + 1: (ESV(dp, i_, "2026-09"), 1e-9) for j, (i_, _, _) in enumerate(EFUNDS) if ESV(dp, i_, "2026-09") is not None})
r_sev = s18.row(["Severe ESG risk, % of covered portfolio (snapshot)"] + [f"={SNref(i_, 'ESG42')}" if SNAP[i_].get("ESG42") is not None else None for i_, _, _ in EFUNDS],
                fmt=[None] + ["0.00"] * len(EFUNDS))
s18.widths(50, *(18,) * len(EFUNDS))
T["waci"] = s18.tag(S18r["ESG2H"], 2); T["coal"] = s18.tag(S18r["ESG62"], 2); T["sev"] = s18.tag(r_sev, 2)

# ---- ESG trends
s19 = BK.sheet(SN("esg_trend", "ESG trends"), "Morningstar ESG indicators by date (12/2023–09/2026), bars by date", slide=NO["esg_trend"],
               title=f"Slide {NO['esg_trend']} – ESG overview: 3-year trends (Morningstar)", note="Same values as on the slide; links to 'Data Morningstar ESG'.")
TQ = ["2023-12", "2024-12", "2025-12", "2026-09"]
for dp, nm in (EDP[0], EDP[1], EDP[2]):
    s19.text(nm, CB.BOLD)
    h = s19.header(["Fund"] + TQ)
    r0 = s19.r
    for i_, n_, _ in EFUNDS:
        s19.row([n_] + [f"={ESref(dp, i_, q)}" if ESV(dp, i_, q) is not None else None for q in TQ], fmt=[None] + ["0.00"] * 4,
                expect={j + 1: (ESV(dp, i_, q), 1e-9) for j, q in enumerate(TQ) if ESV(dp, i_, q) is not None})
    s19.chart("bar", (r0, s19.r - 1, 1), [(h, r0, s19.r - 1, c) for c in range(2, 6)], f"H{h}", nm, colors=["DCDCDC", "AFAFAF", "6E6E6E", "191824"], h=7)
    s19.gap(10)
s19.widths(32, 11, 11, 11, 11)

# ---- Heat map: fund − EM
s22 = BK.sheet(SN("perf_heat", "Heat map vs EM"), "Monthly return of the fund minus Core MSCI EM IMI ETF, pp, EUR, by year and month", slide=NO["perf_heat"],
               title=f"Slide {NO['perf_heat']} – Performance: month by month vs emerging markets (fund − iShares Core MSCI EM IMI ETF, pp, EUR)",
               note="Cells link to column I of 'Calc fund monthly'; Year = compounded fund − compounded EM ETF.")
MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
s22.header(["Year"] + MON + ["Year, pp"])
HEAT, HL, H22r = [], [], {}
HY = list(range(2015, 2027))
for y in HY:
    row, xs = [], []
    for mi in range(12):
        k = f"{y}-{mi + 1:02d}"
        if k <= LM:
            row.append(round(FUNDM[k] - EM[k], 1)); xs.append(f"={CFn}!$I${CFr[k]}")
        else:
            row.append(None); xs.append(None)
    ks = [k for k in months(f"{y}-01", f"{y}-12") if k <= LM]
    yv = (comp(FUNDM, ks) - comp(EM, ks)) * 100
    row.append(round(yv, 1))
    HEAT.append(row); HL.append(str(y) + (" (to Oct)" if y == 2026 else ""))
    H22r[y] = s22.row([HL[-1]] + xs + [f"=(EXP(SUMPRODUCT(LN(1+{MMrng(NF, ks[0], ks[-1])}/100)))-EXP(SUMPRODUCT(LN(1+{MMrng(NEM, ks[0], ks[-1])}/100))))*100"],
                      fmt="0.0", expect={13: (yv, 1e-9)})
s22.heat(f"B{H22r[2015]}:M{H22r[2026]}")
s22.widths(14, *(7,) * 12, 10)
T["h2020"] = s22.tag(H22r[2020], 14); T["h2025"] = s22.tag(H22r[2025], 14); T["h2019"] = s22.tag(H22r[2019], 14)

# ---- Wins and losses
s23 = BK.sheet(SN("perf_risk", "Wins and losses"), "Worst fund months since 07/2014 vs EM ETF; up / down capture vs EM ETF; correlation", slide=NO["perf_risk"],
               title=f"Slide {NO['perf_risk']} – Performance: when the fund wins and loses vs emerging markets (07/2014–10/2026, EUR)", note="Links to 'Calc fund monthly'.")
s23.header(["Worst fund months", "This fund", "EM IMI ETF", "Difference, pp"])
W23 = {}
for k in WORST:
    W23[k] = s23.row([k, f"={CFn}!$B${CFr[k]}/100", f"={CFn}!$D${CFr[k]}/100", f"=B{s23.r}-C{s23.r}"], fmt=[None, P2, P2, P2],
                     expect={1: (FUNDM[k] / 100, 1e-9), 3: ((FUNDM[k] - EM[k]) / 100, 1e-9)})
s23.gap()
s23.header(["Capture vs EM IMI ETF (monthly, geometric)", "Value", "Months"])
_cdn = (f"(EXP(SUMPRODUCT(({CFR('D', '2014-07')}<0)*LN(1+{CFR('B', '2014-07')}/100))/COUNTIF({CFR('D', '2014-07')},\"<0\"))-1)/"
        f"(EXP(SUMPRODUCT(({CFR('D', '2014-07')}<0)*LN(1+{CFR('D', '2014-07')}/100))/COUNTIF({CFR('D', '2014-07')},\"<0\"))-1)")
_cup = _cdn.replace("<0", ">=0")
r_cd = s23.row(["Downside capture", "=" + _cdn, f"=COUNTIF({CFR('D', '2014-07')},\"<0\")"], fmt=[None, P, "0"], expect={1: (CAPD[0], 1e-9), 2: (CAPD[1], 0)})
r_cu = s23.row(["Upside capture", "=" + _cup, f"=COUNTIF({CFR('D', '2014-07')},\">=0\")"], fmt=[None, P, "0"], expect={1: (CAPU[0], 1e-9), 2: (CAPU[1], 0)})
r_co = s23.row(["Correlation of monthly returns, fund vs EM ETF", f"=CORREL({CFR('B', '2014-07')},{CFR('D', '2014-07')})"], fmt=[None, "0.00"], expect={1: (CORR, 1e-9)})
s23.widths(42, 14, 14, 14)
T["capd"] = s23.tag(r_cd, 2); T["capu"] = s23.tag(r_cu, 2); T["corr"] = s23.tag(r_co, 2); T["w0"] = s23.tag(W23[WORST[0]], 2)

# ---- Drawdowns live (daily)
s24 = BK.sheet(SN("perf_dd_live", "Drawdowns live"), "Daily falls from high since 01.11.2024; fund falls ≥ 8% and the alternatives over the same dates", slide=NO["perf_dd_live"],
               title=f"Slide {NO['perf_dd_live']} – Performance: drawdowns in the last two years (daily, EUR)",
               note="Links to 'Data Morningstar daily'. Episode = fall of this fund's NAV of 8% or more, peak to low; other funds over the same dates.")
s24.header(["Episode (fund peak → low)", "This fund (NAV)", "XETRA IBCJ", "Amundi E. Europe ex Russia", "EM IMI ETF", "MSCI World ETF", "Fund − EM, pp"])
E24 = {}
for a, z, f in EPD:
    r = s24.r
    E24[(a, z)] = s24.row([f"{fm(a)} → {fm(z)}"] + [f"={DYref(c, z)}/{DYref(c, a)}-1" for c in ("fund", "xet", "cec", "em", "w")] + [f"=B{r}-E{r}"],
                          fmt=[None] + [P2] * 6, expect={1: (f, 1e-9), 4: (win(DEM, a, z), 1e-9), 6: (f - win(DEM, a, z), 1e-9)})
s24.gap()
r_md = s24.row(["Largest fall since 01.11.2024: fund / EM / Amundi / World", f"=MIN({DYn}!$G${DY1}:$G${DYN})", f"=MIN({DYn}!$H${DY1}:$H${DYN})",
                f"=MIN({DYn}!$I${DY1}:$I${DYN})", f"=MIN({DYn}!$J${DY1}:$J${DYN})"],
               fmt=[None, P2, P2, P2, P2], expect={1: (MDDD["fund"], 1e-9), 2: (MDDD["em"], 1e-9), 3: (MDDD["cec"], 1e-9), 4: (MDDD["w"], 1e-9)})
r_n20 = s24.row(["One day, 19.11.2026 → 20.11.2026: fund NAV", f"={DYref('fund', '2026-11-20')}/{DYref('fund', '2026-11-19')}-1"], fmt=[None, P2], expect={1: (NOV20, 1e-9)})
s24.widths(34, 14, 14, 18, 12, 14, 14)
dy.chart("line", (DY1, DYN, 1), [(DY1 - 1, DY1, DYN, 7), (DY1 - 1, DY1, DYN, 8), (DY1 - 1, DY1, DYN, 9)], "L2", f"Fall from the previous high, daily, EUR (slide {NO['perf_dd_live']})",
         colors=["C01E1E", "AFAFAF", "191824"], fmt="0%", w=30, h=12)
T["dd_live"] = s24.tag(r_md, 2); T["n20"] = s24.tag(r_n20, 2)

# ---- Drawdowns history (monthly)
s25 = BK.sheet(SN("perf_dd_hist", "Drawdowns history"), "Fund falls ≥ 15% (month-end, 02/2011–10/2026) and the World / EM ETFs over the same months; Morningstar capture ratios", slide=NO["perf_dd_hist"],
               title=f"Slide {NO['perf_dd_hist']} – Performance: drawdowns 2011–2026 (monthly, EUR)", note="Links to 'Calc fund monthly' (index columns E, F, J, K) and 'Data Morningstar snapshot'.")
s25.header(["Episode (fund peak month-end → low month-end)", "This fund", "MSCI World ETF", "EM IMI ETF", "Months peak → low"])
E25 = {}
for a, z, f in EPM:
    ra = CF0 if a == "2011-01" else CFr[a]
    em_ok = a >= "2014-07"
    E25[(a, z)] = s25.row([f"{a} → {z}", f"={CFn}!$E${CFr[z]}/{CFn}!$E${ra}-1", f"={CFn}!$F${CFr[z]}/{CFn}!$F${ra}-1",
                           f"={CFn}!$K${CFr[z]}/{CFn}!$K${CFr[a]}-1" if em_ok else "n/a", len(months(a, z)) - 1],
                          fmt=[None, P2, P2, P2, "0"], expect={1: (f, 1e-9), 2: (VW[z] / VW[a] - 1, 1e-9)})
r_mdd = s25.row(["Largest fall from high, whole period", f"=MIN({CFR('G')})", f"=MIN({CFR('H')})", None], fmt=[None, P2, P2], expect={1: (MDD_F, 1e-9), 2: (MDD_W, 1e-9)})
r_mdde = s25.row(["Largest fall from high since 06/2014: fund / EM ETF", f"=MIN({CFR('L', '2014-07')})", None, f"=MIN({CFR('M', '2014-07')})"], fmt=[None, P2, None, P2],
                 expect={1: (MDD_FE, 1e-9), 3: (MDD_EM, 1e-9)})
s25.widths(44, 14, 16, 14, 16)
T["mdd_h"] = s25.tag(r_mdd, 2); T["mdde"] = s25.tag(r_mdde, 4)

# ---- Data check
s26 = BK.sheet(SN("perf_check", "Data check"), "Factsheet USD returns vs index; Morningstar NAV vs XETRA line in EUR", slide=NO["perf_check"],
               title=f"Slide {NO['perf_check']} – Performance: check of the data used", note="Factsheet values link to 'Data documents'; own values are formulas on 'Data Morningstar monthly'.")
s26.header(["Calendar year (EUR, own)", "Fund NAV", "XETRA IBCJ", "Difference, pp"])
C26 = {}
for y in range(2012, 2026):
    ks = months(f"{y}-01", f"{y}-12")
    C26[y] = s26.row([str(y), FCOMP(NF, ks[0], ks[-1]), FCOMP(NX, ks[0], ks[-1]), f"=C{s26.r}-B{s26.r}"], fmt=[None, P2, P2, P2],
                     expect={1: (comp(FUNDM, ks), 1e-9), 2: (comp(XET, ks), 1e-9)})
XK = months("2011-07", LM)
MAXD = max(abs(FUNDM[k] - XET[k]) for k in XK)
r_mx = s26.row(["Largest monthly gap NAV vs XETRA, 07/2011–10/2026, pp", f"=SUMPRODUCT(MAX(ABS({MMrng(NF, '2011-07', LM)}-{MMrng(NX, '2011-07', LM)})))/100"], fmt=[None, P2],
               expect={1: (MAXD / 100, 1e-9)})
s26.widths(36, 14, 14, 14)
T["maxd"] = s26.tag(r_mx, 2)

# ---- Alternatives table
ALT = [("0P0000TJXS", "This fund: iShares MSCI Poland (IBCJ GY)", "M2"), ("F00001T1YT", "Expat Poland WIG20 (PLX GY)", "M3"),
       ("F000013D6B", "Amundi MSCI E. Europe ex Russia (LEER GY)", "M4"), ("F00000TJCR", "iShares Core MSCI EM IMI", "M5")]
s27 = BK.sheet(SN("alt1", "Alternatives"), "Alternatives: cost, size, SFDR, benchmark, holdings, inception (Morningstar)", slide=NO["alt1"],
               title=f"Slide {NO['alt1']} – Possible financial instrument alternatives", note="Links to 'Data Morningstar snapshot'. Expat: management fee (no ongoing charge returned).")
s27.header(["Fund", "Ongoing charge / fee, %", "Fund size, EUR m", "SFDR (Morningstar)", "Benchmark", "Holdings", "Inception", "Portfolio date"])
A27 = {}
for i_, n_, _ in ALT:
    dp = "UB422" if SNAP[i_].get("UB422") is not None else "OS003"
    A27[i_] = s27.row([n_, f"={SNref(i_, dp)}", f"={SNref(i_, 'OF009')}/1000000", f"={SNref(i_, 'EU001')}", f"={SNref(i_, 'OF00L')}",
                       f"={SNref(i_, 'HS008')}" if SNAP[i_].get("HS008") else None, f"={SNref(i_, 'OS00F')}", f"={SNref(i_, 'HS000')}" if SNAP[i_].get("HS000") else None],
                      fmt=[None, "0.00", "#,##0.0"], expect={1: (SNAP[i_][dp], 1e-9), 2: (SNAP[i_]["OF009"] / 1e6, 1e-6)})
s27.widths(42, 16, 14, 14, 36, 10, 12, 12)

# ---- Alternatives performance
ALTP = [("This fund (NAV)", NF, FUNDM), ("Expat Poland WIG20", NEX, EXPAT), ("Amundi E. Europe ex Russia", NCEC, CEC), ("Core MSCI EM IMI", NEM, EM), ("Core MSCI World", NW, WLD)]
YA = [2020, 2021, 2022, 2023, 2024, 2025, 2026]
s28 = BK.sheet(SN("alt2", "Alternatives perf."), "Calendar-year returns and annualised 1/3/5 years, volatility 5 years, EUR", slide=NO["alt2"],
               title=f"Slide {NO['alt2']} – Alternatives: performance, EUR", note="Year = monthly returns compounded; 2026 = January–October.")
h28 = s28.header(["Fund"] + [str(y) if y < 2026 else "2026 Jan–Oct" for y in YA])
A28 = {}
for n_, col, s in ALTP:
    vals, exp = [], {}
    for j, y in enumerate(YA):
        ks = months(f"{y}-01", f"{y}-12" if y < 2026 else LM)
        vals.append(FCOMP(col, ks[0], ks[-1])); exp[j + 1] = (comp(s, ks), 1e-9)
    A28[n_] = s28.row([n_] + vals, fmt=[None] + [P] * len(YA), expect=exp)
s28.chart("bar", (h28 + 1, h28 + len(ALTP), 1), [(h28, h28 + 1, h28 + len(ALTP), c) for c in range(2, 2 + len(YA))], "J2", "Calendar-year return, EUR", fmt="0%")
s28.gap(2)
s28.header(["Fund", "1 year a year", "3 years a year", "5 years a year", "Volatility 5 years", "Difference to this fund, 5 years, pp"])
A28b = {}
R28F = s28.r
for n_, col, s in ALTP:
    i = s28.r
    A28b[n_] = s28.row([n_] + [FANN(col, ks[0], LM) for ks in (K1Y, K3, K5)] + [FVOL(col, K5[0], LM), f"=D{i}-D{R28F}"],
                       fmt=[None] + [P2] * 5, expect={1: (ann(s, K1Y), 1e-9), 2: (ann(s, K3), 1e-9), 3: (ann(s, K5), 1e-9), 4: (vol(s, K5), 1e-9)})
s28.widths(30, *(13,) * 7)
T["cec5"] = s28.tag(A28b["Amundi E. Europe ex Russia"], 4); T["f5"] = s28.tag(A28b["This fund (NAV)"], 4); T["ex5"] = s28.tag(A28b["Expat Poland WIG20"], 4)
T["cecd"] = s28.tag(A28b["Amundi E. Europe ex Russia"], 6, ann(CEC, K5) - ann(FUNDM, K5))

# ---- Alternatives monthly heat
M12 = months("2025-11", LM)
s29 = BK.sheet(SN("alt_heat", "Alternatives monthly"), "Monthly returns, last 12 months (11/2025–10/2026), EUR, %", slide=NO["alt_heat"],
               title=f"Slide {NO['alt_heat']} – Alternatives: monthly returns 11/2025–10/2026, EUR, %", note="Links to 'Data Morningstar monthly'; 12M = compounded.")
s29.header(["Fund"] + [m[2:4] + "/" + m[5:] for m in M12] + ["12M"])
A29 = {}
for n_, col, s in ALTP:
    A29[n_] = s29.row([n_] + [f"={MMref(col, m)}" for m in M12] + [f"=(EXP(SUMPRODUCT(LN(1+{MMrng(col, M12[0], M12[-1])}/100)))-1)*100"],
                      fmt="0.0", expect={13: (comp(s, M12) * 100, 1e-9)})
s29.heat(f"B{A29['This fund (NAV)']}:M{A29['Core MSCI World']}")
s29.widths(28, *(7,) * 12, 8)

ow = BK.sheet("Data internal", "Data from internal sources (Bloomberg, internal assessment), 23.11.2026",
              title="Internal data", note="Blue = input from the internal source named in column D (source, date), as in the DD.")
ow.header(["Item", "Value", "Unit", "Source"])
R_SPR = ow.row(["Average bid-ask spread, IBCJ GY Equity (XETRA), last 30 days to 23.11.2026", SPREAD, "%", OWN], fmt=[None, "0.000"], inputs=(1,))
R_VOL = ow.row(["90-day average daily traded volume, IBCJ GY Equity (XETRA), to 23.11.2026", AVGVOL, "shares (as provided: 252.9k)", OWN], fmt=[None, "#,##0"], inputs=(1,))
R_LIM = ow.row(["Classification for the pension fund's limits", "Europe equity exposure", "text", OWNI], inputs=(1,))
R_SPT = ow.row(["Spread as a share of TER 0.74% (own calculation)", f"=B{R_SPR}/0.74", "%", "mūsų skaičiavimas / own calculation; TER [F1 p.1]"], fmt=[None, "0.0%"])
ow.widths(62, 22, 26, 52)
T_SPR, T_VOL, T_SPT = ow.tag(R_SPR, 2, SPREAD), ow.tag(R_VOL, 2, AVGVOL), ow.tag(R_SPT, 2, SPREAD / 0.74)

NCHECK, BAD = BK.save(XLSX)
print("Excel", XLSX.name, NCHECK, "checks,", len(BAD), "differences", BAD[:5])
assert not BAD, BAD

# ======================================================================= POWERPOINT
W2 = (2.016, 9.978)
d = L.Deck()
MSF = SNAP["0P0000TJXS"]
A5F, A5E, A5W, A5C = ann(FUNDM, K5), ann(EM, K5), ann(WLD, K5), ann(CEC, K5)
ASI, ASW = ann(FUNDM, FK), ann(WLD, FK)
AE14F, AE14E = ann(FUNDM, EK), ann(EM, EK)
VOLF, VOLW, VOL5F, VOL5E = vol(FUNDM, FK), vol(WLD, FK), vol(FUNDM, K5), vol(EM, K5)

# 1 ---------------------------------------------------------------- title
s = d.slide(blank=True)
tb = s.shapes.add_textbox(*(Inches(v) for v in (0.602, 3.0, 12.129, 2.6)))
tf = tb.text_frame; tf.word_wrap = True
for k, (txt, size, major) in enumerate([
        (f"iShares V plc – {FUND}", 32, True),
        (f"Share class: USD (Acc), accumulating ({ISIN}); XETRA line IBCJ, EUR [F1 p.1][W1]", 18, False),
        ("Fund manager: BlackRock Asset Management Ireland Limited (management company); investment manager BlackRock Advisors (UK) Limited [K1 p.1][P1 p.199]", 18, False),
        ("Fund manager contact: iShares Investor Services Team – info@ishares.co.uk, +49 89 427295800 (general contact) [K1 p.3]", 18, False),
        ("Due diligence report prepared by: Analyst", 20, False),
        (f"Date: {REV}", 20, False)]):
    p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
    L.M.no_space(p)
    if major:
        r = p.add_run(); r.text = txt; L.M.run_fmt(r, size, major=True)
    else:
        L.add_runs(p, txt, size, slide_no=1)

# 2 ---------------------------------------------------------------- Conclusion
s = d.slide("Conclusion")
CW, CHh, CY = 1.93, 1.12, 1.40
cards = [("Decision", "Positive", "with the points to watch below; " + sl("decision"), GREENC),
         ("TER", "0.74%", "no entry / exit fee for exchange buyers [K1 p.3][F1 p.1]", AMBER),
         ("Fund size", "USD 1.1bn", "31.10.2026 [F1 p.1]; USD 0.48bn a year earlier [A1 p.253]", GREENC),
         ("SFDR", "Article 6", "“Other” – no ESG criteria [F1 p.1][A1 p.13]", DARK),
         ("Tracks the index", f"{TD_AVG * 100:+.2f} pp", f"a year vs MSCI Poland, 2016–2025, after TER {T['tda']}", GREENC),
         ("Last 5 years, a year", p0(A5F), f"EUR vs {p0(A5E)} emerging markets {T['y5']}", AMBER)]
for k, (lab, val, note, acc) in enumerate(cards):
    L.card(s, (0.602 + k * (CW + 0.11), CY, CW, CHh), lab, val, note, acc, slide_no=2)
BY = CY + CHh + 0.3
COLW = 3.93
blocks = [
    ("Why the fund", DARK, [
        "- **Only liquid way to buy the Polish large and mid caps in one UCITS ETF:** physical, full replication of MSCI Poland (16 stocks) [P1 p.89][F1 p.1]; the other Poland ETF is EUR 2m in size [J2][J1]",
        f"- **Tracks well:** fund − index {TD_AVG * 100:+.2f} pp a year 2016–2025 although TER is 0.74% {T['tda']} – securities lending and dividend tax recover most of the cost {T['rec']}",
        "- **Large and growing:** USD 1,121m (31.10.2026) [F1 p.1] vs USD 484m at 30.11.2025 [A1 p.253]; on XETRA in EUR (IBCJ GY) [W1]",
        f"- **Strong recent record:** {p0(A5F)} a year in EUR over 5 years vs {p0(A5E)} for emerging markets {T['y5']}"]),
    ("Risks to watch", AMBER, [
        f"- **Very volatile single country:** volatility {p0(VOLF)} a year vs {p0(VOLW)} for MSCI World {T['vol']}; ^^fall of {pc(MDD_F)} 01/2018–09/2022^^ {T['mdd']}; index fell 77.59% in 2007–2009 [I1 p.1]",
        f"- **Concentrated:** 16 stocks, top 3 = {p0(TOP3)} {T['top3']}; banks and insurers 47% [W1]; PKO BP {TOP[0][1]:.1f}% – the fund may hold up to 20% (35% in exceptional markets) of one issuer [P1 p.89]",
        f"- **Weak long-term return:** since launch 02/2011 {p0(ASI)} a year in EUR vs {p0(ASW)} for MSCI World {T['si']}",
        "- Realised tracking error 0.41% above the anticipated 0.35% in FY 2025 (dividend tax differences) [A1 p.18][A1 p.21]"]),
    ("Internal checks", DARK, [
        f"- **Limits:** the position is classified as **Europe equity exposure** for the pension fund's limits ({OWNI})",
        f"- **XETRA liquidity:** average bid-ask spread **{SPREAD:.3f}%** (30 days) and 90-day average daily volume **252.9k** shares ({OWN}) {T_SPR}{T_VOL}; the spread equals {SPREAD / 0.74:.0%} of the annual TER (own calculation) {T_SPT}",
        "- Polish dividend withholding tax: not assessed further (internal decision, 23.11.2026); the annual report names it as a tracking-error driver [A1 p.21]",
        f"- Questions and answers: slide {NO['questions']}; data not available: slide {NO['method']}"]),
]
for k, (hd, col, lines) in enumerate(blocks):
    x = 0.602 + k * (COLW + 0.17)
    L.heading(s, (x, BY, COLW, 0), hd, 13, col)
    L.body(s, (x, BY + 0.45, COLW, 6.8 - BY - 0.45), lines, 10, slide_no=2)

# 3 ---------------------------------------------------------------- Fund overview (index funds)
s = d.slide("Fund overview  (index funds)")
L.table(s, [
    ["ISIN", [f"**{ISIN}** – share class USD, accumulating; launched 21.01.2011 [F1 p.1]", "Sub-fund of iShares V plc, Ireland [K1 p.1]; base currency USD, not hedged [P1 p.89][J1]"]],
    ["Ticker", [f"**IBCJ GY Equity** – XETRA, EUR ({OWN}); same line in the iShares listing table: Deutsche Boerse Xetra, IBCJ, EUR, Bloomberg IBCJ GY [W1]",
                "Also: SPOL LN (GBP), IPOL LN (USD), SPOL BW [F1 p.3]; Mexico [F1 p.3]"]],
    ["Fees", ["Total expense ratio **0.74% a year** [F1 p.1]; KID: 0.74% management and other costs + 0.02% transaction costs; no performance fee [K1 p.3]",
              "No entry / exit fee for exchange buyers – broker costs only [K1 p.3]"]],
    ["AUM", f"**USD 1,120.83m** (31.10.2026) [F1 p.1]; USD 1,028m (20.11.2026) [W1]; Morningstar EUR {MSF['OF009'] / 1e6:,.0f}m ({fm(MSF['OF010'])}) [M2]"],
    ["Bid-ask spread", f"**{SPREAD:.3f}%** – average bid-ask spread of IBCJ GY Equity on XETRA, last 30 days to 23.11.2026 ({OWN}) {T_SPR}"],
    ["Liquidity (e.g. 90D Avg Volume)", [f"**252.9k** shares – 90-day average daily volume of IBCJ GY Equity on XETRA, to 23.11.2026 ({OWN}) {T_VOL}", "Fund size USD 1.1bn [F1 p.1]; 6 listings [J1]; underlying: 16 large and mid caps on the Warsaw exchange [I1 p.1]"]],
    ["Replication Method", "**Physical, full replication** – holds all index stocks in index proportions [P1 p.89][F1 p.1][J1]"],
    ["Tracking difference / tracking error", [f"Tracking difference **{TD_AVG * 100:+.2f} pp a year** on average 2016–2025 (fund − index, USD) {T['tda']}; FY 2025: +0.19 pp net, +0.93 pp gross of TER [A1 p.18]",
                                              f"Realised tracking error **0.41%** vs anticipated up to 0.35% [A1 p.18][P1 p.123] – details: slide {NO['track']}"]],
    ["Securities lending / CP risk management", ["Yes; fund keeps **62.5%** of income, BlackRock 37.5% [W1]; average on loan 10.74% of assets, max 18.94%, collateral 109% (year to 30.09.2026) [W1]",
                                                 "Prospectus: expected 0–39%, maximum 43% of NAV [P1 p.125]"]],
    ["Other relevant information (taxes, leverage, etc)", "Ireland, UK reporting fund [F1 p.1]; depositary State Street Custodial Services (Ireland) [K1 p.1]; may use derivatives (index futures) [K1 p.1]; borrowing only temporary, ≤ 10% [P1 p.231]"],
], (0.602, 1.396, 0, 0), W2, 10.5, slide_no=NO["overview"])

# 4 ---------------------------------------------------------------- Index overview
s = d.slide("Index overview (if index is different than used internally)", box=(0.602, 0.846, 12.129, 0.763))
L.table(s, [
    ["Index", ["**MSCI Poland Index** (Net Return, USD) – \"large and mid capitalisation stocks of the Polish equity market\" [P1 p.89]; 16 constituents, about 85% of the Polish equity universe [I1 p.1]",
               "Part of MSCI Emerging Markets (factsheet benchmark name: “MSCI Emerging - Poland in Net USD”) [F1 p.1]"]],
    ["Index weighting method", "Free-float market capitalisation weighted; **quarterly** rebalance [P1 p.89]; turnover 4.00% in the last 12 months [I1 p.1]"],
    ["Country / sector / company / factor actual exposures", ["**Poland 100%**; Financials 46.22%, Consumer Discretionary 15.65%, Energy 15.48%, Materials 7.18% [I1 p.2]",
                                                             "PKO Bank Polski 18.26%, Orlen 15.48%, Allegro 10.22%; top 10 = **85.39%** [I1 p.2]",
                                                             f"P/E 14.49, forward P/E 11.09, dividend yield 3.71% [I1 p.1]; details: slides {NO['portfolio']}, {NO['index_hist']}"]],
    ["Actual credit ratings", "N/A – equity index"],
    ["Interest / spread duration", "N/A – equity index"],
    ["FX exposure (principle)", "**Unhedged** – Polish zloty exposure; base currency USD; the EUR return of the XETRA line depends on PLN/EUR [P1 p.89][K1 p.2]"],
], (0.602, 1.609, 0, 0), W2, 11.5, slide_no=NO["index"])

# 5 ---------------------------------------------------------------- ESG overview I
ev = lambda dp, i_="0P0000TJXS": ESV(dp, i_, "2026-09")  # noqa: E731
CO = "0P0000MEHZ"
s = d.slide("ESG overview")
L.table(s, [
    ["SFDR article", "**Article 6** – factsheet “SFDR Classification: Other” (not Article 8 or 9) [F1 p.1][F1 p.4]; annual report: “Other” [A1 p.13]; Morningstar: not stated [M2]"],
    ["Exclusionary screening", "The fund excludes companies from the investment universe involved in the industries/reasons below"],
    ["No", f"Controversial weapons – no screen; the index is not ESG-screened and ESG integration does not add \"exclusionary screens\" [P1 p.17]; Morningstar: {ev('ESG53'):.2f}% involved vs {ev('ESG53', CO):.2f}% Article 6 MSCI World ETF [M2][M6]"],
    ["No", f"Tobacco – no screen [P1 p.17]; Morningstar: {ev('ESG63'):.2f}% vs {ev('ESG63', CO):.2f}% [M2][M6]"],
    ["No", f"Gambling – no screen [P1 p.17]; Morningstar: {ev('ESG55'):.2f}% vs {ev('ESG55', CO):.2f}% [M2][M6]"],
    ["No", f"Fossil fuel related screens (not fully excluded but sectors like thermal coal, oil sands, etc) – none; ^^thermal coal {ev('ESG62'):.2f}% involved vs {ev('ESG62', CO):.2f}%^^ {T['coal']}; Energy 13.31% of the fund (Orlen) [W1]"],
    ["No", "Red flag/signal related to controversies based on UN Global Compact, OECD Guidelines for Multinational Enterprises or any other relevant principles/standards – no fund-level rule; the fund does \"not commit to considering PAIs\" [P1 p.35]"],
    ["Other important exclusions:", "None – the fund holds all 16 index stocks [P1 p.89][I1 p.1]"],
    ["Positive screening", ""],
    ["n/a", f"The fund manager is UN-PRI signatory – !!n/a: not stated in the fund documents read (slide {NO['method']})!!"],
    ["No", "The fund implements positive screening – no; replicates a standard market-cap index [P1 p.89]"],
    ["Additional comments:", f"Morningstar ESG data uses the portfolio of 20.11.2026 (current) [M2]; ^^carbon intensity {ev('ESG2H'):.0f} vs {ev('ESG2H', CO):.0f} for MSCI World^^ {T['waci']}; "
                             f"severe ESG risk {MSF['ESG42']:.1f}% of portfolio {T['sev']}. Details: slides {NO['esg_ms']}–{NO['esg_trend']}."],
], (0.602, 1.396, 0, 0), W2, 10.5, bold_rows=(1,), bold_cells={(8, 0)}, slide_no=NO["esg1"])

# 6 ---------------------------------------------------------------- ESG overview II
s = d.slide("ESG overview")
L.table(s, [
    ["Engagement", "(voluntary field for passively managed funds)"],
    ["Yes", "The fund manager/company engages with the companies – \"Investment stewardship activities that are undertaken across all Equity Funds\"; engagement with index providers [P1 p.17]"],
    ["Additional comments", f"!!Voting and engagement record for the 16 Polish companies: n/a – not in the documents read (slide {NO['method']})!!"],
    ["Indicators", ""],
    ["No", "The indicators support the targets of sustainability and PAI indicators followed by the Fund – no sustainability targets; PAIs not committed [P1 p.35]"],
    ["Additional comments:", f"Morningstar indicators only (no fund ESG reporting): slides {NO['esg_ms']}–{NO['esg_trend']}"],
    ["Sanctions", ""],
    ["n/a", "Compliance with the list of sanctions (fund management company) (add a screenshot) – see " + sl("sanctions", "sanctions_1", "sanctions_2")],
    ["Yes", ["No issued identified from adverse media checks (add a screenshot) – see " + sl("media"),
             "We do not consider adverse media check Google search results as red flags and see no added value in inspecting board members of this international conglomerate."]],
    ["No issues identified/checks performed", "In case any red flags, risks, unclarities were identified then additional checks (sanctions screening and adverse media checks) on the fund’s Ultimate Beneficial Owners and/or senior management shall be performed."],
], (0.602, 1.396, 0, 0), W2, 11.5, bold_cells={(0, 0), (3, 0), (6, 0)}, slide_no=NO["esg2"])

# 7 ---------------------------------------------------------------- Performance
s = d.slide("Performance")
L.signed_bars(s, (0.602, 1.30, 8.0, 2.95), [str(y) for y in TDY], tuple(round(TD[y] * 100, 2) for y in TDY),
              title="Tracking difference: fund minus MSCI Poland, percentage points, each calendar year, USD, after TER", size=9, number_format='+0.00;-0.00;0.00')
L.table(s, [["Year"] + [str(y) for y in TDY], ["Fund, %"] + [f"{FS_CY[y][0]:.1f}" for y in TDY], ["MSCI Poland, %"] + [f"{FS_CY[y][1]:.1f}" for y in TDY]],
        (0.602, 4.27, 0, 0), (1.3,) + (0.67,) * len(TDY), 9, bold_cols=(0,), head=True, slide_no=NO["perf"])
L.data_note(s, (0.602, 5.17, 8.0, 0.45), "Data: factsheet calendar-year performance, share class vs benchmark, USD [F1 p.2]; bar = fund − index (own calculation) – Excel [C1] sheet "
            f"“{s7.ws.title}”.")
L.table(s, [["EUR, to 10/2026", "Fund", "EM / World"],
            ["Since 02/2011, a year", f"^^{p0(ASI)}^^", f"– / {p0(ASW)}"], ["Since 07/2014, a year", p0(AE14F), f"{p0(AE14E)} / –"],
            ["10 years, a year", p0(ann(FUNDM, K10)), f"{p0(ann(EM, K10))} / {p0(ann(WLD, K10))}"],
            ["5 years, a year", f"**{p0(A5F)}**", f"{p0(A5E)} / {p0(A5W)}"], ["3 years, a year", f"**{p0(ann(FUNDM, K3))}**", f"{p0(ann(EM, K3))} / {p0(ann(WLD, K3))}"],
            ["Volatility since 2011", f"^^{p0(VOLF)}^^", f"– / {p0(VOLW)}"], ["Largest fall (month-end)", f"^^{pc(MDD_F)}^^", f"– / {pc(MDD_W)}"]],
        (8.85, 1.40, 0, 0), (2.0, 0.85, 1.03), 9.5, bold_cols=(0,), head=True, slide_no=NO["perf"])
L.data_note(s, (8.85, 4.35, 3.9, 0.7), "Data: Morningstar monthly total returns in EUR, fund NAV [M1], iShares Core MSCI EM IMI ETF [M5], iShares Core MSCI World ETF [M6], read 23.11.2026; "
            f"own calculation – Excel [C1] sheet “{s7.ws.title}”.")
L.body(s, (0.602, 5.65, 12.129, 1.2), [
    f"- **Index tracking is good:** behind the index by only {abs(TD_AVG) * 100:.2f} pp a year on average with a 0.74% TER {T['tda']}; ahead in {TD_POS} of 10 years {T['tdp']}; worst 2020 {TD[2020] * 100:+.2f} pp {T['td2020']}",
    f"- **The market, not the fund, drives the result:** EUR {p0(A5F)} a year over 5 years {T['y5']}, but only {p0(ASI)} a year since 2011 {T['si']}. **Details:** slides {NO['track']}, {NO['index_hist']}, {NO['perf_heat']}–{NO['perf_check']}"],
    10.5, slide_no=NO["perf"])

# 8 ---------------------------------------------------------------- Investment team
s = d.slide("Investment team")
L.table(s, [
    ["Role", "Who"],
    ["Management company", "**BlackRock Asset Management Ireland Limited** (the “Manager”) – authorised and regulated by the Central Bank of Ireland [K1 p.1]; ultimately wholly owned by BlackRock, Inc. [P1 p.199]"],
    ["Investment manager", "**BlackRock Advisors (UK) Limited** – the Manager delegated investment management to it; regulated by the UK Financial Conduct Authority [P1 p.199][P1 p.12]"],
    ["Sub-investment managers", "BlackRock Asset Management North Asia Limited and BlackRock Asset Management Deutschland AG [A1 p.263]"],
    ["Fund managers", "Index team; the iShares documents read do not name individual portfolio managers [P1 p.199]"],
    ["Service providers", "Depositary: State Street Custodial Services (Ireland) Limited [K1 p.1]; administrator: State Street Fund Services (Ireland) Limited [P1 p.200][W1]"],
    ["Firm", "BlackRock group; iShares V plc is an Irish umbrella with segregated liability between funds [P1 p.1]"],
], (0.602, 1.40, 0, 0), (2.8, 9.193), 12, bold_cols=(0,), head=True, slide_no=NO["team"])

# 9 ---------------------------------------------------------------- Investment strategy & process
s = d.slide("Investment strategy & process")
L.table(s, [
    ["Topic", "What the fund does"],
    ["Objective", "\"a total return, taking into account both capital and income returns, which reflects the return of the MSCI Poland Index\" [P1 p.89]"],
    ["Management style", "**Passive, physical full replication**: \"holding all the securities comprising the Benchmark Index in a similar proportion to their weightings\" [P1 p.89]"],
    ["Single-issuer limit", f"Up to **20%** of NAV in one issuer, **35%** in exceptional market conditions [P1 p.89]; largest holding PKO BP {TOP[0][1]:.2f}% (20.11.2026) {T['pko']}"],
    ["Instruments", ["Equities; financial derivatives may be used, also for direct investment [K1 p.1]; index futures held (WIG20 future, 20.11.2026) [M2]",
                     "Securities lending up to 43% of NAV [P1 p.125]; 10.62% of NAV on loan at 30.11.2025 [A1 p.731]"]],
    ["Key risks", ["- **Concentration:** single country, sector and companies – \"more sensitive to any localised economic, market, political … events\" [F1 p.1]",
                   "- **Emerging-market risk:** liquidity, restrictions on transfer of assets [F1 p.1]; KID risk class **5 of 7** [K1 p.2]",
                   "- **Securities lending counterparty risk** – collateral 109% of loans [W1]"]],
], (0.602, 1.40, 0, 0), (2.4, 9.593), 11, bold_cols=(0,), head=True, slide_no=NO["strategy"])

# 10 --------------------------------------------------------------- Current portfolio characteristics
s = d.slide("Current portfolio characteristics", sub="including impact on current portfolio")
SHORTS = ["Financials", "Cons. Disc.", "Energy", "Materials", "Staples", "Comm.", "Utilities", "IT", "Industrials"]
L.column_chart(s, (0.602, 1.70, 7.2, 2.6), SHORTS, [("Fund 19.11.2026", tuple(a for _, a, _ in SECT[:-1])), ("MSCI Poland 30.10.2026", tuple(b for _, _, b in SECT[:-1]))],
               title="Sector weights, %", size=8, number_format='0.0', axis_format='0', palette=["191824", "AFAFAF"])
L.data_note(s, (0.602, 4.32, 7.2, 0.4), f"Data: product page sector breakdown [W1]; MSCI index factsheet [I1 p.2] – Excel [C1] sheet “{s10.ws.title}”.", 8.5)
L.table(s, [["Holding, 20.11.2026", "Weight"]] + [[n_, f"{w:.2f}%"] for n_, w in TOP[:10]] + [["Top 10 / top 3", f"{TOP10 * 100:.1f}% / {TOP3 * 100:.1f}%"]],
        (8.1, 1.40, 0, 0), (3.2, 1.4), 9, bold_cols=(0,), head=True, slide_no=NO["portfolio"])
L.data_note(s, (8.1, 5.55, 4.6, 0.4), f"Data: Morningstar holdings, portfolio 20.11.2026 [M2] – Excel [C1] sheet “{s10.ws.title}”.", 8.5)
L.heading(s, (0.602, 4.85, 7.2, 0), "Impact on our portfolio", 12)
L.body(s, (0.602, 5.25, 7.2, 1.6), [
    "- Adds **one country (Poland) and mostly banks, insurers, energy and retail** – 47% financials [W1]; P/E 14.5 for the index [I1 p.1]",
    f"- Three stocks = {p0(TOP3)} {T['top3']}: a position in this ETF is in practice a position in PKO BP, Orlen and Allegro",
    "- Valuation: P/E 18.49, P/B 2.26 (fund, factsheet) [F1 p.1]"], 9.5, slide_no=NO["portfolio"])

# 11 --------------------------------------------------------------- Other
s = d.slide("Other")
for k, (lab, txt) in enumerate([("Legal analysis", "usually standard investments do not need legal analysis."),
                                ("Operational risk analysis", "usually standard investments do not need operational risk analysis.")]):
    L.card(s, (0.602 + k * 6.12, 1.55, 5.95, 1.25), lab, "Not needed", txt, DARK, vsize=18)
L.heading(s, (0.602, 3.15, 5.95, 0), "Why no extra analysis", 13)
L.body(s, (0.602, 3.6, 5.95, 3.1), [
    f"- Since the {FUND} is a UCITS fund compliant with European regulatory standards, it typically does not require additional legal or operational risk analysis.",
    "- UCITS funds are subject to EU regulations, ensuring a high level of investor protection, which means the standard oversight and risk management practices are sufficient.",
    "- Physical replication; borrowing only temporary and ≤ 10% [P1 p.231]; securities lending is the only non-standard activity [P1 p.125]."], 12, slide_no=NO["other"])
L.heading(s, (6.72, 3.15, 5.95, 0), "Fit with our strategy", 13)
L.body(s, (6.72, 3.6, 5.95, 3.1), [
    f"- {FUND} aligns with investment strategy within the SAA by providing exposure to Polish large and mid-cap equities (Central and Eastern Europe / emerging markets).",
    "- The fund’s objective is to track the MSCI Poland Index (Net Return).",
    "- Additionally, as a passively managed UCITS-compliant fund, it aligns with regulatory standards and aligns with our long-term investment goals."], 12, slide_no=NO["other"])

# ======================================================================= ADDITIONAL ANALYSIS
d.divider()

s = d.slide("Conclusion and Investment decision")
y = L.table(s, [["SFDR", "Asset class", "Region", "UCITS"], ["Article 6 (“Other”) [F1 p.1]", "Equity (long only)", "Poland – Central and Eastern Europe", "Yes [K1 p.1][P1 p.10]"]],
            (0.602, 1.45, 0, 0), (3.0, 3.0, 3.0, 2.993), 12, bold_rows=(0,), slide_no=NO["decision"])
ctrl = [
    ["1", "Possible financial instrument alternatives are evaluated to determine the most attractive investment in terms of return and risk (historical returns, risk results, acquisition costs, management fees, settlement terms, etc.).", "Yes", sl("alt1", "alt2", "alt_heat", "alt3")],
    ["2", "A comprehensive analysis of the historical performance of the selected financial instrument is conducted.", "Yes", sl("perf", "index_hist", "perf_heat", "perf_risk", "perf_dd_live", "perf_dd_hist", "perf_check")],
    ["3", "A cost analysis of the financial instrument is performed.", "Yes", sl("overview", "fees", "track")],
    ["4", "The suitability of the financial instrument’s form for the fund is assessed (direct investment in corporate stocks/bonds, government bonds, harmonized collective investment schemes, alternative investment entities, etc.).", "Yes", sl("acc_law", "acc_evid", "acc_class")],
    ["5", "Other parameters of the investment instrument are evaluated according to the financial instrument’s prospectus, Key Investor Information Document (KIID), rules, data description (factsheet), or other available financial instrument documents.", "Yes", sl("strategy", "overview", "esg1", "acc_evid")],
    ["6", "It is assessed whether investing in the selected financial instrument is possible according to the pension fund’s strategy, laws, and the fund's rules and investment guidelines.", "Yes", sl("strategy", "questions", "acc_law", "acc_class")],
]
y = L.table(s, ctrl, (0.602, y + 0.25, 0, 0), (0.45, 9.0, 0.8, 1.743), 11, slide_no=NO["decision"])
y = L.table(s, [["Based on the latest performance results and risk assessments, the investment decision is", "Positive"]], (0.602, y + 0.25, 0, 0), (9.45, 2.543), 14, bold_cols=(1,), slide_no=NO["decision"])
L.data_note(s, (0.602, y + 0.08, 12.1, 0.45), f"The decision rests on the fund documents P1, K1, F1, A1, the index factsheet I1, Morningstar data M1–M6, and internal data: limit classification ({OWNI}) and XETRA liquidity ({OWN}). Data not available: slide {NO['method']}.", 9)

s = d.slide("Questions before investing and answers")
y = L.table(s, [
    ["#", "Question", "Why it matters", "Answer (23.11.2026)", "Decisive?"],
    ["1", "Does a single-country Polish equity position fit our limits for emerging markets / Central and Eastern Europe, and what size?",
     f"Volatility {p0(VOL5F)} a year over 5 years vs {p0(VOL5E)} for emerging markets {T['vol5']}; falls of {pc(MDD_F)} happened {T['mdd']}", f"Classified as Europe equity exposure ({OWNI})", "**Yes**"],
    ["2", "How liquid is IBCJ on XETRA (bid-ask spread, average daily volume) for our trade size?",
     "Not in public documents; the ETF is large (USD 1.1bn [F1 p.1]) but the EUR line may trade less than London", f"Spread {SPREAD:.3f}%, volume 252.9k shares a day ({OWN}) {T_SPR}{T_VOL}", "**Yes**"],
    ["3", "How is Polish dividend withholding tax treated for the fund?",
     "Main driver of the 0.41% tracking error [A1 p.21]", "Not assessed further (internal decision, 23.11.2026)", "No"],
    ["4", "Do we accept 0.74% when the Amundi Eastern Europe ETF costs 0.50% and moved almost the same?",
     f"5 years: Amundi {pp(ann(CEC, K5) - A5F)} a year vs this fund {T['cecd']}; but it is a swap-based regional fund [J3]", "Different exposure: regional, swap-based [J3]; this fund retained", "No"],
    ["5", "Who is the BlackRock contact?", "No named contact in the public documents", "iShares Investor Services Team, info@ishares.co.uk [K1 p.3]", "No"],
], (0.602, 1.40, 0, 0), (0.4, 3.9, 3.9, 2.85, 0.879), 10, bold_cols=(0,), head=True, slide_no=NO["questions"])
L.data_note(s, (0.602, y + 0.08, 12.1, 0.4), f"Questions 1 and 2 decide the size and the way of buying; both answered with internal data (Excel [C1] sheet “{ow.ws.title}”). Data not available: slide {NO['method']}.")

# fees
s = d.slide("Fund overview: fees")
rows = [["Fund", "Ongoing charge", "Difference to this fund", "Note"]]
for n_, i_ in FEEF:
    dp = "UB422" if SNAP[i_].get("UB422") is not None else "OS003"
    v = SNAP[i_][dp]
    rows.append([("**" + n_ + "**") if i_ == "0P0000TJXS" else n_, f"{v:.2f}%" + ("" if dp == "UB422" else " (mgmt fee)"),
                 "–" if i_ == "0P0000TJXS" else pp((v - 0.74) / 100, 2), {"0P0000TJXS": "Poland, physical [K1 p.3]", "F00001T1YT": "Poland WIG20, EUR 2m; TER 1.38% on justETF [J2]",
                                                                         "F000013D6B": "regional, swap-based [J3]", "F00000TJCR": "all emerging markets [M5]"}[i_]])
y = L.table(s, rows, (0.602, 1.40, 0, 0), (3.4, 1.6, 1.8, 2.8), 10.5, bold_cols=(0,), head=True, slide_no=NO["fees"])
L.data_note(s, (0.602, y + 0.05, 9.6, 0.4), f"Data: Morningstar KIID ongoing charge (UB422) or management fee (OS003), read 23.11.2026 [M2][M3][M4][M5] – Excel [C1] sheet “{s15.ws.title}”.")
L.body(s, (0.602, y + 0.55, 6.4, 6.85 - y - 0.55), [
    f"- **Real cost is lower than the TER:** fund − index {TD_AVG * 100:+.2f} pp a year 2016–2025 {T['tda']} – about {(0.0074 + TD_AVG) * 100:.2f} pp of the 0.74% came back {T['rec']} (securities lending income, dividend tax) [A1 p.18]",
    "- Exchange buyers pay no entry or exit fee, only broker costs; transaction costs inside the fund 0.02% [K1 p.3]",
    "- No cheaper ETF on the same index; the WIG20 ETF is dearer and tiny [J2][M3]"], 10, slide_no=NO["fees"])
png, sz = L.pdf_crop(LOCAL["K1"], 3, "Composition of Costs", TMPI / "kid_costs.png", pad=(6, 10, 0, 150), width_pt=560)
L.heading(s, (7.25, y + 0.5, 5.48, 0), "Copy of the official document", 12)
L.evidence_picture(s, png, sz, (7.25, y + 0.95, 5.48, 6.3 - y - 0.95), "PRIIPs KID USD (Acc) (IE00B4M7GH52), 09.04.2026, “Composition of Costs” [K1 p.3]", slide_no=NO["fees"])

# tracking & lending
s = d.slide("Fund overview: tracking and securities lending")
y = L.table(s, [["Annual report, year to 30.11.2025", "Value"], ["Fund return / benchmark return, USD", "60.52% / 60.33%"], ["Difference after TER (own)", f"++{pp(0.0019, 2)}++"],
                ["TER / tracking difference gross of TER", "0.74% / 0.93%"], ["Realised / anticipated tracking error", "^^0.41%^^ / up to 0.35%"],
                ["Main drivers", "TER compounding; dividend tax rate vs index assumption"]],
            (0.602, 1.40, 0, 0), (3.2, 2.9), 10, bold_cols=(0,), head=True, slide_no=NO["track"])
L.data_note(s, (0.602, y + 0.03, 6.1, 0.5), f"Data: [A1 p.18] table and notes a, m [A1 p.21]; anticipated tracking error also [P1 p.123]; difference own calculation – Excel [C1] sheet “{s16.ws.title}”.", 8.5)
L.column_chart(s, (6.95, 1.30, 5.8, 3.0), [x[0] for x in SL], [("Average on loan", tuple(x[2] for x in SL)), ("Maximum on loan", tuple(x[3] for x in SL))],
               title="Securities on loan, % of fund assets, years to 30 September", size=8, number_format='0', axis_format='0', palette=["AFAFAF", "191824"])
L.data_note(s, (6.95, 4.32, 5.8, 0.5), f"Data: product page securities lending table [W1]; lending return 0.01–0.17% a year; collateral 108–112% of loans – Excel [C1] sheet “{s16.ws.title}”.", 8.5)
L.body(s, (0.602, 4.9, 12.129, 1.95), [
    f"- **FY 2025:** the fund beat the index by {pp(0.0019, 2)} after a 0.74% TER {T['ar_net']}; ^^realised tracking error above the anticipated level by {pp(0.0006, 2)}^^ {T['ar_te']} because of dividend tax differences [A1 p.21]",
    "- **Securities lending** is moderate: on average 7–33% of assets lent, maximum 57.5% in 2016/17 and 18.94% in 2025/26 [W1]; prospectus cap 43% [P1 p.125]; the fund keeps 62.5% of income [W1]",
    "- Lending income is small (0.04% in 2025/26 [W1]) – most of the gap between TER and tracking difference comes from dividend tax treatment [A1 p.21]"], 10, slide_no=NO["track"])

# index history
s = d.slide("Index overview: MSCI Poland long history")
L.column_chart(s, (0.602, 1.30, 7.6, 3.3), [str(y) for y in IDX_CY], [("MSCI Poland", tuple(v[0] for v in IDX_CY.values())), ("MSCI Emerging Markets", tuple(v[1] for v in IDX_CY.values()))],
               title="Annual net return, USD, %", size=8, number_format='0', axis_format='0', palette=["C01E1E", "AFAFAF"])
L.data_note(s, (0.602, 4.62, 7.6, 0.45), f"Data: MSCI index factsheet, annual performance 2012–2025 [I1 p.1] – Excel [C1] sheet “{s17.ws.title}”.", 8.5)
y = L.table(s, [["To 30.10.2026, USD", "Poland", "EM", "ACWI IMI"], ["Volatility 3 years", "18.02%", "17.88%", "12.02%"], ["Volatility 5 years", "^^28.45%^^", "18.79%", "14.99%"],
                ["Volatility 10 years", "^^28.23%^^", "17.68%", "14.96%"], ["Maximum drawdown since 2000", "^^77.59%^^", "65.25%", "58.59%"],
                ["10 years a year", "10.41%", "8.15%", "12.10%"], ["Since 29.12.2000 a year", "^^5.91%^^", "8.73%", "7.54%"]],
            (8.45, 1.40, 0, 0), (1.9, 0.75, 0.75, 0.9), 9.5, bold_cols=(0,), head=True, slide_no=NO["index_hist"])
L.data_note(s, (8.45, y + 0.03, 4.3, 0.5), f"Data: [I1 p.1] index risk and return characteristics; maximum drawdown 29.10.2007–17.02.2009 – Excel [C1] sheet “{s17.ws.title}”.", 8.5)
L.body(s, (0.602, 5.2, 12.129, 1.65), [
    f"- **Boom and bust market:** Poland beat emerging markets in {IDX_UP} of 14 years {T['idx_up']}, but the gaps are huge both ways (2019 −24 pp, 2025 +41 pp) [I1 p.1]",
    "- Since 2000 the index returned 5.91% a year vs 8.73% for emerging markets [I1 p.1] – **the recent strength (2023, 2025) follows a weak decade**; a long holding period does not remove the single-country risk"], 10, slide_no=NO["index_hist"])


# ESG: Morningstar indicators vs alternatives
def mark(vals, lower_better=True):
    out, nums = [], [v for v in vals if isinstance(v, (int, float))]
    best = (min(nums) if lower_better else max(nums)) if nums else None
    for k, v in enumerate(vals):
        if not isinstance(v, (int, float)):
            out.append("–"); continue
        t = f"{v:.2f}" if abs(v) < 20 else f"{v:.0f}"
        if k == 0 and isinstance(vals[-1], (int, float)) and v > vals[-1]:
            t = "^^" + t + "^^"
        elif v == best:
            t = "++" + t + "++"
        out.append(t)
    return out


s = d.slide("ESG overview: Morningstar indicators vs alternatives")
rows = [["Indicator (Morningstar, 30.09.2026)"] + [n for _, n, _ in EFUNDS],
        ["SFDR article"] + ["Art. 6 “Other” [F1 p.1]", "not stated", "not stated", "not stated"]]
for dp, nm in EDP:
    rows.append([nm] + mark([ESV(dp, i_, "2026-09") for i_, _, _ in EFUNDS]))
rows += [["Severe ESG risk, % of covered portfolio"] + mark([SNAP[i_].get("ESG42") for i_, _, _ in EFUNDS]),
         ["Holdings date used by Morningstar"] + [fm(SNAP[i_]["HS000"]) for i_, _, _ in EFUNDS],
         ["Source"] + [f"[{m}]" for *_, m in EFUNDS]]
y = L.table(s, rows, (0.602, 1.40, 0, 0), (3.6, 2.15, 2.15, 2.15, 2.079), 9.5, bold_cols=(0,), head=True, shade_col=1, slide_no=NO["esg_ms"])
y += 0.35
L.data_note(s, (0.602, y + 0.05, 12.1, 0.5), f"Data: Morningstar connector, read 23.11.2026; values at 30.09.2026 copied as returned – Excel [C1] sheet “{s18.ws.title}”. "
            "Red: this fund worse than the Article 6 MSCI World ETF; green: best value in the row. Amundi: swap-based, Morningstar returned carbon values only.", 8.5)
L.body(s, (0.602, y + 0.6, 12.129, 6.85 - y - 0.6), [
    f"- **No ESG profile:** worse than the Article 6 MSCI World ETF on carbon intensity ({ev('ESG2H'):.0f} vs {ev('ESG2H', CO):.0f}) {T['waci']}, thermal coal ({ev('ESG62'):.2f}% vs {ev('ESG62', CO):.2f}%) {T['coal']} and severe ESG risk ({MSF['ESG42']:.1f}% vs {SNAP[CO]['ESG42']:.1f}%) {T['sev']}",
    "- Zero tobacco, gambling and controversial weapons – because no such company is in the 16-stock index, not because of a screen [P1 p.17]",
    "- ^^Low reliability:^^ the fund's carbon intensity jumps between single month-ends (e.g. 51.4 in 01/2025, 110.4 in 02/2026) without an index change [M2]"], 10, slide_no=NO["esg_ms"])

# ESG trends
s = d.slide("ESG overview: 3-year trends (Morningstar)")
SHORTE = ["This fund", "Amundi E. Europe", "Core EM IMI", "Core MSCI World (Art. 6)"]
TQL = ["12/2023", "12/2024", "12/2025", "09/2026"]
for k, (dp, ttl, fmt_) in enumerate([("ESG2H", "Carbon intensity, Scope 1+2 (t CO2e / USD m revenue)", "0"), ("ESG2A", "Historical carbon risk score", "0.0"),
                                     ("ESG62", "Thermal coal, % of portfolio involved", "0.0")]):
    x = 0.602 + k * 4.08
    L.column_chart(s, (x, 1.30, 3.95, 3.6), SHORTE, [(TQL[j], tuple(ESV(dp, i_, q) for i_, _, _ in EFUNDS)) for j, q in enumerate(TQ)], title=ttl, size=8,
                   number_format=fmt_, axis_format=fmt_, palette=["DCDCDC", "AFAFAF", "6E6E6E", "191824"])
L.data_note(s, (0.602, 4.95, 12.1, 0.45), "Data: Morningstar monthly ESG series, values at 12/2023, 12/2024, 12/2025 and 09/2026 (bars light to dark) [M2][M4][M5][M6]; gap = no value. "
            f"Portfolio dates are current at each point. Excel [C1] sheet “{s19.ws.title}”.", 8.5)
L.body(s, (0.602, 5.5, 12.129, 1.35), [
    f"- Carbon intensity of the fund fell from {ESV('ESG2H', '0P0000TJXS', '2023-12'):.0f} (12/2023) to {ESV('ESG2H', '0P0000TJXS', '2024-12'):.0f} (12/2024) and rose again to {ev('ESG2H'):.0f} (09/2026) [M2] – it changes with the index composition and company data, not with a fund policy [P1 p.89]",
    "- No improvement trend can be expected: the index has no ESG criteria [P1 p.89][P1 p.17]"], 10, slide_no=NO["esg_trend"])

# sanctions
L.sanctions_slides(d, SANC, slide_no=NO["sanctions"])

# media
s = d.slide("Media checks")
y = L.table(s, [L.MEDIA_HEAD,
    ["26.09.2022, updated 20.10.2026", "**Central Bank of Ireland** warning notice", "Unauthorised firms cloned the name, Central Bank number and address of BlackRock Asset Management Ireland Limited to deceive consumers; \"no connection between the Central Bank authorised firm and the Unauthorised Firm\"", "Fraudsters using the name of the management company", "Warning in force – BlackRock is the victim", "Low – check counterparties and payment instructions only via official channels", "[N1]"],
    ["11.2024 – 09.2026", "**Texas Attorney General**; Climate Litigation Database", "Texas and other states sued BlackRock, Inc., State Street and Vanguard (antitrust: coal holdings used to cut coal output); Vanguard settled 02/2026; court granted defendants partial judgment on several claims", "BlackRock, Inc. (parent)", "**Ongoing** for BlackRock", "Low for this fund – US parent, US antitrust claim; not the Irish management company", "[N2][N3]"],
    ["11.2021 – 11.2026", "Web search 23.11.2026", "No fines, sanctions or enforcement actions found for BlackRock Asset Management Ireland Limited, BlackRock Advisors (UK) Limited or iShares V plc in the documents and pages read", "Manager; investment manager; the fund", "–", "None", "search terms below"],
], (0.602, 1.50, 0, 0), L.MEDIA_W, 9, head=True, slide_no=NO["media"])
L.heading(s, (0.602, y + 0.2, 5.9, 0), "How the check was done", 12)
L.body(s, (0.602, y + 0.62, 5.9, 6.85 - y - 0.62), [
    "- **Names:** BlackRock Asset Management Ireland Limited; BlackRock Advisors (UK) Limited; iShares V plc; iShares MSCI Poland UCITS ETF",
    "- **Words:** fine, sanction, settlement, lawsuit, investigation, Central Bank of Ireland, FCA, SEC; **period:** 5 years (11.2021 – 11.2026); search date 23.11.2026; pages in the DD documents ZIP"], 10, slide_no=NO["media"])
L.heading(s, (6.85, y + 0.2, 5.9, 0), "Conclusion", 12)
L.body(s, (6.85, y + 0.62, 5.9, 6.85 - y - 0.62), [
    "- One clone-fraud warning (BlackRock is the victim) and one US lawsuit against the parent – **not a red flag** for this UCITS ETF",
    "- We do not consider adverse media check Google search results as red flags and see no added value in inspecting board members of this international conglomerate."], 10, slide_no=NO["media"])

# heat map vs EM
s = d.slide("Performance: month by month vs emerging markets")
y = L.heatmap(s, (0.602, 1.50, 12.129, 3.9), HL, MON + ["Year"], HEAT, size=8.5, scale=6.0, label_w=1.45, slide_no=NO["perf_heat"])
L.data_note(s, (0.602, y + 0.05, 12.1, 0.5), "Data: cell = fund (NAV) return − iShares Core MSCI EM IMI ETF return in the month, percentage points, EUR [M1][M5]; "
            f"Year = compounded fund − compounded EM ETF (own calculation) – Excel [C1] sheet “{s22.ws.title}”.", 8.5)
L.body(s, (0.602, y + 0.65, 12.129, 6.85 - y - 0.65), [
    f"- **Long red spells and short green bursts:** 2019 {HEAT[4][-1]:+.1f} pp {T['h2019']}, 2020 {HEAT[5][-1]:+.1f} pp {T['h2020']}, then 2025 {HEAT[10][-1]:+.1f} pp {T['h2025']}",
    "- Poland is not a smaller version of emerging markets – it moves on its own (banks, energy, zloty); use it as a deliberate country bet"], 10.5, slide_no=NO["perf_heat"])

# wins and losses
s = d.slide("Performance: when the fund wins and loses")
L.column_chart(s, (0.602, 1.30, 6.3, 3.4), list(WORST), [("This fund", tuple(round(FUNDM[k], 1) for k in WORST)), ("EM IMI ETF", tuple(round(EM[k], 1) for k in WORST))],
               title="Six worst fund months 07/2014–10/2026: monthly return, %, EUR", size=9, number_format='0.0', axis_format='0', palette=["C01E1E", "AFAFAF"])
L.data_note(s, (0.602, 4.72, 6.3, 0.45), f"Data: Morningstar monthly returns, EUR [M1][M5]; own selection – Excel [C1] sheet “{s23.ws.title}”.", 8.5)
y = L.table(s, [["Worst fund months", "This fund", "EM ETF", "Difference"]] +
            [[k, pc(FUNDM[k] / 100), pc(EM[k] / 100), f"^^{FUNDM[k] - EM[k]:+.1f} pp^^"] for k in WORST],
            (7.15, 1.45, 0, 0), (1.9, 1.25, 1.25, 1.2), 10, bold_cols=(0,), head=True, slide_no=NO["perf_risk"])
y = L.table(s, [["Relative risk vs EM ETF, 07/2014–10/2026", "Value"], ["Downside capture (monthly)", p0(CAPD[0], 0) + f" ({CAPD[1]} months)"],
                ["Upside capture (monthly)", p0(CAPU[0], 0) + f" ({CAPU[1]} months)"], ["Correlation of monthly returns", f"{CORR:.2f}"]],
            (7.15, y + 0.2, 0, 0), (3.3, 2.3), 10, bold_cols=(0,), head=True, slide_no=NO["perf_risk"])
L.data_note(s, (7.15, y + 0.03, 5.6, 0.4), f"Data: own calculation (capture = geometric average month of the fund ÷ of the EM ETF in its down / up months) – Excel [C1] sheet “{s23.ws.title}”.", 8.5)
L.body(s, (0.602, 5.3, 6.3, 1.55), [
    f"- Worst month 03/2020: {pc(FUNDM[WORST[0]] / 100)} {T['w0']}; in each of the six worst months the fund lost more than emerging markets",
    f"- Downside capture {p0(CAPD[0], 0)} {T['capd']}, upside {p0(CAPU[0], 0)} {T['capu']}: **falls harder than emerging markets** and gains only a little more in up months; correlation {CORR:.2f} {T['corr']}"],
    9.5, slide_no=NO["perf_risk"])

# drawdowns live
s = d.slide("Performance: drawdowns in the last two years (daily)")
MLB = {"01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr", "05": "May", "06": "Jun", "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"}
DLAB = [(MLB[x[5:7]] + " " + x[2:4]) if (i == 0 or x[5:7] != DK[i - 1][5:7]) and x[5:7] in ("01", "04", "07", "10") else " " * (i % 9 + 1) for i, x in enumerate(DK)]
ch = L.line_chart(s, (0.602, 1.30, 6.55, 3.35), DLAB, [("This fund (NAV)", tuple(round(FD["fund"][x] * 100, 2) for x in DK)),
                                                       ("Core MSCI EM IMI ETF", tuple(round(FD["em"][x] * 100, 2) for x in DK)),
                                                       ("Amundi E. Europe ex Russia", tuple(round(FD["cec"][x] * 100, 2) for x in DK))],
                  title=f"Fall from the previous high, %, daily, EUR, {fm(DK[0])}–{fm(DK[-1])}", size=8, axis_format='0', palette=["C01E1E", "AFAFAF", "191824"], widths=[2.25, 1.5, 1.5])
L.tick_skip(ch, 1)
L.data_note(s, (0.602, 4.68, 6.55, 0.55), "Data: Morningstar Daily Return Index (NAV / prices in EUR), read 23.11.2026 [M1][M4][M5]; line = value ÷ highest value so far − 1 (own calculation) – "
            f"Excel [C1] sheet “{dy.ws.title}” (chart there; table on “{s24.ws.title}”).", 8.5)
rows = [["Fund fall ≥ 8%", "Fund", "XETRA", "Amundi", "EM ETF", "World"]]
for a, z, f in EPD:
    rows.append([f"{fm(a)[:5]}.{a[2:4]}–{fm(z)[:5]}.{z[2:4]}", f"^^{pc(f)}^^", pc(win(DX, a, z)), pc(win(DCEC, a, z)), pc(win(DEM, a, z)), pc(win(DW, a, z))])
y = L.table(s, rows, (7.35, 1.40, 0, 0), (1.6, 0.8, 0.75, 0.8, 0.75, 0.68), 8.5, bold_cols=(0,), head=True, slide_no=NO["perf_dd_live"])
L.data_note(s, (7.35, y + 0.03, 5.38, 0.6), f"Data: same series; falls of the fund NAV of 8% or more, peak to low; other funds over the same dates – Excel [C1] sheet “{s24.ws.title}”.", 8)
L.body(s, (0.602, 5.3, 12.129, 1.55), [
    f"- Largest fall in the period {pc(MDDD['fund'])} for the fund vs {pc(MDDD['em'])} for emerging markets and {pc(MDDD['cec'])} for the Amundi Eastern Europe ETF {T['dd_live']}",
    f"- 20.11.2026: the fund's NAV fell {pc(NOV20, 2)} in one day {T['n20']} (USD −5.43% on the product page [W1]); cause: !!n/a – not available from public sources or Bloomberg as of 23.11.2026!!",
    "- ^^Daily comparisons are distorted by timing^^: NAV (Warsaw close) and XETRA / other ETF prices are struck at different times – multi-day episodes are used"], 9.5, slide_no=NO["perf_dd_live"])

# drawdowns history
s = d.slide("Performance: drawdowns 2011–2026 (monthly)")
lab = lambda a, z: f"{a[5:]}/{a[2:4]}–{z[5:]}/{z[2:4]}"  # noqa: E731
L.column_chart(s, (0.602, 1.30, 6.3, 3.3), [lab(a, z) for a, z, _ in EPM], [("This fund", tuple(round(f * 100, 1) for _, _, f in EPM)),
                                                                          ("MSCI World ETF", tuple(round((VW[z] / VW[a] - 1) * 100, 1) for a, z, _ in EPM))],
               title="Fund falls of 15% or more (month-end): fund vs MSCI World ETF over the same months, %", size=9, number_format='0.0', axis_format='0', palette=["C01E1E", "AFAFAF"])
L.data_note(s, (0.602, 4.63, 6.3, 0.5), f"Data: Morningstar monthly returns, EUR [M1][M5][M6]; peak month-end to lowest month-end of the fund; others compounded over the same months (own calculation) – Excel [C1] sheet “{s25.ws.title}”.", 8.5)
rows = [["Episode (peak → low)", "Fund", "World", "EM", "Months"]]
for a, z, f in EPM:
    rows.append([lab(a, z), f"^^{pc(f)}^^", pc(VW[z] / VW[a] - 1), pc(VEM[z] / VEM[a] - 1) if a >= "2014-07" else "n/a", str(len(months(a, z)) - 1)])
rows.append(["**Largest fall, whole period**", pc(MDD_F), pc(MDD_W), "–", ""])
y = L.table(s, rows, (7.1, 1.40, 0, 0), (2.0, 0.95, 0.95, 0.95, 0.78), 9.5, bold_cols=(0,), head=True, slide_no=NO["perf_dd_hist"])
L.data_note(s, (7.1, y + 0.03, 5.63, 0.5), f"Data: same series – Excel [C1] sheet “{s25.ws.title}”; Morningstar capture ratios for this fund were not returned by the connector.", 8)
L.body(s, (0.602, 5.2, 12.129, 1.65), [
    f"- **2018–2022: −55%** {T['mdd_h']} – the deepest fall ran for almost five years (COVID, then the 2022 war in Ukraine); since 06/2014 the fund's largest fall was {pc(MDD_FE)} vs {pc(MDD_EM)} for emerging markets {T['mdde']}",
    "- Same pattern in the index history: −77.59% in 2007–2009 [I1 p.1]. A position must be sized so that a 50% fall is acceptable"], 9.5, slide_no=NO["perf_dd_hist"])

# data check
s = d.slide("Performance: check of the data used")
y = L.table(s, [["Period to 31.10.2026, USD", "Fund (factsheet)", "Index (factsheet)", "Difference"],
                ["1 year", "37.17%", "37.10%", "+0.07 pp"], ["3 years a year", "34.47%", "34.29%", "+0.18 pp"], ["5 years a year", "15.63%", "15.63%", "0.00 pp"],
                ["Since 21.01.2011 a year", "3.24%", "3.31%", "−0.07 pp"]],
            (0.602, 1.45, 0, 0), (3.0, 2.0, 2.0, 1.6), 10.5, bold_cols=(0,), head=True, slide_no=NO["perf_check"])
L.data_note(s, (0.602, y + 0.05, 9.0, 0.45), "Data: factsheet cumulative and annualised performance [F1 p.2]; MSCI index factsheet shows the same 1, 3 and 5-year index figures [I1 p.1].", 8.5)
yrs = list(range(2016, 2026))
y2 = L.table(s, [["EUR, own calculation"] + [str(x) for x in yrs], ["Fund NAV, %"] + [f"{comp(FUNDM, months(f'{x}-01', f'{x}-12')) * 100:.1f}" for x in yrs],
                 ["XETRA IBCJ, %"] + [f"{comp(XET, months(f'{x}-01', f'{x}-12')) * 100:.1f}" for x in yrs]],
             (0.602, y + 0.65, 0, 0), (2.1,) + (1.0,) * len(yrs), 9.5, bold_cols=(0,), head=True, slide_no=NO["perf_check"])
L.data_note(s, (0.602, y2 + 0.05, 12.1, 0.45), f"Data: Morningstar monthly total returns in EUR, fund NAV [M1] and XETRA line [M2], compounded – Excel [C1] sheet “{s26.ws.title}”.", 8.5)
L.body(s, (0.602, y2 + 0.6, 12.129, 6.85 - y2 - 0.6), [
    f"- **Morningstar NAV and the XETRA price move together:** largest monthly gap {MAXD:.2f} pp since 07/2011 {T['maxd']} – the EUR line follows the fund",
    "- The factsheet (USD) and the MSCI factsheet give the same index returns, so the tracking figures are consistent [F1 p.2][I1 p.1]",
    "- Own EUR figures differ from the USD factsheet because of the USD/EUR rate – EUR is the relevant currency for us"], 10, slide_no=NO["perf_check"])

# extracts
s = d.slide("Key facts – copies of the official documents")
L.body(s, (0.602, 1.30, 12.1, 0.4), ["Figures used on slides " + ", ".join(str(NO[k]) for k in ("overview", "perf", "track", "strategy")) + " as published (files: Documents ZIP, same keys)."], 11, slide_no=NO["extracts"])
png, sz = L.pdf_crop(LOCAL["F1"], 2, "Share Class", TMPI / "fs_cy.png", pad=(14, 22, 0, 26), width_pt=560)
y_ = L.evidence_picture(s, png, sz, (0.602, 1.85, 6.0, 1.1), "Factsheet USD (Acc), October 2026, calendar year performance [F1 p.2]", slide_no=NO["extracts"])
png, sz = L.pdf_crop(LOCAL["F1"], 2, "CUMULATIVE & ANNUALISED PERFORMANCE", TMPI / "fs_cum.png", pad=(4, 6, 0, 76), width_pt=560)
y_ = L.evidence_picture(s, png, sz, (0.602, y_ + 0.1, 6.0, 1.3), "Factsheet, cumulative and annualised performance to 31.10.2026 [F1 p.2]", slide_no=NO["extracts"])
png, sz = L.pdf_crop(LOCAL["P1"], 89, "MSCI Poland UCITS ETF Investment Objective", TMPI / "p1_pol.png", pad=(60, 8, 0, 250), width_pt=520)
L.evidence_picture(s, png, sz, (0.602, y_ + 0.1, 6.0, 6.75 - y_ - 0.5), "Prospectus, iShares MSCI Poland UCITS ETF – objective and policy [P1 p.89]", slide_no=NO["extracts"])
# A1 p.18 pasuktas 270° (pdftotext ribos – pasukto puslapio pločio 597), todėl iškarpa – tiesiogiai pdftoppm (gulsčias 842 × 595 pt)
import subprocess  # noqa: E402
_x0, _y0, _w, _h, _k = 19, 76, 812, 374, 150 / 72
subprocess.run(["pdftoppm", "-f", "18", "-l", "18", "-r", "150", "-png", "-singlefile", "-x", str(int(_x0 * _k)), "-y", str(int(_y0 * _k)),
                "-W", str(int(_w * _k)), "-H", str(int(_h * _k)), str(LOCAL["A1"]), str(TMPI / "ar_td")], check=True)
png, sz = TMPI / "ar_td.png", (_w, _h)
L.evidence_picture(s, png, sz, (6.75, 1.85, 6.0, 4.5), "Annual report 2025, performance summary, tracking difference and tracking error – last row iShares MSCI Poland UCITS ETF [A1 p.18]", slide_no=NO["extracts"])

# alternatives table
s = d.slide("Possible financial instrument alternatives")
rows = [["Fund", "Type", "Cost", "Size, EUR m", "Benchmark", "Holdings", "Since"]]
TYPES = {"0P0000TJXS": "Physical ETF, Poland", "F00001T1YT": "Physical ETF, Poland (Bulgaria)", "F000013D6B": "Swap ETF, regional", "F00000TJCR": "Physical ETF, all EM"}
for i_, n_, m_ in ALT:
    dp = "UB422" if SNAP[i_].get("UB422") is not None else "OS003"
    rows.append([f"{n_} [{m_}]", TYPES[i_], f"{SNAP[i_][dp]:.2f}%" + ("*" if dp == "OS003" else ""), f"{SNAP[i_]['OF009'] / 1e6:,.1f}",
                 SNAP[i_]["OF00L"], str(SNAP[i_].get("HS008", "–")) if i_ != "F000013D6B" else "swap", SNAP[i_]["OS00F"][:4]])
y = L.table(s, rows, (0.602, 1.40, 0, 0), (3.6, 2.0, 0.85, 1.05, 2.7, 0.9, 0.9), 9.5, bold_cols=(0,), head=True, slide_no=NO["alt1"])
L.data_note(s, (0.602, y + 0.05, 12.1, 0.45), f"Data: Morningstar snapshot, read 23.11.2026 (fund size all classes, EUR) – Excel [C1] sheet “{s27.ws.title}”. *management fee; justETF shows TER 1.38% and EUR 2m [J2]. "
            "Amundi inception: share class 2019 [M4]; fund history from 2005 [J3]. XETRA tickers: IBCJ [W1], PLX [J2], LEER [J3].", 8.5)
L.body(s, (0.602, y + 0.6, 12.129, 6.85 - y - 0.6), [
    "- **Only two Poland ETFs exist:** this fund and Expat Poland WIG20 [J1][J2]. Expat is EUR 2m, TER 1.38%, Bulgarian domicile – not investable for us at size [J2]",
    "- **Amundi MSCI Eastern Europe ex Russia** – cheaper (0.50%), swap-based (unfunded swap), regional index [J3]; its Poland weight: !!n/a in the pages read!!",
    "- **iShares Core MSCI EM IMI** – cheap (0.18%) and diversified, but it is a different exposure (all emerging markets) [M5]",
    "- Direct purchase of the 16 stocks: possible but needs 16 Warsaw trades and rebalancing; the ETF does this at 0.74% with tracking difference close to zero [C1 “" + s7.ws.title + f"” D{r_tda}]"], 10, slide_no=NO["alt1"])

# alternatives performance
s = d.slide("Possible financial instrument alternatives: performance")
YAL = [str(y) for y in YA[:-1]] + ["2026*"]
series = [(n_, tuple(round(comp(s_, months(f"{y}-01", f"{y}-12" if y < 2026 else LM)) * 100, 1) for y in YA)) for n_, _, s_ in ALTP]
L.column_chart(s, (0.602, 1.30, 7.4, 3.6), YAL, series, title="Calendar-year return, EUR, % (*2026: Jan–Oct)", size=8, number_format='0', axis_format='0',
               palette=["C01E1E", "B45F06", "191824", "AFAFAF", "4561A7"])
L.data_note(s, (0.602, 4.92, 7.4, 0.45), f"Data: Morningstar monthly total returns in EUR, compounded per calendar year (own calculation) [M1][M3][M4][M5][M6] – Excel [C1] sheet “{s28.ws.title}”.", 8.5)
rws = [["To 10/2026, a year", "1 year", "3 years", "5 years", "Vol. 5Y"]]
for n_, col, s_ in ALTP:
    rws.append([n_] + [p0(ann(s_, ks)) for ks in (K1Y, K3, K5)] + [p0(vol(s_, K5))])
y = L.table(s, rws, (8.2, 1.40, 0, 0), (1.9, 0.65, 0.7, 0.7, 0.65), 9, bold_cols=(0,), head=True, slide_no=NO["alt2"])
L.data_note(s, (8.2, y + 0.03, 4.6, 0.6), f"Data: same series; volatility = standard deviation of monthly returns × √12 (own calculation) – Excel [C1] sheet “{s28.ws.title}”.", 8)
L.body(s, (0.602, 5.45, 12.129, 1.4), [
    f"- **Amundi Eastern Europe ex Russia moved almost like this fund:** 5 years {p0(A5C)} vs {p0(A5F)} a year {T['cec5']}{T['f5']} – Poland drives the regional index; it is the closest substitute",
    f"- **Expat WIG20** lagged by about 9 pp a year over 5 years ({p0(ann(EXPAT, K5))}) {T['ex5']}; emerging markets and MSCI World were far less volatile"], 10, slide_no=NO["alt2"])

# alternatives monthly heat
s = d.slide("Possible financial instrument alternatives: monthly")
hv = [[round(s_[m], 1) for m in M12] + [round(comp(s_, M12) * 100, 1)] for _, _, s_ in ALTP]
y = L.heatmap(s, (0.602, 1.50, 12.129, 2.6), [n for n, _, _ in ALTP], [m[2:4] + "/" + m[5:] for m in M12] + ["12M"], hv, size=9, scale=6.0, label_w=2.4, slide_no=NO["alt_heat"])
L.data_note(s, (0.602, y + 0.05, 12.1, 0.45), f"Data: Morningstar monthly total returns in EUR, %, 11/2025–10/2026 [M1][M3][M4][M5][M6]; 12M = compounded (own calculation) – Excel [C1] sheet “{s29.ws.title}”. "
            "Colour: green = gain, red = loss in the month.", 8.5)
diff = [round(FUNDM[m] - CEC[m], 1) for m in M12]
y2 = L.heatmap(s, (0.602, y + 0.65, 12.129, 0.7), ["This fund minus Amundi E. Europe, pp"], [m[2:4] + "/" + m[5:] for m in M12], [diff], size=9, scale=2.0, label_w=2.4, slide_no=NO["alt_heat"])
L.body(s, (0.602, y2 + 0.2, 12.129, 6.85 - y2 - 0.2), [
    f"- **What it shows:** this fund and the Amundi regional ETF differ by at most {max(abs(v) for v in diff):.1f} pp in any month [M1][M4]; Expat WIG20 moves with them but with bigger gaps [M3]",
    "- Emerging markets and MSCI World follow a different path – Poland adds diversification to a global equity portfolio, but with higher risk"], 10.5, slide_no=NO["alt_heat"])

# assessment
s = d.slide("Possible financial instrument alternatives: assessment")
y = L.table(s, [
    ["Criterion", "This fund", "Expat WIG20", "Amundi E. Europe ex Russia", "Core MSCI EM IMI", "Best"],
    ["Exposure: Poland only", "++Yes, 16 stocks++", "Yes, 20 stocks [J2]", "^^Regional^^", "^^All EM^^", "This fund"],
    ["Cost", "0.74%", "^^1.38% [J2]^^", "0.50%", "++0.18%++", "EM IMI (other exposure)"],
    ["Size", "++EUR 936m++", "^^EUR 2m^^", "EUR 728m", "++EUR 36bn++", "This fund (for Poland)"],
    ["Replication", "++Physical++", "Physical", "^^Swap^^", "++Physical++", "This fund, EM IMI"],
    ["5 years a year, EUR", p0(A5F), f"^^{p0(ann(EXPAT, K5))}^^", f"++{p0(A5C)}++", p0(A5E), "Amundi ≈ this fund"],
    ["Volatility 5 years", f"^^{p0(VOL5F)}^^", p0(vol(EXPAT, K5)), p0(vol(CEC, K5)), f"++{p0(VOL5E)}++", "EM IMI"],
    ["XETRA EUR line", "IBCJ [W1]", "PLX [J2]", "LEER [J3]", "–", "–"],
], (0.602, 1.50, 0, 0), (2.3, 1.9, 1.7, 2.3, 1.9, 1.929), 10, bold_cols=(0,), head=True, shade_col=1, slide_no=NO["alt3"])
L.data_note(s, (0.602, y + 0.08, 12.1, 0.45), f"Funds: Expat Poland WIG20 UCITS ETF, BGPLWIG04173 [M3][J2]; Amundi MSCI Eastern Europe ex Russia UCITS ETF Acc, LU1900066462 [M4][J3]; iShares Core MSCI EM IMI UCITS ETF, IE00BKM4GZ66 [M5]. "
            f"Data: summary of slides {NO['alt1']}–{NO['alt_heat']}; size: Morningstar EUR, all classes.")
L.heading(s, (0.602, y + 0.62, 12.0, 0), "Final comparison view", 12, AMBER)
L.body(s, (0.602, y + 1.02, 12.129, 6.9 - y - 1.02), [
    "- **Critically:** if the goal is Poland, this fund is the only investable choice – large, physical, tracks the index closely. If the goal is Central and Eastern Europe, the Amundi ETF delivered almost the same return for 0.24 pp less a year, but through a swap. If the goal is emerging markets, a broad EM ETF is far cheaper and less volatile",
    f"- Three alternatives were considered. iShares MSCI Poland UCITS ETF was selected for its pure Polish exposure through physical full replication, a tracking difference of {TD_AVG * 100:+.2f} pp a year {T['tda']} and a fund size of USD 1.1bn [F1 p.1]. **The main risk is the market itself:** a single, concentrated, volatile country that fell 55% in 2018–2022 {T['mdd_h']} – the position size should reflect this."],
    10.5, slide_no=NO["alt3"])

# accounting: PSPKĮ 49
from pspki49 import P49, P49b  # noqa: E402
s = d.slide("Additional information for the fund accounting group", box=(0.602, 0.30, 12.129, 0.664))
L.table(s, P49, (0.535, 0.964, 0, 0), (1.3, 10.0, 0.7), 10.5, bold_rows=(0,), slide_no=NO["acc_law"])
s = d.slide("Additional information for the fund accounting group", box=(0.602, 0.30, 12.129, 0.664))
L.table(s, P49b, (0.535, 0.964, 0, 0), (1.3, 10.0, 0.7), 11, bold_rows=(0,), slide_no=NO["acc_law2"])

s = d.slide("Evidence: PSPKĮ 49.1 and fund documents", box=(0.602, 0.30, 12.129, 0.664))
L.table(s, [
    ["PSPKĮ 49.1", "Reikalavimo raktiniai žodžiai", "Fondo dokumentuose (citata)", "Šaltinis"],
    ["1)", "licencijuotas valstybėje narėje; **Direktyva 2009/65/EB** (UCITS)", ["\"authorised by the Central Bank pursuant to the European Communities (**Undertakings for Collective Investment in Transferable Securities**) Regulations 2011\"",
                                                                          "\"“Directive”, **Directive No. 2009/65/EC** of the European Parliament and of the Council of 13 July 2009\"",
                                                                          "\"authorised by the Central Bank of Ireland as a Undertaking for Collective Investment in Transferable Securities (“UCITS”)\""], "[P1 p.1] [P1 p.10] [K1 p.1]"],
    ["2)", "turto atskyrimas, skolinimasis, skolinimas, turto, kurio neturi, pardavimas", ["\"umbrella investment company … having **segregated liability between its Funds**\"",
                                                                                         "\"may not **borrow**, other than borrowings which in the aggregate do not exceed **10%** of the Net Asset Value of the Fund … on a temporary basis\"",
                                                                                         "\"A Fund may not carry out **uncovered sales** of transferable securities\"; depositary State Street Custodial Services (Ireland) Limited"], "[P1 p.1] [P1 p.231] [P1 p.231] [K1 p.1]"],
    ["3)", "pusmečio ir metų ataskaitos", "\"The **annual report and audited financial statements** will be published within four months … a **semi-annual report and unaudited financial statements** … within two months\"", "[P1 p.207]"],
    ["4)", "ne daugiau kaip **10 %** – kiti KIS", "\"investments made by a Fund in units of other CIS may not exceed, in aggregate, **10%** of the assets of the Fund\" – the Poland fund's policy has no exception", "[P1 p.229] [P1 p.89]"],
    ["UCITS", "Diversifikavimas: UCITS – Yes", "\"UCITS – Yes\" (product page key facts)", "[W1]"],
], (0.535, 0.964, 0, 0), (1.0, 3.0, 6.9, 1.4), 10.5, bold_rows=(0,), slide_no=NO["acc_evid"])

s = d.slide("Additional information for the fund accounting group")
L.table(s, [["BBG tickeris", f"**IBCJ GY Equity** ({OWN}); iShares listing table: Xetra, Bloomberg IBCJ GY [W1]"], ["Tipas", "KIS"], ["Potipis", "KIS 3"],
            ["Įtraukta/neįtraukta į prekybą rinkoje", "Taip – Deutsche Boerse Xetra (IBCJ, EUR), taip pat London, Berne, Mexico [W1][F1 p.3]"],
            ["Perleidžiamieji VP", "Taip"], ["Pinigų rinkos priemonė", "Ne"], ["SFDR", "Article 6 („Other“) [F1 p.1]"]],
        (0.602, 1.45, 0, 0), (4.0, 7.993), 14, bold_cols=(0,), slide_no=NO["acc_bbg"])
L.body(s, (0.602, 5.3, 12.0, 1.3), [L.ucits_bbg_line(ISIN, proven=f"skaidrė {NO['acc_evid']}, [P1 p.10][K1 p.1]")], 14, slide_no=NO["acc_bbg"])

s = d.slide("Additional information for the fund accounting group", box=(0.602, 0.30, 12.129, 0.664))
L.table(s, AF.pptx_rows({
    "Type of issue": ("Investment Fund", "UCITS ETF [K1 p.1]"),
    "Trading Venue": ("Regulated Market", "Deutsche Boerse Xetra [W1]"),
    "Exchange": ("Germany GY", f"Xetra, EUR line IBCJ [W1]; {OWN}"),
    "Ticker": ("IBCJ GY Equity", f"{OWN}; Bloomberg IBCJ GY in the iShares listing table [W1]"),
    "ISIN code": (ISIN, "[F1 p.1][K1 p.1]"),
    "Currency": ("EUR", "currency of the Xetra line [W1]; share class currency – USD [F1 p.1]"),
    "Securitie class": ("Investment Funds", "UCITS ETF shares"),
    "Securitie type": ("Equity Funds", "Polish equities, MSCI Poland [P1 p.89]"),
    "Asset class affiliation": ("Equity", "asset class: Equity [F1 p.1]"),
    "Fund decimal units": ("0", "exchange-traded – whole shares only"),
    "Settlements T+": ("T+2", "EU exchange trades – T+2 (Regulation (EU) No 909/2014, Art. 5(2)); T+1 from 2027-10-11 (European Commission)"),
    "Investment Guidelines Class (High level classification)": ("Equity", ""),
    "Investment Guidelines Subclass (Low level classification)": ("Global Equities", "only equity subclass in the list; Poland is part of MSCI Emerging Markets (“MSCI Emerging - Poland”) [F1 p.1]"),
    "Collective investment undertaking types (investment funds)": ("Equity", "equity fund (KIS 3)"),
}, "en"), (0.602, 1.0, 0, 0), (2.9, 4.2, 5.03), 9.5, bold_cols=(0,), head=True, shade_col=0, slide_no=NO["acc_form"])
L.data_note(s, (1.4, 6.95, 11.3, 0.3), AF.TXT["en"]["lead"] + " " + AF.TXT["en"]["incomplete"])

s = d.slide("Additional information for the fund accounting group", box=(0.602, 0.30, 12.129, 0.664))
Gs = lambda opts, sel: [("++" + o + "++") if o == sel else ("~~" + o + "~~") for o in opts]  # noqa: E731
L.table(s, [
    ["Regionai:", Gs(["Vidurio ir Rytų Europos", "Vakarų Europos", "Japonijos", "Šiaurės Amerikos", "Sektoriniai fondai", "Besivystančios rinkos", "Global markets"], "Vidurio ir Rytų Europos"),
     "Lenkija (Poland) – Vidurio ir Rytų Europa", ""],
    ["KIS Tipai:", Gs(["KIS 1", "KIS 2", "KIS 3", "KIS 4", "KIS 5", "KIS 6", "KIS 7"], "KIS 3"), AF.kis_full_lines("KIS 3"), "~~KIS – kolektyvinio investavimo subjektas; apibrėžimai – Lietuvos banko standartas, pažodžiui~~"],
    ["Turto klasė:", "++Akcijos++", "Priskiriama prie KIS kurie investuoja į akcijas.", ""],
    ["Positioning Report:", "Equity", Gs(["JAV", "Europe", "Japan", "Developed Equities", "Emerging Market Equities"], "Europe"), OWNI],
    ["Diversifikavimas:", "~~KIS Netenkinantys PSKP 49.1~~", "++UCITS – Yes++", ""],
    ["Kita:", "++ETF++", "", ""],
    ["Rizikos lygis:", "++Rizikingos++", "", ""],
], (0.602, 1.0, 0, 0), (1.8, 2.2, 6.6, 1.5), 10, bold_cols=(0,), slide_no=NO["acc_class"])

# method
s = d.slide("Method and data availability")
L.body(s, (0.602, 1.40, 12.129, 5.4), [
    "**Method**",
    "- Facts are taken from the fund documents (prospectus P1, KID K1, factsheet F1, annual report A1), the MSCI index factsheet I1 and saved web pages (W#, J#, N#); every sentence carries its source and page",
    "- Return, risk, ESG and cost data: Morningstar (M#), retrieved 23.11.2026; own calculations in the Excel workbook C1 (one sheet per slide), checked automatically against the slides",
    f"- Internal data: Bloomberg ticker, XETRA bid-ask spread and 90-day average volume ({OWN}); classification for limits ({OWNI}) – Excel [C1] sheet “{ow.ws.title}”",
    "**Not available from public sources or data providers as of 23.11.2026**",
    "- Cause of the one-day fall of −5.43% (USD NAV) on 20.11.2026 – not found in public sources or Bloomberg (internal check, 23.11.2026)",
    "- Named portfolio managers and a named client contact – iShares documents name none; the general Investor Services contact is used [K1 p.3]",
    "**Not covered by the documents read for this DD**",
    "- UN PRI signatory status of the manager; BlackRock voting and engagement record for the 16 Polish companies; Poland weight in the Amundi MSCI Eastern Europe ex Russia ETF",
    "**Not assessed (internal decision, 23.11.2026)**",
    "- Polish dividend withholding-tax rate suffered by the fund vs the rate assumed by the MSCI net index",
    "**Recorded outside this file**",
    f"- Sanctions screening screenshots (2 companies × 3 lists) – taken by the analyst in the internal systems at filing; slides {NO['sanctions']}–{NO['sanctions_2']} hold the boxes"], 11, slide_no=NO["method"])

# sources
s = d.slide("Sources")
items = [(k, v[0], v[1]) for k, v in SOURCES.items()] + [(k, v, URL.get(k)) for k, v in DOCS.items()]
half = (len(items) + 1) // 2
for k, chunk in enumerate((items[:half], items[half:])):
    tb = s.shapes.add_textbox(Inches(0.602 + k * 6.15), Inches(1.50), Inches(6.0), Inches(5.3))
    tf = tb.text_frame; tf.word_wrap = True
    for i, (key, ttl, url) in enumerate(chunk):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        L.M.no_space(p)
        r = p.add_run(); r.text = f"{key}: "; L.M.run_fmt(r, 7.5, True)
        r = p.add_run(); r.text = ttl + f" – {L.where(key, url, None)}"; L.M.run_fmt(r, 7.5)
        if url:
            r.hyperlink.address = url; r.font.color.rgb = L.M.DARK

assert d.n == len(ORDER), (d.n, len(ORDER))
d.save(OUT)
print(OUT, d.n, "slides")
BR = L.bare_refs(OUT)
print("bare refs:", len(BR), BR[:8])
seen = []
for n, t in L.MISSING:
    if (n, t) not in seen:
        seen.append((n, t))
(EVD / "iShares_Poland_not_found.json").write_text(json.dumps(seen, ensure_ascii=False, indent=1))
print(len(seen), "NOT FOUND / patikrinti")
SRC_ALL = dict(SOURCES)
for k, u in URL.items():
    SRC_ALL[k] = (DOCS[k], u)
DOCS_ONLY = {k: v for k, v in DOCS.items() if k not in URL}
L.sources_json(EVD / f"sources_ishares_poland_{ISIN}.json", SRC_ALL, DOCS_ONLY, {}, EXPECT)

# ------------------------------------------------------------------ README (PDF)
_rows = "".join(f"<tr><td><b>{k}</b></td><td>{v[0]}</td></tr>" for k, v in SOURCES.items() if not k.startswith("X")) + \
    "".join(f"<tr><td><b>{k}</b></td><td>{v}</td></tr>" for k, v in DOCS.items())
L.html_pdf(f"""<html><head><meta charset="utf-8"><style>body{{font-family:Arial;font-size:10.5pt;color:#191824}} h1{{font-size:17pt}} h2{{font-size:12.5pt;margin-top:18px}}
td{{vertical-align:top;padding:3px 6px;border-bottom:1px solid #ddd;font-size:9pt}} li{{margin:4px 0}}</style></head><body>
<h1>DD {FUND} USD (Acc) ({ISIN}) – how to check</h1>
<p>DD of {REV}. Prepared by: Analyst. This ZIP holds only PDF and Excel files.</p>
<h2>What is in the ZIP</h2><ul>
<li><b>{XLSX.name}</b> [C1] – every own calculation and every copied number behind a table or chart. One sheet per slide (“S07 Performance” = slide 7), the same table and chart as on the slide; inputs in blue with their source on the row (sheets “Data …”).</li>
<li><b>P1, K1, F1, A1</b> – iShares documents downloaded from ishares.com on 23.11.2026; <b>I1</b> – MSCI index factsheet from msci.com. In the prospectus P1 the PDF page equals the printed page; in the annual report A1 the tag gives the PDF page (the printed number is one lower).</li>
<li><b>W1, J#, N#</b> – web pages saved as PDF on the day of download. Morningstar data [M#] is not a document: it is in the Excel sheets “Data Morningstar …” with SecId, datapoint code, link and retrieval date.</li>
<li>Bloomberg ticker IBCJ GY Equity: {OWN}; the same line is in the iShares listing table [W1].</li>
<li>Internal data (XETRA bid-ask spread {SPREAD:.3f}%, 90-day average volume 252.9k shares – {OWN}; limit classification Europe equity exposure – {OWNI}) are on the Excel sheet “Data internal”.</li>
<li>Data not available, not covered or not assessed are listed on the slide “Method and data availability”.</li></ul>
<h2>How to check a sentence</h2><ol>
<li>Read the tag after the sentence, e.g. <b>[P1 p.89]</b> = prospectus, page 89. Open the file starting with “P1_” and go to page 89.</li>
<li>A tag like <b>[C1 “S07 Performance” D14]</b> = Excel workbook, sheet “S07 Performance”, cell D14; click the cell to see the formula.</li>
<li>Screenshots on the slides are copies of the documents; the caption names the document and page, e.g. [K1 p.3].</li></ol>
<h2>How to check a table or chart</h2><ol>
<li>The “Data:” line under the table names the Excel sheet.</li>
<li>On that sheet the table / chart is the same as on the slide; formulas point to the “Data …” sheets.</li>
<li>On the “Data …” sheets each input row names its source: document and page, or the Morningstar SecId, datapoint and link.</li>
<li>Before delivery every number on these sheets was compared automatically with the number on the slide ({NCHECK} checks, {len(BAD)} differences); no tag without a page or cell ({len(BR)} found).</li></ol>
<h2>Important data notes</h2><ul>
<li>Tracking difference (fund vs MSCI Poland) uses the factsheet in USD, the fund's base currency. All other own return statistics are in EUR (Morningstar).</li>
<li>Emerging markets = iShares Core MSCI EM IMI UCITS ETF; world equities = iShares Core MSCI World UCITS ETF (XETRA, EUR) – investable proxies.</li>
<li>Morningstar ESG values of this fund jump at single month-ends (carbon intensity) – treat them as indicative.</li></ul>
<h2>Source keys</h2><table>{_rows}</table></body></html>""", PKG / "README_how_to_check_iShares_Poland_DD.pdf")
print("README ok")
