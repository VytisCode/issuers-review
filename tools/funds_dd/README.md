# Funds DD generatoriai (fondai / ETF, PowerPoint)

Perkelta iš Vyčio Claude projekto failų 2026-11-26, kad abu redaktoriai naudotų tą pačią versiją (rules/team_work.md 6 d.).
Taisyklės – [rules/funds_dd.md](../../rules/funds_dd.md) ir [rules/analysis_standards.md](../../rules/analysis_standards.md); šis failas tik paaiškina, kaip paleisti.

## Kas čia
| Failas | Kam |
|---|---|
| `Liquid_Funds_DD_Template.pptx` | Tuščias SEB „Liquid Funds DD Template“ (13 skaidrių, tuščios lentelės). Sukurtas `make_funds_dd.py` pagal aprašymą ir sutikrintas su originalu (2026-11-01). Jame nėra jokio DD turinio. |
| `make_funds_dd.py` | Šablono atkūrimas kode (tema, master, šablono lentelės) – iš jo `funds_dd_lib.Deck` kuria kiekvieną DD. `python make_funds_dd.py` perkuria tuščią šabloną. |
| `funds_dd_lib.py` | Bendri įrankiai: skaidrės, lentelės, žymės (`**bold**`, `!!raudona!!`, `[S#]` nuorodos), diagramos, PDF iškarpos, šaltinių sąrašas, sankcijų skaidrės, README PDF. |
| `calc_book.py` | Skaičiavimų Excel [C1]: lapai „S## …“ kiekvienai skaidrei, „Data …“ su šaltiniais, automatinė patikra Excel = skaidrė. |
| `make_dd_ishares_poland.py` | **Pavyzdys** – indekso ETF DD (iShares MSCI Poland, 2026-11-23, 41 skaidrė). Naujam ETF – kopijuok ir keisk tik fondo blokus. |
| `pspki49.py` | PSPKĮ 49 str. lentelės apskaitos skaidrėms (anksčiau buvo `make_dd_xdwc.py` viduje). |
| `accounting_fields.py` | Apskaitos formos „Field Name / Data“ 14 laukų ir leistinos reikšmės (bendra obligacijų ir fondų DD). |
| `apskaita_fondams.md`, `standartiniai_tekstai.md` | Apskaitos skaidrių tekstai ir standartiniai sakiniai (pažodžiui, kaip pateikė Vytis). |
| `struktura_etf_dd.md` | DD struktūra ir metodika (v4.2 → v3; viršesnė naujesnė). |

Aktyvaus fondo pavyzdys (`make_dd_aqr_delphi.py`, `aqr_calc_book.py`) ir kiti senesni DD generatoriai kol kas liko tik Vyčio projekto failuose.

## Kas NĖRA repozitorijoje (ir neturi būti)
- **Fondo dokumentai, Morningstar duomenys, sugeneruoti DD** (`.pptx`, Excel, ZIP) – tik darbo aplanke (Claude projekto failuose) ir užšifruotame archyve (`python tools/dd_vault.py add …`).
- **SEB vidinis eNameChecker adresas.** Generatorius jį ima iš aplinkos kintamojo `ENAMECHECKER_URL`; jei jo nėra, skaidrėje lieka pavadinimas be nuorodos. Adresą kolegos perduoda asmeniškai.

## Kaip paleisti
Reikia: Python 3 su `python-pptx`, `lxml`, `Pillow`, `openpyxl` (`pip install python-pptx lxml Pillow openpyxl`), LibreOffice (`soffice`, Excel perskaičiavimui), Poppler (`pdftotext`, `pdftoppm`), Chromium (README PDF; Claude Code debesyje jau yra).

```bash
# darbo aplankas: dd_evidence/<fondas>/docs (PDF), Morningstar json; rezultatas rašomas ten pat
export FUNDS_DD_WORK=/mnt/project-files/funds_dd      # numatyta; kitur – savo kelias
export ENAMECHECKER_URL=...                           # nebūtina; tik savo aplinkoje
python3 tools/funds_dd/make_dd_ishares_poland.py
```
Pavyzdys tikrina pats save: „277 checks, 0 differences“, „bare refs: 0“, „README ok“. Jo duomenys (`dd_evidence/iShares_MSCI_Poland_IE00B4M7GH52/`) yra tik Vyčio projekto failuose, todėl kolegos projekte jis kaip pavyzdys skaitomas, bet nepaleidžiamas.

**Naujas DD:** nukopijuok `make_dd_ishares_poland.py` į savo darbo aplanką (ne į repozitoriją – DD turinys nėra viešas), pakeisk eilutę `sys.path.insert(0, str(HERE))` į `sys.path.insert(0, "<repo kelias>/tools/funds_dd")` ir keisk fondo blokus. Bendrų įrankių pataisymus (`funds_dd_lib.py`, `calc_book.py` ir kt.) daryk čia, per PR (rules/team_work.md).
