# Apskaitos forma „Field Name / Data“ (visi DD ir forma iš laiško)

> Perkelta iš CLAUDE.md 2026-11-20 pažodžiui (savininko sprendimas: kiekvienam darbui – vienas taisyklių failas). Taisyklę keisk tik čia; dashboard užklausos į šį failą tik nurodo.

- **Apskaitos forma „Field Name / Data“** (savininkas 2026-11-11, visi DD – obligacijų docx ir fondų pptx): 14 laukų formos tvarka, pavadinimai pažodžiui, šalia – visas leistinų reikšmių sąrašas (pasirinkta žaliai); laukai, sąrašai ir patikra – `tools/funds_dd/accounting_fields.py` (`docx_table`, `pptx_rows`, `kis_lines`); lauke su sąrašu VISADA reikšmė iš sąrašo (ir CIU obligacijai, pvz. Fixed Income), išimtis – Exchange (sąrašas neišsamus, trūkstama pridedama); DD duoda atsakymą, ne klausimą (netinka tiksliai – artimiausia + komentaras su šaltiniu); Ticker: fondas „XDWC GY Equity“, įmonės obligacija „<ISIN> Corp“, vyriausybės „<ISIN> Gov“; Credit – įmonių, Fixed Income – vyriausybių obligacijos; Settlements T+ – nauja emisija: įkainojimo ir atsiskaitymo datos, antrinė rinka ES T+2 (nuo 2027-10-11 T+1), nelistinguojamas fondas – prospekto terminas; kontekstas pastabose, jei yra ką pasakyti. Kalba: obligacijų DD – DD kalba; fondų DD lentelė – anglų k. (apskaitininkas gali būti ne lietuvis). Fondų DD KIS klasifikacijos lentelėje, KIS eilutėje – pilni KIS 1–7 apibrėžimai pažodžiui (`AF.KIS_FULL`, pasirinktas žaliai). **Apskaitos forma iš laiško** (`#acc`, `accPage`, darbas `acc`; savininkas 2026-11-11 – kai DD daro kitas, o jam reikia formos): įklijuojamas laiškas „please set up…“ → užklausa Claude → Claude užpildo tą pačią 14 laukų formą (`accounting_fields.rows()` patikra), Excel `/mnt/project-files/dd/accounting/` + VIEW {isin, name, table {cols [Field Name, Data, Comment], rows}, check} → `dd_vault.py screen <xlsx> --kind accform --view .. --title "Accounting form <ISIN> <pavadinimas>"`, build, push; rezultatas – tame pačiame puslapyje („Užpildytos formos“, atrakinus archyvą: „Kopijuoti į Excel“, Excel). Kodas reikšmių nespėja (savininkas: visada prašys AI). Pavyzdys – `/mnt/project-files/dd/make_dd_ignitis_lt_v2.py`. Slaptažodžio nereikia ir jis niekur nesaugomas (tik savininko naršyklėje). Savininkas pamiršo slaptažodį → `python tools/dd_vault.py rekey` ir naują slaptažodį perduok tik žinute.

## Ankstesnė pilna dashboard užklausa (iki 2026-11-20, perkelta pažodžiui – visi jos reikalavimai galioja)

Dabar dashboard užklausa trumpa ir nurodo šį failą. Tekstas žemiau – su numatytosiomis įvestimis ([ISIN], [quarter] ir pan.).

### Accounting form from an email

```text
Fill in the accounting form "Field Name / Data" from this email.
Rules – CLAUDE.md "Apskaitos forma" and tools/funds_dd/accounting_fields.py: 14 fields in the form order, values only from the lists (except Exchange), a short comment with source for each; check the terms / prospectus and the Nasdaq or issuer page. If the issuer is already in the tool, use its data.
Result: Excel (Field Name | Data | Comment | Source) in project files /mnt/project-files/dd/accounting/ and `python tools/dd_vault.py screen <xlsx> --kind accform --view <VIEW.json> --title "Accounting form <ISIN> <name>"` (VIEW: {"isin", "name", "table": {"cols": ["Field Name","Data","Comment"], "rows": [...]}}), then build and push. Tell me briefly what to check.

Email:
EMAIL_TEXT
```
