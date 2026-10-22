# Emitentų apžvalga — instrukcijos Claude

Įrankis turimų obligacijų / akcijų emitentų apžvalgai (Baltijos šalių emitentai, ~20 emitentų, ~40 ISIN). Savininkas Vytis – ne programuotojas, bendrauja lietuviškai.
Viešas repo ir svetainė: tik vieši duomenys, jokių slaptažodžių ar asmeninių duomenų. Jei užduotis prieštarauja taisyklėms – sustok ir paklausk.

## Taisyklės
1. **Viskas nemokama.** Jei kažkas kainuotų – nedaryk, paaiškink ir lauk patvirtinimo. AI naudojamas tik paprašius (mygtukas ar žinutė), automatinis rinkimas – be AI.
2. **Faktai, ne interpretacijos.** Kiekvienas skaičius ir teiginys – su šaltiniu `[S# p.N]`. Nerasta → rašyk „nerasta“ ir žymėk `review`. Niekada nepildyk iš atminties.
3. **Skaičiai – kaip skelbia emitentas.** Paties apskaičiuoti rodikliai žymimi `verified: calc` ir tekste „apskaičiuota įrankio“. Iš santraukos, ne iš originalo teksto, paimti – `verified: check`.
4. **Istorija neperrašoma:** `data/history/*.csv` tik papildoma (taisyti galima tik aiškią klaidą). Senos ataskaitos `reports/*/*.yaml` netrinamos – nauja analizė = naujas failas.
5. **Kalbos:** `config/settings.yaml` `languages` (savininko sprendimas 2026-10-22: kol kas tik `lt`). Tekstus rašyk bent `lt`; `en`/`lv`/`et` – kai bus įjungtos.
6. Sankcijų ir žiniasklaidos patikras daro savininkas – įrankis tik pateikia radinius (`found` / `not_found` / `review`) su nuorodomis.

## Failų žemėlapis
- `config/issuers.yaml` – emitentai (šalis, sektorius, ataskaitų puslapiai, žiniasklaidos kalbos, SEB politikos); `config/securities.csv` – ISIN, kuponai, datos (`status` matured skaičiuojamas automatiškai).
- `config/seb_policies.yaml` – SEB politikų santrauka (`verified: no` – SEB PDF nepavyko perskaityti automatiškai) ir ESG klausimynas.
- `config/metrics.yaml`, `config/i18n.yaml` – rodiklių ir antraščių vertimai.
- `reports/<emitentas>/<laikotarpis>.yaml` – ataskaitos turinys; `data/history/<emitentas>.csv` – rodikliai (metric, period, value, unit, source, page, verified, note).
- `tools/build_reports.py` (Word/PDF → `docs/files/`), `tools/archive_sources.py` (šaltinių ZIP → GitHub Release `sources`), `tools/build_site.py` (`docs/data.js`, kopijuoja `site/` → `docs/`), `tools/fetch_news.py`, `tools/check_filings.py`, `tools/fetch_market.py` (ECB palūkanos, infliacija → `data/market.csv`), `tools/fetch_prices.py` (Nasdaq Baltic obligacijų kotiruotės → `data/prices.csv`), `tools/common.py`. PDF šriftas – Inter (`assets/fonts`), Word – Calibri.
- `site/` – dashboard šaltinis (`index.html`, `app.js`, `style.css`); `docs/` – generuojama, ranka netaisyti.

## Kaip atlikti analizę (nauja emitento ataskaita)
1. Rask naujausią ataskaitą (emitento IR puslapis iš `config/issuers.yaml`, `data/filings.csv` nauji įrašai). Perskaityk ataskaitą, pranešimą spaudai, prezentaciją; reitingų puslapį.
2. Papildyk `data/history/<emitentas>.csv`: naujo laikotarpio rodikliai + palyginamasis laikotarpis, su šaltiniu ir puslapiu.
3. Nukopijuok ankstesnį `reports/<emitentas>/*.yaml` kaip pagrindą → naujas failas `<laikotarpis>.yaml` (pvz. `9M2026`). Atnaujink: `analysis_date`, `period`, `periods`, `sources`, visas skiltis.
4. **Planai ir rezultatai:** perkelk ankstesnius planus, atnaujink `outcome` ir `status` (met / partly / not_met / pending / changed), pridėk naujus emitento pažadus (prognozės, strategijos tikslai, dividendų politika).
5. **Investuotojui:** susiek įmonės būklę su turimais VP (`config/securities.csv` held=yes): rangas, reitingas, išpirkimai, call/tender, sverto ribos, dividendai.
6. **ESG:** kiekvienam klausimyno punktui `status` + trumpas įrodymas + šaltinis pagal emitento `seb_policies`.
7. **Patikros:** peržiūrėk `data/news.csv` (flag=review) ir greitą paiešką emitento šalies kalba; `checks.media` – `found` su nuorodomis arba `not_found`. `checks.sanctions` visada `review`.
   **Spauda:** `press:` – 1–4 punktai (lt/en), ką rašė žiniasklaida per ~90 d. (svarbiausia investuotojui: sandoriai, obligacijos, vadovai, priežiūros institucijos), šaltiniai `N#` iš `data/news.csv` (`common.news_digest`). ESG `media` įrodymas turi atitikti šią skiltį (automatinė suvestinė pridedama pati).
   **Rinka ir aplinka:** `market:` – 2–4 punktai (lt/en), kas iš aplinkos lėmė rodiklių pokyčius ar veikia VP (palūkanos, kainos, paklausa, reguliavimas), tik kaip rašo emitentas ar žiniasklaida, su šaltiniu. ECB lentelė pridedama automatiškai iš `data/market.csv`.
   **Emisijų reitingai:** `config/securities.csv` `issue_rating` / `issue_rating_note` / `issue_rating_source` – tik jei emitento ar agentūros puslapis nurodo emisijos (klasės) reitingą; kitaip tuščia (rodoma „atskiro nerasta“).
8. Paleisk `python tools/build_reports.py` ir `python tools/build_site.py`; patikrink, kad be klaidų.
9. Commit ir push į `main`. Savininkui – trumpai, kas pasikeitė nuo ankstesnės analizės.

## Patikra prieš push
`python tools/build_reports.py && python tools/build_site.py` be klaidų; jei keistas `site/` – atidaryti `docs/index.html` naršyklėje (Playwright) be JS klaidų.
