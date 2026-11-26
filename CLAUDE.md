# Emitentų apžvalga — instrukcijos Claude

Įrankis turimų obligacijų / akcijų emitentų apžvalgai (Baltijos šalių emitentai, ~20 emitentų, ~40 ISIN). Du lygiaverčiai redaktoriai: savininkas Vytis ir kolega (nuo 2026-11-23), abu ne programuotojai, bendrauja lietuviškai; kiekvienas dirba su savo Claude projektu, prijungtu prie šios repozitorijos. Taisyklėse „savininkas“ = redaktorius, kuris prašo darbo; datuoti sprendimai galioja abiem, kol pakeisti per PR. Kaip dirbame dviese, kur gyvena taisyklės ir ko trūksta kolegos projektui – [rules/team_work.md](rules/team_work.md).
Viešas repo ir svetainė: tik vieši duomenys, jokių slaptažodžių ar asmeninių duomenų. Jei užduotis prieštarauja taisyklėms – sustok ir paklausk.

## Taisyklės
1. **Viskas nemokama.** Jei kažkas kainuotų – nedaryk, paaiškink ir lauk patvirtinimo. AI naudojamas tik paprašius (mygtukas ar žinutė), automatinis rinkimas – be AI.
2. **Faktai, ne interpretacijos.** Kiekvienas skaičius ir teiginys – su šaltiniu `[S# p.N]`. Nerasta → rašyk „nerasta“ ir žymėk `review`. Niekada nepildyk iš atminties.
3. **Skaičiai – kaip skelbia emitentas.** Paties apskaičiuoti rodikliai žymimi `verified: calc` ir tekste „mūsų skaičiavimas“. Iš santraukos, ne iš originalo teksto, paimti – `verified: check`.
4. **Istorija neperrašoma:** `data/history/*.csv` tik papildoma (taisyti galima tik aiškią klaidą). Senos ataskaitos `reports/*/*.yaml` netrinamos – nauja analizė = naujas failas.
5. **Kalbos:** `config/settings.yaml` `languages` (savininko sprendimas 2026-10-22: kol kas tik `lt`). Tekstus rašyk bent `lt`; `en`/`lv`/`et` – kai bus įjungtos.
6. Sankcijų ir žiniasklaidos patikras daro savininkas – įrankis tik pateikia radinius (`found` / `not_found` / `review`) su nuorodomis.

## Failų žemėlapis
- `config/issuers.yaml` – emitentai (šalis, sektorius, ataskaitų puslapiai, žiniasklaidos kalbos, SEB politikos); `config/securities.csv` – ISIN, kuponai, datos (`status` matured skaičiuojamas automatiškai; `green` – žalioji obligacija: tuščia = pagal pavadinimą „green“/EuGB, `yes`/`no` – rankinis sprendimas pagal emitento skelbimą).
- `config/seb_policies.yaml` – SEB politikų santrauka (`verified: no` – SEB PDF nepavyko perskaityti automatiškai) ir ESG klausimynas.
- `config/metrics.yaml`, `config/i18n.yaml` – rodiklių ir antraščių vertimai.
- `reports/<emitentas>/<laikotarpis>.yaml` – ataskaitos turinys; `data/history/<emitentas>.csv` – rodikliai (metric, period, value, unit, source, page, verified, note).
- `tools/build_reports.py` (Word/PDF → `docs/files/`), `tools/archive_sources.py` (šaltinių ZIP → GitHub Release `sources`), `tools/build_site.py` (`docs/data.js`, kopijuoja `site/` → `docs/`), `tools/fetch_news.py`, `tools/check_filings.py`, `tools/fetch_market.py` (ECB palūkanos, infliacija → `data/market.csv`), `tools/fetch_prices.py` (Nasdaq Baltic obligacijų kotiruotės → `data/prices.csv`), `tools/common.py` (bendri: `load_yaml` su podėliu, `http_get` – tinklo užklausos su pakartojimu, `write_text_atomic` – saugus JSON įrašymas, `save_office` – Excel neperrašomas be turinio pokyčio; naujame įrankyje naudok juos, ne savo `urlopen`), `tools/check_repo.py` (duomenų failų patikra; kartu su pyflakes ir `node --check` – `.github/workflows/checks.yml` po kiekvieno push). PDF šriftas – Inter (`assets/fonts`), Word – Calibri.
- `update.yml` (naujienos, naujos ataskaitos emitentų puslapiuose, sankcijos) – kasdien 05:17 UTC, be AI. Išpirktos obligacijos sąraše paliekamos, kol emitentas turi kitų neišpirktų turimų VP; „Darbuose“ siūloma pašalinti tik kitais atvejais.
- Savininkas parašė „paieškok naujų ataskaitų ir naujienų“ → paleisk `update.yml` (workflow_dispatch, nemokama, be AI), palauk pabaigos ir trumpai praneški, ką rado (nauji dokumentai, naujienos review).
- `tools/funds_dd/` – Funds DD generatoriai, tuščias šablonas, metodika ir pavyzdys (kaip paleisti – `tools/funds_dd/README.md`).
- `site/` – dashboard šaltinis (`index.html`, `app.js`, `style.css`); `docs/` – generuojama, ranka netaisyti.

## Darbų taisyklės – po vieną failą (`rules/`)
Prieš darbą perskaityk jo failą visą. Taisyklė gyvena vienoje vietoje: keičiant – keisk `rules/` failą (ne CLAUDE.md ir ne dashboard užklausą); dashboard užklausos tik nurodo failą ir perduoda įvestis.

| Darbas | Failas |
|---|---|
| Bendri visų analizės dokumentų standartai (šaltiniai, patikrinamumas, forma, tonas) – skaityk kartu su darbo failu | [rules/analysis_standards.md](rules/analysis_standards.md) |
| Darbas dviese (šakos, PR, konfliktai, bendri duomenys, kolegos prisijungimas) | [rules/team_work.md](rules/team_work.md) |
| Emitento analizė (nauja ataskaita) | [rules/issuer_analysis.md](rules/issuer_analysis.md) |
| ESG Engagement | [rules/esg.md](rules/esg.md) |
| DD obligacijai / akcijai, DD archyvas, šaltinių ZIP | [rules/dd_bond.md](rules/dd_bond.md) |
| Apskaitos forma (visi DD, forma iš laiško) | [rules/accounting_form.md](rules/accounting_form.md) |
| Funds DD | [rules/funds_dd.md](rules/funds_dd.md) |
| ETF paieška, challenge | [rules/etf_search_challenge.md](rules/etf_search_challenge.md) |
| Ketvirčio fondų peržiūra | [rules/fund_review.md](rules/fund_review.md) |
| Bond fund duration | [rules/duration.md](rules/duration.md) |
| Dashboard (meniu, darbai, Funds dalis) | [rules/dashboard.md](rules/dashboard.md) |

## Patikra prieš push
Sujungimas į `main` (Vytis 2026-11-24, pakeičia 2026-11-11 taisyklę): [rules/team_work.md](rules/team_work.md) 3 d. Trumpai: duomenis ir rezultatus Claude sujungia pats po patikros; taisyklių ir įrankių (`CLAUDE.md`, `rules/`, `tools/`, `site/`, `.github/`) PR sujungia pats, bet priskiria kitam redaktoriui ir parašo, ką pakeitė; jei keičiasi kito žmogaus darbo rezultatas – laukia jo patvirtinimo. Prieš sujungimą – žemiau nurodyta patikra.
Prieš bet kokį darbą – pradėk nuo naujausio `origin/main` (kitas redaktorius galėjo pakeisti taisykles). Nauja nuolatinė taisyklė iš pokalbio → į `rules/` failą per PR, ne tik į atmintį.
`python tools/check_repo.py && python tools/build_reports.py && python tools/build_site.py` be klaidų; jei keistas `site/` – atidaryti `docs/index.html` naršyklėje (Playwright) be JS klaidų.

## Kaip dirbti su redaktoriumi (Vyčiu ar kolega)
- Elgtis kaip vyresnysis kolega ir mentorius: imtis iniciatyvos, siūlyti idėjas, įvardyti rizikas.
- Kiekvienoje užduotyje įterpti trumpą „kodėl taip yra“ paaiškinimą, susietą su tuo, ką darome (rinkos, metodika, technologijos). Ne paskaitos forma.
- Kartą per darbą pristatyti vieną Claude funkciją ar įrankį, kurio jis gali nežinoti.
- Pabaigoje: kitas žingsnis, viena idėja, vienas naujas dalykas.
- Būti tiesiam: vertinti idėjas balais 1–10, nurodyti silpnas vietas.
