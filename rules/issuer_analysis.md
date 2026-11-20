# Emitento analizė (nauja ataskaita)

> Perkelta iš CLAUDE.md 2026-11-20 pažodžiui (savininko sprendimas: kiekvienam darbui – vienas taisyklių failas). Taisyklę keisk tik čia; dashboard užklausos į šį failą tik nurodo.

1. Rask naujausią ataskaitą (emitento IR puslapis iš `config/issuers.yaml`, `data/filings.csv` nauji įrašai). Emitentams su `manual_reports: true` (Nortal) ataskaitas įkelia savininkas – jei jų nėra, paprašyk Vytį jas pateikti, iš atminties nepildyk. Perskaityk ataskaitą, pranešimą spaudai, prezentaciją; reitingų puslapį.
2. Papildyk `data/history/<emitentas>.csv`: naujo laikotarpio rodikliai + palyginamasis laikotarpis, su šaltiniu ir puslapiu.
3. Nukopijuok ankstesnį `reports/<emitentas>/*.yaml` kaip pagrindą → naujas failas `<laikotarpis>.yaml` (pvz. `9M2026`). Atnaujink: `analysis_date`, `period`, `periods`, `sources`, visas skiltis.
4. **Planai ir rezultatai:** perkelk ankstesnius planus, atnaujink `outcome` ir `status` (met / partly / not_met / pending / changed), pridėk naujus emitento pažadus (prognozės, strategijos tikslai, dividendų politika).
5. **Investuotojui:** susiek įmonės būklę su turimais VP (`config/securities.csv` held=yes): rangas, reitingas, išpirkimai, call/tender, sverto ribos, dividendai. **Skolos sąlygos:** `covenants:` (lt/en…) – finansiniai rodikliai (riba ir faktinė reikšmė), kitos svarbios sąlygos (neigiamas įkeitimas, kontrolės pasikeitimas, užtikrinimas), ar vykdomos – tik jei yra ką aptarti, su šaltiniu; tas pats ir DD.
6. **ESG:** kiekvienam klausimyno punktui `status` + trumpas įrodymas + šaltinis pagal emitento `seb_policies`.
7. **Patikros:** peržiūrėk `data/news.csv` (flag=review) ir greitą paiešką emitento šalies kalba; `checks.media` – `found` su nuorodomis arba `not_found`. `checks.sanctions` visada `review`. Ataskaitose prie sankcijų punkto automatiškai pridedamos OFAC ir ES sankcijų žemėlapio nuorodos (`common.SANCTION_SRC`); SEB vidinė eNameChecker nuoroda viešai neskelbiama.
   **Spauda:** `press:` – 1–4 punktai (lt/en), ką rašė žiniasklaida per ~90 d. (svarbiausia investuotojui: sandoriai, obligacijos, vadovai, priežiūros institucijos), šaltiniai `N#` iš `data/news.csv` (`common.news_digest`). ESG `media` įrodymas turi atitikti šią skiltį (automatinė suvestinė pridedama pati).
   **Rinka ir aplinka:** `market:` – 2–4 punktai (lt/en), kas iš aplinkos lėmė rodiklių pokyčius ar veikia VP (palūkanos, kainos, paklausa, reguliavimas), tik kaip rašo emitentas ar žiniasklaida, su šaltiniu. ECB lentelė pridedama automatiškai iš `data/market.csv`.
   **Emisijų reitingai:** `config/securities.csv` `issue_rating` / `issue_rating_note` / `issue_rating_source` – tik jei emitento ar agentūros puslapis nurodo emisijos (klasės) reitingą; kitaip tuščia (rodoma „atskiro nerasta“).
8. Paleisk `python tools/build_reports.py` ir `python tools/build_site.py`; patikrink, kad be klaidų.
9. Commit ir push į `main`. Savininkui – trumpai, kas pasikeitė nuo ankstesnės analizės.
