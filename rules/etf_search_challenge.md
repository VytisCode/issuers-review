# ETF paieška ir turimų pozicijų peržiūra (challenge)

> Perkelta iš CLAUDE.md 2026-11-20 pažodžiui (savininko sprendimas: kiekvienam darbui – vienas taisyklių failas). Taisyklę keisk tik čia; dashboard užklausos į šį failą tik nurodo.

- **ETF paieška** (prieš DD, 2026-10-28): DD puslapio režimas „ETF paieška“ (`#dd/screen`, `screenPage`, darbas `screen`) – kriterijų gidas + laisvas komentaras → užklausa. Claude suranda visas tinkamas klases (justETF profiliai, valdytojų faktų lapai/KID; SFDR – tik kaip skelbia valdytojas, kitaip 6), palygina ir siūlo finalistus (kritiškai, kaip DD alternatyvose). Rezultatas – PowerPoint anglų k. (`/mnt/project-files/funds_dd/screens/make_screen_*.py`, pavyzdys S&P 500 Equal Weight) + `VIEW.json` → `python tools/dd_vault.py screen <failas.pptx> --view <VIEW.json> --title ..` (užšifruota; dashboard'e „Atliktos ETF paieškos ir peržiūros“ – peržiūra, PowerPoint, mygtukas „Pilnas DD“ užpildo Funds DD formą). **Visada su įrodymais** (savininkas 2026-10-28): `--sources <šaltiniai.json>` (dokumentų ZIP per update.yml) ir `--evidence <katalogas>` (duomenų išrašai, pvz. pension-fund-tracker pozicijos, savi skaičiavimai CSV, generatorius); esamam įrašui – `dd_vault.py sources <id> ..` / `evidence <id> ..`. Pavyzdys – `screens/evidence_sp500ew/`.
- **Turimų pozicijų peržiūra (challenge)** (2026-10-28): DD puslapio režimas (`#dd/challenge`, `chalPage`, darbas `chal`) – pozicijos ranka (bet kuri šalis) arba iš LT pensijų fondų portfelių (pension-fund-tracker `docs/data_pf.js`, kraunama naršyklėje; valdytojas, pakopa, fondas, kategorija, regionas, SFDR, ketvirtis), mandato taisyklės (PAB…), valdytojų pasiūlymai, rinkos kontekstas → užklausa. Metodika – `/mnt/project-files/funds_dd/challenge/metodika.md` (verdiktas SWITCH / KEEP / WATCH: vartai, grynoji nauda ≥ 0,05 pp ir atsipirkimas ≤ 2 m., strategijos variantai, kolegų signalai, rinkos kontekstas – faktai ir scenarijai be prognozių). Rezultatas – PowerPoint anglų k. → `dd_vault.py screen <pptx> --view <VIEW.json> --title .. --kind challenge` (VIEW `table` {cols, rows, hl}). Dokumentų stilius – dalykinis, be asmeninių frazių.

## Ankstesnė pilna dashboard užklausa (iki 2026-11-20, perkelta pažodžiui – visi jos reikalavimai galioja)

Dabar dashboard užklausa trumpa ir nurodo šį failą. Tekstas žemiau – su numatytosiomis įvestimis ([ISIN], [quarter] ir pan.).

### Challenge

```text
Task: Challenge review.
Funds:
[ISINs, or a description such as “all US equity funds we hold”]
Do we hold them: Yes – held: keep or switch?
Method: /mnt/project-files/funds_dd/challenge/metodika.md – held: SWITCH / KEEP / WATCH (hard gates, net advantage ≥ 0.05 pp a year, break-even ≤ 2 years); candidate: ADD / DO NOT ADD / WATCH against the closest held exposure (section 8a; a short or backtested record caps the verdict at WATCH); peers from pension-fund-tracker; market context as facts and scenarios. Latest example: /mnt/project-files/funds_dd/challenge/aqr_delphi_fusion/.
If the list is a description, take the held funds from the latest quarterly review in the archive (say which you used) and also look for new market leaders with the same exposure.
Sources: the manager's own documents first (factsheet, KID, prospectus at one common date); justETF only for price series and cross-checked; reference fund with ISIN in every comparison; every analytical statement backed by a table or chart.
Output: PowerPoint in English + `python tools/dd_vault.py screen <pptx> --view <VIEW.json> --kind challenge --title "Challenge: …" --sources <sources.json> --evidence <folder>`. VIEW: table (Fund, ISIN, Verdict, …; hl = verdict column), picks = funds worth a DD, plan = [{date, what, isin, task}] only for a date I asked for; never schedule separate re-reviews (e.g. “after 12 months live”) – such milestones are checked in the quarterly fund review if the fund is held – the dashboard shows these on the fund cards and in the calendar. Not in the public repo.
Report back in Lithuanian: verdicts and the next step.
Missing data (VIEW.missing, at most 15, only what is not public): each item must tell me exactly what to look up – what = the exact figure as the source names it, with the share class and the date or period (e.g. “Effective duration of class I EUR Acc at 31 Oct 2026”); where = the exact place step by step (system → screen → field, or the manager page → document → section); url = the direct link to that page or document whenever one exists (checked to open, never guessed); search = words for a web search with the ISIN (e.g. “factsheet October 2026”); how = one practical tip; fmt = an example answer; why = what it changes.
```

### Find ETFs → challenge → DD

```text
Task: Find ETFs → challenge → DD. Three steps; STOP after step 2 and wait for my OK.
Looking for: [index, sector, strategy or theme]
Must have: UCITS, listed on XETRA (GY), EUR trading line
Step 1 – search: every UCITS share class that fits (justETF profiles, manager factsheets and KIDs; SFDR only as the manager states it, otherwise 6). Explain each index rulebook (selection, exclusions, weighting, rebalancing), holdings and sector differences, price history (growth, drawdowns, rolling 12 months), size trend, peers (pension-fund-tracker); weighted scorecard with the weights shown, a sensitivity line and a blunt one-line verdict per fund.
Step 2 – challenge every finalist against what we hold with the same exposure (metodika.md) and name one winner.
Archive steps 1–2 as one entry: `python tools/dd_vault.py screen <pptx> --view <VIEW.json> --kind screen --title "Find ETFs: …" --sources <sources.json> --evidence <folder>`; VIEW picks = finalists, winner first. Then report back in Lithuanian: winner, runner-up and why.
Step 3 – only after my OK: the Funds DD for the winner (rules as in the Due diligence task), based on this entry.
Missing data (VIEW.missing, at most 15, only what is not public): each item must tell me exactly what to look up – what = the exact figure as the source names it, with the share class and the date or period (e.g. “Effective duration of class I EUR Acc at 31 Oct 2026”); where = the exact place step by step (system → screen → field, or the manager page → document → section); url = the direct link to that page or document whenever one exists (checked to open, never guessed); search = words for a web search with the ISIN (e.g. “factsheet October 2026”); how = one practical tip; fmt = an example answer; why = what it changes.
```
