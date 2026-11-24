# Darbas dviese (du lygiaverčiai redaktoriai)

> Sukurta 2026-11-23 (Vytis: „kad mums dirbant abiems, mes abu galėtume tobulinti ir naudotis vienas kito patobulinimais“). Taisyklę keisk tik čia.

## 1. Kur gyvena žinios
- **Vienintelis bendras šaltinis – ši repozitorija** (`main` šaka): `CLAUDE.md` (rodyklė), `rules/*.md` (kiekvieno darbo taisyklės), `rules/analysis_standards.md` (bendri visų analizės dokumentų standartai), `tools/`, `site/`, `config/`.
- Kiekvieno redaktoriaus Claude projektas turi **savo** pokalbius, atmintį ir projekto failus. Kito redaktoriaus Claude jų **nemato**. Todėl taisyklė, kurią pasakei tik pokalbyje ar kuri liko tik projekto atmintyje, kolegai neegzistuoja.
- Claude, gavęs naują nuolatinę taisyklę („nuo šiol…“, „visada…“), įrašo ją į atitinkamą `rules/` failą ir sujungia į `main` (žr. 3 d.). Į projekto atmintį – tik tai, kas asmeniška arba laikina (pvz., „laukiu kolegos atsakymo“).
- Paskyros „skills“ (claude.ai) yra asmeniniai: kolega jų negauna ir negali redaguoti. Todėl bendras standartas laikomas `rules/analysis_standards.md`, o `.claude/skills/analysis-standards/SKILL.md` tik nurodo į jį (Claude Code jį randa repozitorijoje).
- Repozitorija **vieša**: jokių slaptažodžių, raktų, asmens duomenų, SEB vidinių dokumentų, portfelio sumų ar DD turinio atviru tekstu.

## 2. Kaip į repozitoriją patenka pakeitimas
1. Kiekviena Claude gija dirba **savo šakoje** (`claude/...`), niekada tiesiai kito žmogaus šakoje.
2. Prieš pradėdamas Claude pasiima naujausią `main` (`git fetch origin main` ir darbą pradeda nuo jo), kad naudotų kolegos naujausias taisykles.
3. Pakeitimas → pull request (PR) į `main` su trumpu aprašu lietuviškai: ką keičia ir kodėl.
4. Patikra (CLAUDE.md „Patikra prieš push“) + GitHub „Checks“ žali.
5. Sujungimas pagal 3 dalies taisyklę.

## 3. Kas sujungia į `main`
**Galioja nuo 2026-11-24** (Vytis pasirinko šią tvarką; pakeičia 2026-11-11 taisyklę „Claude viską sujungia pats“):
- *Duomenys ir rezultatai* (`reports/`, `data/`, ataskaitos, DD archyvo įrašai, `config/securities.csv`): Claude sujungia pats po patikros, kaip dabar.
- *Taisyklės ir įrankiai* (`CLAUDE.md`, `rules/`, `tools/`, `site/`, `.github/`): Claude sujungia pats po patikros, bet PR priskiria **kitam** redaktoriui (assignee) ir pranešime parašo vieną sakinį „ką pakeitė kitam“. Jei pakeitimas keičia kito žmogaus darbo rezultatą (pvz., DD formą, Duration lentelę) – nesujungia, kol kitas redaktorius nepatvirtina PR.
- Kodėl ne „visada laukti patvirtinimo“: abu dirbate ne kasdien, o laukiantis PR genda (pasensta, atsiranda konfliktų). Kodėl ne „niekada nelaukti“: kitas žmogus kitą rytą gautų pasikeitusį DD formatą be įspėjimo.

## 4. Kaip išvengti konfliktų
- `docs/` – generuojama. Konflikto `docs/` failuose netaisyti ranka: paimti `main` versiją ir paleisti `python tools/build_reports.py && python tools/build_site.py` iš naujo.
- Automatiniai darbai (`update.yml`, `isin-sync.yml`, `duration.yml`, `fund-docs.yml`) rašo tiesiai į `main` kasdien. Todėl šaka prieš sujungimą visada atnaujinama nuo `main`.
- Tą pačią dieną tas pats `rules/` failas – tik vienoje gijoje. Jei abu keičiate tą patį darbą (pvz., Funds DD), susitarkite, kas jį keičia, arba antras PR daromas po pirmojo sujungimo.
- Istorija neperrašoma (CLAUDE.md 4 taisyklė): nauja analizė = naujas failas, todėl du žmonės to paties emitento ataskaitos neperrašo.

## 5. Bendri ir atskiri duomenys
| Kas | Bendra ar atskira | Pastaba |
|---|---|---|
| Emitentų ataskaitos, istorija, naujienos, ESG, dashboard | Bendra | Vienas `main`, viena svetainė. |
| ISIN sąrašas (`config/securities.csv` ← Google Apps Script saugykla) | Bendra | Vienas sąrašas visiems. Žr. žemiau. |
| DD archyvas (`data/dd/`, užšifruotas) | Bendra | Pridėti gali bet kas (viešasis raktas). Atidaryti – tik su slaptažodžiu. |
| Claude projekto failai (`/mnt/project-files/`) | **Atskira** | Kiekvienas Claude projektas – savo aplankas (žr. 6 d.). |
| Projekto atmintis, pokalbiai, paskyros skills | Atskira | Žr. 1 d. |

**ISIN sąrašas.** Saugyklos adresas yra viešame `config/settings.yaml`, o PIN neįjungtas, todėl sąrašą dabar gali pakeisti **bet kas internete**, o kasdienis `isin-sync.yml` jį perkelia į repozitoriją. Galimybės:
- (a) **Rekomenduojama:** vienas bendras sąrašas + PIN. Google Apps Script → Project Settings → Script properties → `PIN` = kodas; kodą perduoti kolegai asmeniškai (ne repozitorijoje, ne el. laišku atviru tekstu). Dashboard'as PIN paklausia pats; laukas „Jūsų vardas“ rodo žurnale, kas keitė.
- (b) Du atskiri sąrašai – reikėtų perrašyti saugyklą ir dashboard'ą, o ataskaitos vis tiek būtų vienoje svetainėje. Verta tik jei portfeliai iš esmės skirtingi.

## 6. Ko kolegos Claude neturės ir kaip tai išspręsti
| Priklausomybė | Kur dabar | Kaip pasidalinti saugiai |
|---|---|---|
| Darbų generatoriai ir metodika (~100 Python failų, `metodika.md`, `struktura_etf_dd.md`, `review_lib.py`, `funds_dd_lib.py`, `make_duration.py`, `accounting_fields.py`, …) | Tik Vyčio projekto failuose | Ne konfidencialius – perkelti į repozitoriją (`tools/` ar `methods/`), patikrinus, kad nėra vidinių duomenų. Kol neperkelta, Funds DD, Fund review, Duration ir Apskaitos formos darbai veikia tik Vyčio projekte. |
| SEB „Liquid Funds DD Template“ (pptx) | Vyčio projekto failai | Į viešą repo **nekelti**. Kolega įkelia jį į savo Claude projekto failus pats (iš SEB vidinių sistemų). |
| SEB politikų PDF | Vyčio projekto failai | Jei tai vieši SEB tinklapio dokumentai – kolega atsisiunčia pats; santrauka jau yra `config/seb_policies.yaml`. |
| DD archyvo slaptažodis | Tik Vyčio galvoje / slaptažodžių tvarkyklėje | Perduoti asmeniškai. Niekada į repo, pokalbį ar atmintį. |
| DD manifest (`/mnt/project-files/dd/manifest.json`) ir pradiniai DD failai | Vyčio projekto failai | `dd_vault.py add` veikia ir be jo (įrašas į manifest praleidžiamas), bet `replace`/`rekey` reikalauja manifest – tai daro tas, kurio projekte DD sukurtas. |
| GitHub Actions paslaptys | Tik nebūtinas `CLAUDE_CODE_OAUTH_TOKEN` (`analyze.yml`; jei įrašytas – naudoja jį įrašiusiojo Claude planą) | Kolegai nieko nereikia; paslapčių nekopijuoti. Visi workflow naudoja GitHub suteikiamą raktą; kolega su „Write“ teise gali juos paleisti (Actions → workflow → Run workflow). |
| Google Apps Script (ISIN saugykla) | Atskira Google paskyra | Kolegai paskyros nereikia; reikia tik PIN (jei įjungtas). |

## 7. Kolegos prisijungimas (vieną kartą)
1. Priimti GitHub kvietimą į `vytiscode/issuers-review` (teisė „Write“) – el. laiške arba https://github.com/notifications.
2. claude.ai → Claude Code → prijungti GitHub (https://claude.ai/connect-github) ir įdiegti Claude GitHub App savo paskyrai, jei klausia.
3. Susikurti savo Claude projektą ir jame pridėti repozitoriją `vytiscode/issuers-review`. Projekto instrukcijose užtenka vienos eilutės: „Dirbk pagal repozitorijos CLAUDE.md ir rules/; nuolatines taisykles rašyk į rules/ per PR (rules/team_work.md).“
4. Į savo projekto failus įkelti tai, ko jam reikia ir kas negali būti viešai (pvz., SEB DD šabloną).
5. Gauti iš Vyčio asmeniškai: DD archyvo slaptažodį ir ISIN sąrašo PIN (jei įjungtas).
6. Dashboard: https://vytiscode.github.io/issuers-review/ – peržiūrai nieko nereikia; ISIN sąrašo keitimui – PIN; DD archyvui – slaptažodis; mygtukai „Run workflow“ veda į GitHub – reikia būti prisijungus su „Write“ teise; „Copy prompt“ mygtukų tekstą įklijuoti į **savo** Claude projektą.
7. Pirmas bandymas: savo projekte parašyti „paieškok naujų ataskaitų ir naujienų“ (nemokama, be AI) ir patikrinti, kad Claude paleidžia `update.yml`.
