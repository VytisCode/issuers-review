# Funds DD struktūra – v4.1 (2026-11-12) – viskas patikrinama (savininkas tikrina kiekvieną sakinį ir skaičių); galioja kartu su v4 ir v3 žemiau
1. **Vienas Excel [C1]** (`calc_book.py`; pavyzdys `aqr_calc_book.py`): „Start here“ → lapai „S## <skaidrė>“ (ta pati lentelė ir diagrama kaip skaidrėje, skaičiai – formulės) → lapai „Data …“ (įvestys mėlynai, šaltinis kiekvienoje eilutėje / stulpelyje). Morningstar finansiniai ir ESG duomenys – Excel'e su SecId, datapoint, nuoroda ir data.
2. **Automatinė patikra**: kiekvienas skaidrės skaičius registruojamas `expect` ir po LibreOffice perskaičiavimo lyginamas su Excel – neatitikimų turi būti 0 (AQR rasta: „behind in 2 of 13 years“ → 1 of 13).
3. **Data eilutė**: „… [C1] – Excel sheet “S07 Performance”“; [C1] tekste – tik kai skaičius yra Excel'e.
4. **Citatos**: [P1 p.N] tikrinama, ar citata / skaičius yra tame puslapyje (teksto kopijoje puslapis prasideda eilute „N“; PDF – `L.txt_pages_pdf`, PDF puslapis = spausdintas; AQR rasta 88 nuorodos +1 psl.). Ekrano iškarpos – tik skaidrėse (laikinas katalogas), antraštėje [D1 p.30].
5. **ZIP** (dashboard'e vienas mygtukas „Documents and calculations (ZIP)“): tik PDF ir Excel – Excel [C1], README PDF, P1/K1/D# (gauti), E#/N#/J# (tinklalapiai PDF), „00_Source_list.pdf“. Be json, py, png, __pycache__. Įrodymai – `dd_evidence/<fondas>/package/` (tik Excel + README).

# Funds DD struktūra – v4 (2026-11-12, kolegos pastabos AQR DD) – GALIOJA kartu su v3 žemiau (v4 viršesnė)

1. **Nenaudojamų šablono skaidrių nerodome** (pvz. „Fund overview (index funds)“ aktyviam fondui, „ESG overview (climate future index fund)“). `Deck.na_slide` išjungtas. Likusios šablono skaidrės – šablono tvarka.
2. **Portfelio duomenų data.** Prieš DD patikrinti Morningstar `HS000` (portfolio date). Jei portfelis senesnis už fondą / strategijos pakeitimą – sektorių ir pozicijų skaidrių **nededame**; „Current portfolio characteristics“ – iš **paskutinės valdytojo mėnesio ataskaitos** (ekspozicijos, grąžos šaltiniai, pozicijų skaičius; šalių gali nebūti) + raudonai „to be updated when current holdings are received“.
3. **ESG skaičiai lieka** (vėliau jų prašo), bet kiekvienoje ESG lentelėje ir skaidrėje – ryški pastaba apie Morningstar portfelio datą; jei reikšmė šokinėja be naujo portfelio – parašyti, kad patikimumas žemas.
4. **Grafikai:** fondas vs indeksas per metus – viena serija „fondas − indeksas, pp“ žalia/raudona (`L.signed_bars`) + mažoji lentelė su abiem grąžomis; ESG tendencijos – stulpeliai pagal datas (12/YYYY … paskutinė), ne linijos (susilieja); dienos linijoms – `L.tick_skip` ir mėnesių etiketės.
5. **Strategija:** valiutos rizikos nerašyti (mūsų portfeliuose ji visada yra).
6. **Skyrius „Performance: drawdowns“** (visiems DD, po „when the strategy wins and loses“): (a) gyvas laikotarpis – kritimas nuo piko (Morningstar Daily Return Index, EUR) fondas vs indekso ETF vs turimas ETF, rinkos kritimų ≥ 2 % lentelė tais pačiais datų langais, savaitinis downside/upside capture; NAV ir biržos uždarymo laikas skiriasi – tik kelių dienų / savaitės langai; (b) ilga istorija (proforma ar faktinė) – indekso kritimų ≥ 7 % epizodai, atsigavimo laikas, mėnesinis capture, mėnesiai < −3 %; (c) Morningstar capture ratios (RR158/RR152/RR159) fondui ir alternatyvoms su pastaba, kokį laikotarpį jos apima. Rinkos įvykio priežastis – tik su perskaitytu šaltiniu (pvz. [N9]).
7. **Peržiūros klausimai** (pvz. mokesčiai) – Conclusion „Open points“: klausimas + ką sako dokumentai (puslapiai) + kas turi patvirtinti.
8. Savininko pateikti duomenys (pvz. Bloomberg ticker) – „owner-provided, <šaltinis>, <data>“.

---
# Funds DD struktūra – v3 (2026-11-01, Vytis po AQR DD peržiūros) – GALIOJA

Šablonas `Liquid_Funds_DD_Template.pptx` sutikrintas su SEB originalu (Copilot skirtumų nerado, 2026-11-01).
Bendri įrankiai – `funds_dd_lib.py` (`TEMPLATE`, `Deck.na_slide`, `Deck.divider`, `data_note`, `heatmap`, `pdf_crop` + `evidence_picture`,
`MEDIA_HEAD`, `ucits_bbg_line`). Pavyzdys – `make_dd_aqr_delphi.py` (aktyvus fondas, 36 skaidrės).

## 1. Šablono dalis – skaidrės 1–13 griežtai šablono tvarka, po vieną
1 Fund name · 2 Conclusion · 3 Fund overview (index funds) · 4 Fund overview (mutual funds) · 5 Index overview · 6–7 ESG overview ·
8 ESG overview (climate future index fund) · 9 Performance · 10 Investment team · 11 Investment strategy & process ·
12 Current portfolio characteristics · 13 Other.
- Netinkama skaidrė **lieka** su „**Not applicable** – priežastis“ ir nuoroda, kur yra tinkama (`d.na_slide`), kad matytųsi visas šablonas.
- Šablono skaidrėje – santrauka ir „**Details:** slides N–M“ į papildomas skaidres.
- ESG lentelėse prie prospekto fakto – Morningstar faktas tame pačiame sakinyje: „no screen found in the Sustainability Annex [P1 p.x]; Morningstar: 1.43% of portfolio involved (data) [M#]“.

## 2. Skyriklis „Additional analysis“ (`d.divider()`), po jo – papildomos skaidrės
Pavadinimas prasideda šablono skaidrės vardu („Performance: …“, „ESG overview: …“, „Current portfolio characteristics: …“,
„Possible financial instrument alternatives: …“), ≤ ~52 ženklų (35 pt telpa į vieną eilutę); turinys nuo y = 1.5.
Tvarka: Conclusion and Investment decision (3 lentelės) → strategijos schema → mokesčiai (+ KIID / faktų lapo iškarpa) →
ESG Morningstar rodikliai → ESG 3 m. tendencijos → sankcijos → žiniasklaida → Performance (šilumos žemėlapis, kada laimi/pralaimi, gyvi rezultatai) →
portfelis (sektoriai/regionai, didžiausios pozicijos) → alternatyvos (lentelė, grąžos, mėnesių šilumos žemėlapis, „Final comparison view“) →
apskaita (LT) → Sources.

## 3. Patikrinamumas (vietoj ekrano kopijų)
- Po **kiekviena** lentele / diagrama – `L.data_note`: „Data: duomenų rinkinys (šaltinis, klasė, laikotarpis, paėmimo data); kaip skaičiuota žodžiais; Excel lapas [C#]“.
- Skaičiavimų Excel su formulėmis – `dd_evidence/<fondas>/…calculations….xlsx` ([C2]); įrodymų ZIP – `dd_vault.py evidence <id> dd_evidence/<fondas>`.
- Oficialaus dokumento iškarpa (`L.pdf_crop` + `L.evidence_picture`) – mokesčiai iš KID/KIID, kai yra PDF; parašas su dokumentu, puslapiu, data.
- Sakinyje nurodyti, kur rašoma, ir lyginti šaltinius: „Morningstar shows 0.41% … [M8]; KIID shows 0.47% [K1]“; tikslios citatos su puslapiu (Ctrl+F).
- Kiekvienas „vs x%“ įvardija, su kuo lyginama (pvz. „vs 0.67% Article 6 MSCI World ETF“) – Vytis 2026-11-01.
- Stulpelinės diagramos – visada su reikšmėmis virš stulpelių (bibliotekos numatytoji). Sharpe nerodyti.

## 4. ESG – Morningstar (jungtis, be Excel įkėlimo)
Prašomi rodikliai: WACI Scope 1+2 (EUR), PAI GHG intensity, oil sands, PAI fossil fuel, tobacco, gambling, thermal coal, thermal coal power,
PAI controversial weapons, Arctic oil & gas, PAI UNGC/OECD, % AUM severe controversies, carbon solutions.
Jungtis duoda (fondo lygiu): ESG2H WACI Scope 1+2 **USD** terms, ESG2A historical carbon risk score, ESG1X carbon risk score, ESG1W low carbon designation,
ESG62 thermal coal, ESG63 tobacco, ESG55 gambling, ESG53 controversial weapons, ESG42 % severe ESG risk, HS000 portfolio date; kategorijos vidurkiai ESG1I–ESG1U.
Licencija atsako (403): fondo fossil fuel, oil sands, Arctic, thermal coal power, carbon solutions; PAI ir controversies rodiklių jungtyje nėra.
- Kiekvienam DD – sąžiningai: kas rasta / nerasta kiekvienai investicijai; palyginti su SFDR priedu / periodine ataskaita.
- Lentelė: fondas vs alternatyvos + **Article 6** palyginamasis (pvz. iShares Core MSCI World 0P0000MEHZ) – Art. 8/9 fondas turi būti geresnis už Art. 6;
  raudona – blogiau už Art. 6, žalia – geriausia eilutėje. Visada nurodyti Morningstar holdings datą (sena → raudonai).
- 3 m. linijiniai grafikai (ketvirčių pabaigos + paskutinė reikšmė): WACI, historical carbon risk, thermal coal, gambling (arba kiti reikšmingi).

## 5. Performance ir portfelis
- Performance: kalendoriniai metai, rizikos lentelė, mėnesių šilumos žemėlapis (fondas − atskaitos indeksas, pp), kada laimi/pralaimi (penktadaliai, blogiausi mėnesiai), gyvi rezultatai; proforma visada pažymėta.
- Portfelis: aktyvus fondas – lyginti su atskaitos indeksu; **pasyvus screened ETF – su baziniu (core) indeksu**, kad matytųsi skirtumas. Sektoriai, regionai, pozicijų skaičius, įmonių dydis, 10 didžiausių pozicijų.

## 6. Alternatyvos
Lentelė su šaltiniu kiekvienoje eilutėje + „What was compared / What it shows / Key insight“ komentarai; grąžų diagrama su reikšmėmis; mėnesių šilumos žemėlapis (12 mėn.) ir skirtumo eilutė; galutinė skaidrė – „**Final comparison view**“ (ne „Honest view“).

## 7. Žiniasklaida
Lentelė `L.MEDIA_HEAD`: Date · Published by (kas paskelbė) · What happened · Who is involved · Status / outcome (baigta ir kaip / vyksta) · Relevance · Source.
Tikrinti valdytoją, valdymo įmonę ir fondą (ir buvusį pavadinimą); 5 metai; žodžiai, data – skaidrėje. Turi užtekti be papildomo tyrimo.

## 8. Apskaita
Bloomberg `FUND_EURO_DIRECT_UCITS` eilutė – tik jei UCITS neįrodytas dokumentuose (`L.ucits_bbg_line(isin, proven=...)`), kitaip ne raudona.

---
## Ankstesnė versija (2026-10-28)

# ETF DD struktūra (projektas, 2026-10-28)

Pagrindas: SEB „Liquid Funds DD Template“ (13 skaidrių, forma nekeičiama) + praktika iš 3 atliktų DD (XDWC, IS0E, LYBK; 32–43 skaidrės).
Netinkamos šablono skaidrės (mutual funds, climate future index fund) neįtraukiamos. ESG lentelių antraštė – „ESG overview“.
Kalba: pagrindinis tekstas – dashboard kalba (pavyzdžiuose anglų), apskaitos skaidrės – lietuviškai. Šablono etiketės – nekeičiamos.

| # | Skaidrė | Turinys | Šaltinis |
|---|---|---|---|
| 1 | Titulinė | Fondas, valdytojas, kontaktas (viešas el. paštas, jei nėra – „not found“), prepared by, data | valdytojo svetainė |
| 2 | Conclusion and Investment decision | sprendimas (Positive / Negative / Positive with conditions), kontrolinis sąrašas, raudonos vėliavos aiškiai | – |
| 3 | Investment strategy & process | valdymo būdas, replikacija, indeksas, visata, svoriai, limitai, rebalansavimas, išvestinės, skolinimas, rizikos | prospektas, KID, faktų lapas |
| 4 | Fund overview (index funds) | šablono lentelė; kiekvienoje eilutėje data ir šaltinis; spread su vienetais ir birža | faktų lapas, birža, justETF / Bloomberg |
| 5 | Index overview | šablono lentelė, aprašymai vietoj „Reference to the fact sheet“ | indekso teikėjas |
| 6 | Index details | indekso metodika, sudėtis (top 10, šalys, sektoriai) | indekso faktų lapas |
| 7–8 | ESG overview I, II | šablono lentelės; Yes/No + šaltinis kiekvienoje eilutėje; problemos raudonai | prospekto SFDR priedas, UN PRI, valdytojo politika |
| 9 | ESG analysis | fondo tvarumo rodikliai, anglies intensyvumas, iškastinio kuro dalis | SFDR periodinė ataskaita, viešas ESG reitingas |
| 10 | Compliance with the list of sanctions | 3 nuorodos + ekrano nuotraukos (OFAC, ES žemėlapis – įrankis; eNameChecker – savininkas). **Nuo 2026-11-13: `L.sanctions_slides(d, names)`** – žingsnių lentelė + rezultatų lentelė vardas × sistema („Match: Yes / No“), po jos po skaidrę kiekvienam vardui su 3 dideliais langeliais (vardai × 3 sistemos); **tikrinamos tik įmonės – valdymo įmonė ir investicijų valdytojas, fondas / skėtis netikrinamas** (savininkas 2026-11-13), t. y. paprastai 2 × 3 = 6 nuotraukos; ORDER – `L.sanction_keys(names)` | OFAC, sanctionsmap.eu, eNameChecker |
| 11 | Media checks | dokumentuota paieška: žodžiai, laikotarpis (12 mėn.), radiniai su nuorodomis, išvada | Google / žiniasklaida |
| 12–13 | Performance | 1Y, 3Y, 5Y, Max: fondas vs indeksas vs alternatyvos; kalendoriniai metai | valdytojo NAV istorija (mūsų skaičiavimas), faktų lapas |
| 14 | Investment team | valdytojas, valdomas turtas, komanda | valdytojo svetainė |
| 15 | Current portfolio characteristics | fondo sudėtis + poveikis mūsų portfeliui | faktų lapas + savininko duomenys |
| 16–18 | Possible financial instrument alternatives | 4–6 alternatyvos: TER, AUM, likvidumas, tracking, rezultatai, SFDR, persidengimas; nuoroda kiekvienoje eilutėje | justETF, valdytojai |
| 19 | Other | šablono tekstas + UCITS paaiškinimas | – |
| 20 | Sources | visų šaltinių sąrašas su nuorodomis ir datomis | – |
| 21–22 | Additional information for the fund accounting group (LT) | Tipas, Potipis, prekyba, perleidžiami VP, pinigų rinkos priemonė, SFDR, PSPKĮ 49 analizė; regionas, KIS tipas, turto klasė, Positioning, diversifikavimas, UCITS, ETF, rizikos lygis | prospektas |

## Taisyklės, kurių pavyzdžiuose trūko (įrankis laikysis visada)
- Kiekvienas skaičius – su data ir šaltiniu; spread – vienetais (% arba bp) ir nurodant biržą.
- Tracking difference ir tracking error – atskirai (faktinis ir prospekte nurodytas).
- Tuščias laukas neleidžiamas: „not found“ arba „N/A – priežastis“.
- Patikra, ar tekste neliko kito fondo pavadinimo, indekso ar ISIN (LYBK klaida).
- Media patikra visada su paieškos žodžiais, laikotarpiu ir išvada.

## Šaltiniai ir dokumentų ZIP (v3.1, 2026-11-01 – AQR DD pataisymas)
- Sources skaidrė ir „Dokumentai (ZIP)“ sutampa 1:1; prie kiekvieno šaltinio parašyta, kur jis yra (`funds_dd_lib.where`): Documents ZIP / received copy / Evidence ZIP (M#, C#) / link only (X#) / headline only.
- Šaltiniu dedamas tik perskaitytas dokumentas, kurį savininkas gali atsidaryti. URL – tiesiai į dokumentą (ne svetainės pradinis puslapis). Jei viešos nuorodos nėra, bet failą pateikė savininkas – `LOCAL` (failas į ZIP).
- Mokamas / prisijungimo straipsnis – tik „headline only“ (`HEADLINE`), su nuoroda, ZIP'e tekstinis failas; ne įrodymas – faktą turi patvirtinti kitas, perskaitytas šaltinis (pvz. Hedgeweek vietoj Bloomberg).
- Savininko pateikti failai (D#) – visi į ZIP ta pačia numeracija (savininko sprendimas 2026-11-01), užšifruotai.
- `L.sources_json(path, SOURCES, DOCS, LOCAL, EXPECT, headline=HEADLINE)` – sustabdo generatorių, jei liko pradinis puslapis ar paywall be HEADLINE; `EXPECT` – žodžiai, kurie turi būti atsisiųstame dokumente (ISIN iš pavadinimo tikrinamas automatiškai).
- PDF iškarpos (`L.pdf_crop` + `L.evidence_picture`) – svarbiausiems skaičiams ir sąlygoms, kai šaltinis yra PDF; iš teksto kopijos iškarpų nedaryti (tik tiksli citata su puslapiu).

### v4.2 – tikslios nuorodos kiekvienam skaičiui (savininkas 2026-11-13)
Savininkas nerado „worst rolling 12 months −9.8 pp [C1]“ (Excel'e 21 lapas) ir „gross ~355% of NAV [D3]“ (D3 saugo 3.555, ne 355).
- **Joks [C1] ar Excel šaltinis be lapo ir langelio:** `[C1 “S07 Performance” B26]` – `Sheet.tag(row, col, val)` (calc_book.py), kuris kartu įtraukia skaičių į automatinę patikrą. Tai galioja ir tekste (išvadose, rizikose), ne tik po lentelėmis.
- **Valdytojo Excel [D#]:** lapas ir langelis / eilutė: `[D3 'Live Long-Short Attribution' J21]`. PDF [D#] – visada su puslapiu `[D1 p.27]`.
- **Jei šaltinis saugo skaičių kitu pavidalu** (dalis vietoj %, USD vs mln.), parašyk, kaip jį rasti: „gross 355% of NAV (D3 shows 3.555) [D3 'Live Long-Short Attribution' J21, 30.09.2026]“. Apvalinant arba imant vidurkį – pasakyti, ką tiksliai (paskutinis mėnuo / vidurkis 04–09/2026).
- **Kiekvienas teksto skaičius, kuris yra Excel'e, eina į `Book.checks`** – ne tik lentelių skaičiai.
- Prieš pateikiant – patikra: nėra nė vienos `[C1]`, `[D#]` (PDF) be puslapio / lapo ir langelio (regex per sugeneruotą pptx tekstą).
