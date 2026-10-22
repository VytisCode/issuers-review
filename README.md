# Emitentų apžvalga (issuers-review)

Įrankis kasmetinei (ir dažnesnei) turimų obligacijų ir akcijų emitentų apžvalgai: renka viešą informaciją,
pildo vienodą ataskaitos šabloną su šaltiniais ir rodo viską dashboard'e.

- **Dashboard:** https://vytiscode.github.io/issuers-review/ (GitHub Pages iš `docs/`)
- **Ataskaitos:** `reports/<emitentas>/<laikotarpis>.yaml` – tekstas (taisomas GitHub'e); Word / PDF / Excel generuojami į `docs/files/`.
- **Rodiklių istorija:** `data/history/<emitentas>.csv` – kiekvienas skaičius su šaltiniu ir puslapiu.
- **Sąrašai:** `config/issuers.yaml` (emitentai), `config/securities.csv` (ISIN, kuponai, išpirkimo datos), `config/seb_policies.yaml` (SEB politikų santrauka ir ESG klausimynas).
- **Automatika (nemokama):** `.github/workflows/update.yml` – kas pirmadienį naujienos (`tools/fetch_news.py`) ir naujų ataskaitų paieška (`tools/check_filings.py`), be AI.
- **AI analizė:** paprašykite Claude projekte arba dashboard mygtuku (žr. žemiau).

## Ką reikia įjungti vieną kartą
1. **GitHub Pages:** Settings → Pages → Source: *Deploy from a branch* → Branch: `main`, folder `/docs` → Save.
2. *(Neprivaloma)* **Automatinė AI analizė mygtuku:** kompiuteryje su Claude Code paleiskite `claude setup-token`, gautą raktą įrašykite
   Settings → Secrets and variables → Actions → New repository secret, pavadinimas `CLAUDE_CODE_OAUTH_TOKEN`. Naudojamas jūsų Claude plano limitas.

Darbo taisyklės Claude – `CLAUDE.md`.
