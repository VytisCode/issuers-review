# Fund and issuer analysis standards

> Moved into the repo 2026-11-23 from the account skill "fund-analysis-standards" (text kept, personal working-style notes left out – those are in CLAUDE.md "Kaip dirbti"). This file is the only copy: change a rule here, by pull request, so both editors' Claude sessions get it. "Owner" below means the editor who asked for the work.

These rules apply to every investment analysis document: fund/ETF due diligence, ETF screening, challenge of held positions (SWITCH / KEEP / WATCH), quarterly fund reviews, issuer and bond reviews. If the project has its own methodology file (for example `metodika.md`, `struktura_etf_dd.md`, CLAUDE.md), follow it first. This skill covers the shared principles.

The readers are colleagues who check numbers and skim. Documents are filed internally as investment documentation and must pass an internal check.

## 1. Facts and sources
- Every number and every claim has a source tag (`[S1 p.4]`), and the tag links to the source. Never fill anything in from memory.
- Official sources come first: the manager's factsheet, KID/KIID, prospectus, annual or semi-annual report, and the issuer's own report. Aggregators (justETF, Morningstar, broker pages) are for time series and cross-checks only. Label them as such, and say when no official figure exists.
- Compare sources inside the sentence and give exact quotes with page numbers, so the reader can find them with Ctrl+F. Example: "Morningstar shows 0.41% [M8]; KIID shows 0.47% [K1 p.2]".
- Not found: say "not found" and name the document that was searched (for example "no exclusion screen found in the Sustainability Annex [P1]"). Use red for a draft and grey for a final version. Never guess.
- **Final version** (owner, 2026-11-23): no red anywhere and no "NOT FOUND" wording; missing items read "n/a – <reason>" in grey. Add a "Method and data availability" section before the sources, with these groups: "Not available from public sources or data providers as of <date>", "Not covered by the documents read", "Not assessed (internal decision)", "Recorded outside this file" (for example sanctions screenshots). Under the decision, say what the conclusion rests on (documents, data providers, internal data).
- Mark our own calculations as "own calculation". Mark figures taken from a summary rather than the original text.
- **Pinpoint references, so every number can be found in under a minute** (owner, 2026-11-13). A bare tag such as `[C1]` or `[D3]` is not enough.
  - PDF: always give the page, for example `[D1 p.27]`.
  - Excel, whether our calculations workbook or a manager's file: give the sheet and the cell or row, for example `[C1 “S07 Performance” B26]` or `[D3 'Live Long-Short Attribution' J21]`. This applies in running text too (conclusions, risks), not only under tables.
  - If the source stores the number in a different form, say how to find it. Examples: a ratio instead of a percent ("gross 355% of NAV; D3 shows 3.555"), or millions instead of units. If the number is rounded, averaged or picked from a series, say which: last month, an average, or the worst value.
  - Every number in the text that also appears in the workbook goes into the automatic slide-vs-Excel check.
  - Before delivery, scan the document for tags without a page, sheet or cell. The list must be empty.
- Use public data only (no Bloomberg unless the owner pastes it). Store dated copies of source documents (with sha256) in the evidence package.
- **Internal data from the owner** (owner, 2026-11-23): label each value only with its source and date, for example "(Bloomberg, 2026-11-23)" or "(internal assessment, 2026-11-23)". Never write "owner-provided" or "owner" in a document; write "internal decision", not "owner decision". Keep the values on their own workbook sheet ("Data internal") with pinpoint tags. If a pasted value has no unit or is ambiguous (for example "0,065" or "252,9k"), use the most likely unit (the data provider field's unit), write it in the document, and name the assumption in the reply so the owner can confirm it.

## 2. Verifiability (instead of screenshots)
- Under every table and chart, add a grey "Data:" line with the dataset, source, share class or ISIN, period, retrieval date, how the figure was calculated (in words), and the Excel sheet.
- Provide a calculations workbook: inputs with link and date, every result as an Excel formula, every chart also as a native Excel chart, and validation status ("not validated" until a second person checks it).
- Where a PDF exists, add a crop of the official document for key facts (fees, SFDR article), captioned with the document, page and date.
- Every comparison names what it is compared with, in the same sentence or table row: "0.00% vs 0.67% for the Article 6 MSCI World ETF", never a bare "vs 0.67%".
- Name the comparison fund or index with its ISIN in every table. Show the fund return, the reference return and the difference in separate columns. Say how they are compared: same index, different index, hedged, or different market.
- Main performance period: 3 years to the latest month-end with official factsheets. A grey "Latest" column may run to the last daily price. Both dates go on the cover.
- Evidence ZIP for non-IT colleagues contains only PDF and Excel: a README PDF (how to check), one Excel workbook with a "Start here" sheet and one sheet per slide that has a table or chart, and source PDFs named `<key – document (date)>`. Never include json, py, png or cache files.

## 3. Layout
- When a template exists, the first slides follow it strictly in structure and order. Leave out template slides that do not apply (for example "Fund overview (index funds)" for an active fund). Then add a divider slide, "Additional analysis", followed by the extra slides. Do not change the template's form.
- Order everywhere: a short comment (what it shows) comes first, then the table or chart, then notes (period, method, colours, source). No "Read first" boxes.
- Every analytical text has a supporting table or chart, and every chart or table has a comment.
- Bar charts always show value labels. Do not use Sharpe ratios. Use monthly heat maps (fund minus reference, pp) where the data is reliable, also for alternatives.
- Check PowerPoint files as PowerPoint renders them, not only LibreOffice. For example, a table cell must have exactly one fill, because PowerPoint uses the first fill and LibreOffice the last.
- Tables and sections do not break across pages. Navigation in long PDFs is clickable.
- Each category or section stands on its own: list, conclusion, signals, decisions, evidence chart, and open data.

## 4. Content rules
- **Active fund:** compare with its benchmark. **Passive screened ETF:** compare with its parent or core index, so the effect of the screen is visible.
- **ESG:** report what the manager publishes (SFDR article as the manager states it). Add third-party indicators (Morningstar: WACI, carbon risk, product involvement, severe controversies), saying honestly per fund what was and was not available, with the holdings date. An Article 8 or 9 fund must be shown against an Article 6 alternative. Use 3-year trend lines.
- **Media and sanctions:** cover the manager, the management company, the fund and any former name, over 5 years. For each item give the date, publisher, what happened, who is involved, the status or outcome, and the relevance. The reader must not need to research further. Sanctions are only links and what to check; the owner does the check. Screen companies only: the management company and the investment manager, not the fund or umbrella. Give one screenshot box per company per list, for example 2 companies × 3 lists = 6 boxes.
- **Alternatives:** use instruments that can be bought now, with a source on every row. Add comments for what was compared, what it shows and the key insight. The final slide is titled "Final comparison view".
- **Market context:** facts and scenarios with sources, no forecasts. State causes only where a source states them.
- **Verdicts:** use one word (for example SWITCH / KEEP / WATCH, or Positive). Give red flags clearly. The document is final: give the reader no tasks.

## 5. Tone and language
- Use formal documentary English in third person. Do not write "you / we / our" or "as you asked". Be blunt in substance (clear verdicts, numbers), not in tone.
- Use plain words without jargon or unexplained acronyms. Every observation carries figures and ISINs, because a sentence without a number lowers trust.
- Emails to colleagues are information updates only. They never ask for approval or decisions.
- Talk with the owner in Lithuanian. Documents are in English unless the request says otherwise.

## 6. Working with the owner (the editor who asked)
- Missing data: list at most 15 items, only data that is not public, with fields for the owner to fill in (Morningstar Direct, Bloomberg, internal limits). Each item says exactly what to look up (the figure, the share class, the date), where (system → screen → field), a checked link, a search phrase, an example answer and why it matters. Then write the data into the same file and re-issue it as the final version.
- **Speed** (owner, 2026-11-23, after a slow ETF DD): start from the latest generator of the same kind (an index ETF or an active fund) and change only the fund-specific blocks. In umbrella documents (an annual report of several hundred pages), find the fund's pages first and read only those. Do not read dashboard or tool code when the project rules already give the command. Render and check only the slides that changed.
- Anything that costs money needs the owner's approval first. AI work runs only on request.
- Files with non-public content stay in project files or an encrypted archive, never in a public repo.