/* Emitentų apžvalgos dashboard. Duomenys - data.js (generuoja tools/build_site.py). */
(function () {
  const D = window.DATA;
  const $app = document.getElementById("app");
  const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const byId = Object.fromEntries(D.issuers.map(i => [i.id, i]));
  const gh = p => `https://github.com/${D.repo}/${p}`;
  const ST = {
    met: ["Įvykdyta", "good"], partly: ["Iš dalies", "warn"], not_met: ["Neįvykdyta", "bad"], pending: ["Dar vykdoma", "info"], changed: ["Pakeista", "warn"],
    yes: ["Atitinka", "good"], no: ["Neatitinka", "bad"], review: ["Reikia peržiūrėti", "warn"], found: ["Rasta", "bad"], not_found: ["Nerasta", "good"]
  };
  const chip = s => { const [t, c] = ST[s] || [s, "info"]; return `<span class="chip ${c}">${esc(t)}</span>`; };
  const fmt = (v, u) => {
    if (v === undefined || v === null || v === "") return "–";
    const dec = String(v).includes(".") ? String(v).split(".")[1].length : 0;
    let s = Number(v).toLocaleString("lt-LT", { minimumFractionDigits: dec, maximumFractionDigits: dec });
    return u === "%" ? s + " %" : u === "x" ? s + "x" : s;
  };
  const unitLbl = u => ({ EURm: "mln. EUR", EUR: "EUR", GW: "GW", Mt: "mln. t", m: "mln." }[u] || "");
  const days = d => Math.round((new Date(d) - new Date(D.today)) / 864e5);
  const recentNews = (id, n = 30) => D.news.filter(x => x.issuer === id && x.date && days(x.date) >= -n);
  const ACTIONS = gh("actions/workflows/update.yml");
  const issueUrl = (title, body) => gh(`issues/new?labels=analize&title=${encodeURIComponent(title)}&body=${encodeURIComponent(body)}`);

  /* ---------- paprastas stulpelinis grafikas (viena serija, su užuominomis) ---------- */
  function barChart(points, unit) {
    if (!points.length) return '<p class="muted">Nėra duomenų.</p>';
    const W = 560, H = 220, L = 48, B = 28, T = 12;
    const max = Math.max(...points.map(p => p.v), 0) * 1.12 || 1, min = Math.min(0, ...points.map(p => p.v));
    const y = v => T + (H - T - B) * (1 - (v - min) / (max - min));
    const bw = Math.min(56, (W - L - 10) / points.length * 0.6), step = (W - L - 10) / points.length;
    const ticks = [0, .25, .5, .75, 1].map(f => min + (max - min) * f);
    let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Grafikas">`;
    ticks.forEach(t => { s += `<line class="ax" x1="${L}" x2="${W - 6}" y1="${y(t)}" y2="${y(t)}" stroke-width="1"/><text x="${L - 6}" y="${y(t) + 4}" text-anchor="end">${fmt(+t.toFixed(t >= 100 ? 0 : 1), unit)}</text>`; });
    points.forEach((p, i) => {
      const x = L + step * i + (step - bw) / 2, y0 = y(Math.max(0, p.v)), h = Math.abs(y(p.v) - y(0));
      const r = Math.min(4, bw / 2, h);
      const path = p.v >= 0
        ? `M${x},${y0 + h} V${y0 + r} Q${x},${y0} ${x + r},${y0} H${x + bw - r} Q${x + bw},${y0} ${x + bw},${y0 + r} V${y0 + h} Z`
        : `M${x},${y(0)} H${x + bw} V${y(0) + h} H${x} Z`;
      s += `<path class="bar" d="${path}" data-tip="${esc(p.label)}: ${esc(fmt(p.v, unit))} ${esc(unitLbl(unit))}${p.src ? " · " + esc(p.src) : ""}"/>`;
      s += `<rect x="${L + step * i}" y="${T}" width="${step}" height="${H - T - B}" fill="transparent" data-tip="${esc(p.label)}: ${esc(fmt(p.v, unit))} ${esc(unitLbl(unit))}${p.src ? " · " + esc(p.src) : ""}"/>`;
      s += `<text x="${x + bw / 2}" y="${H - 9}" text-anchor="middle">${esc(p.label)}</text>`;
    });
    return s + "</svg>";
  }
  const tip = document.getElementById("tip");
  document.addEventListener("mousemove", e => {
    const t = e.target.closest && e.target.closest("[data-tip]");
    if (!t) { tip.hidden = true; return; }
    tip.textContent = t.getAttribute("data-tip"); tip.hidden = false;
    tip.style.left = Math.min(e.clientX + 12, innerWidth - 290) + "px"; tip.style.top = e.clientY + 14 + "px";
  });

  /* ---------- puslapiai ---------- */
  function overview() {
    const st = D.status || {};
    let h = `<h1>Emitentų apžvalga</h1><p class="sub">${D.issuers.length} emitentai, ${D.securities.filter(s => s.held === "yes" && s.status === "active").length} turimi vertybiniai popieriai. Naujausi duomenys: ${esc(D.generated)}.</p>`;
    h += `<div class="row" style="margin-bottom:14px"><a class="btn primary" href="${ACTIONS}" target="_blank" rel="noopener">↻ Ieškoti naujų ataskaitų ir naujienų dabar</a>
      <span class="muted">Automatiškai – kiekvieną pirmadienį. Paskutinė paieška: naujienos ${esc(st.news?.date || "dar nevykdyta")}, ataskaitos ${esc(st.filings?.date || "dar nevykdyta")}.</span></div><div class="grid">`;
    D.issuers.forEach(i => {
      const secs = D.securities.filter(s => s.issuer === i.id && s.held === "yes");
      const matured = secs.filter(s => s.status === "matured").length, soon = secs.filter(s => s.status === "active" && s.maturity_date && days(s.maturity_date) < 365).length;
      const flagged = recentNews(i.id).filter(n => n.flag === "review").length;
      const chips = [];
      if (i.new_filings.length) chips.push(`<span class="chip bad">● Nauja ataskaita (${i.new_filings.length}) – reikia analizės</span>`);
      if (flagged) chips.push(`<span class="chip warn">${flagged} naujienos peržiūrai (30 d.)</span>`);
      if (matured) chips.push(`<span class="chip warn">${matured} išpirkta – išimti iš sąrašo</span>`);
      if (soon) chips.push(`<span class="chip info">${soon} išperkama per 12 mėn.</span>`);
      if (i.esg_review) chips.push(`<span class="chip info">ESG: ${i.esg_review} peržiūrai</span>`);
      h += `<div class="card"><h3><a href="#emitentas/${i.id}">${esc(i.name)}</a></h3>
        <div class="kv"><span>Šalis, sektorius</span><b>${esc(i.country)} · ${esc(i.sector)}</b><span>Reitingas</span><b>${esc(i.rating.replace(/\[.*?\]/g, ""))}</b>
        <span>Paskutinė analizė</span><b>${esc(i.period)} (${esc(i.analysis_date)})</b><span>Kita ataskaita</span><b>${esc(i.calendar)}</b><span>Priemonės</span><b>${secs.map(s => esc(s.isin)).join(", ")}</b></div>
        <div>${chips.join("") || '<span class="chip good">Naujų įvykių nėra</span>'}</div>
        <p style="margin:10px 0 0">${esc((i.summary[0] || "").replace(/\s*\[.*?\]/g, ""))}</p>
        <div class="row" style="margin-top:10px"><a class="btn" href="#emitentas/${i.id}">Atidaryti</a>${dl(i, true)}</div></div>`;
    });
    return h + "</div>";
  }
  function dl(i, compact) {
    const last = i.files.filter(f => f.period === i.period);
    const name = { docx: "Word", pdf: "PDF", xlsx: "Excel" };
    return last.map(f => `<a class="btn" href="${f.url}" download>⬇ ${name[f.ext]}${compact ? "" : " (" + f.lang.toUpperCase() + ")"}</a>`).join("");
  }
  function issuer(id) {
    const i = byId[id]; if (!i) return "<p>Emitentas nerastas.</p>";
    const secs = D.securities.filter(s => s.issuer === id);
    let h = `<p class="muted"><a href="#apzvalga">← Apžvalga</a></p><h1>${esc(i.name)}</h1><p class="sub">${esc(i.legal_name)} · ${esc(i.sector)} · ${esc(i.owner)}</p>`;
    h += `<div class="row">${dl(i)}<a class="btn" href="${gh("edit/main/" + i.yaml)}" target="_blank" rel="noopener">✎ Taisyti ataskaitą</a>
      <a class="btn" href="${issueUrl("Analizė: " + i.name, "Atnaujinti " + i.name + " apžvalgą pagal naujausią ataskaitą.\n\n(Sukurta iš dashboard)")}" target="_blank" rel="noopener">✦ Paprašyti naujos analizės</a>
      <a class="btn" href="${esc(i.ir_page)}" target="_blank" rel="noopener">Emitento ataskaitos ↗</a></div>`;
    h += `<div class="kv"><span>Analizės data</span><b>${esc(i.analysis_date)} (pagal ${esc(i.period)})</b><span>Reitingas</span><b>${esc(i.rating)}</b><span>Kita ataskaita</span><b>${esc(i.next_report)}</b></div>`;
    if (i.new_filings.length) h += `<div class="note"><b>Rasti nauji dokumentai po paskutinės analizės:</b><ul class="b">${i.new_filings.map(f => `<li><a href="${esc(f.url)}" target="_blank" rel="noopener">${esc(f.title)}</a> <span class="muted">(${esc(f.kind)}, rasta ${esc(f.first_seen)})</span></li>`).join("")}</ul></div>`;
    h += `<h2>Trumpa santrauka</h2><ul class="b">${i.summary.map(s => `<li>${esc(s)}</li>`).join("")}</ul>`;
    h += `<h2>Priemonės</h2><div class="tbl-wrap"><table><tr><th>ISIN</th><th>Priemonė</th><th>Turima</th><th class="n">Kuponas</th><th>Išpirkimas</th><th>Būsena</th><th>Pastaba</th></tr>${secs.map(s => secRow(s)).join("")}</table></div>`;
    // rodikliai
    h += `<h2>Pagrindiniai rodikliai</h2><div class="tbl-wrap"><table><tr><th>Rodiklis</th>${i.periods.map(p => `<th class="n">${p}</th>`).join("")}</tr>`;
    i.ratios.forEach(m => {
      const v = i.history[m]; if (!v) return;
      const u = Object.values(v)[0].u;
      h += `<tr><td>${esc(D.metrics[m] || m)}${unitLbl(u) ? ", " + unitLbl(u) : ""}</td>${i.periods.map(p => {
        const r = v[p]; if (!r) return '<td class="n">–</td>';
        const mark = r.ok === "check" ? "*" : r.ok === "calc" ? "ᶜ" : "";
        return `<td class="n" data-tip="Šaltinis: ${esc(r.s)}${r.pg ? " p." + esc(r.pg) : ""}${r.ok === "check" ? " – reikia patikrinti" : r.ok === "calc" ? " – apskaičiuota" : ""}">${fmt(r.v, u)}${mark}</td>`;
      }).join("")}</tr>`;
    });
    h += `</table></div><p class="muted">* – reikia patikrinti originale; ᶜ – apskaičiuota įrankio. Užveskite pelę ant skaičiaus – matysite šaltinį.</p>`;
    const opts = i.ratios.filter(m => i.history[m]);
    h += `<div class="two"><div><h2>Rodiklio dinamika</h2><div class="filters"><select id="mpick">${opts.map(m => `<option value="${m}">${esc(D.metrics[m] || m)}</option>`).join("")}</select></div><div class="chart" id="mchart"></div></div>`;
    h += `<div><h2>Ankstesni planai ir rezultatai</h2><div class="tbl-wrap"><table><tr><th>Planas</th><th>Rezultatas</th><th>Būsena</th></tr>${i.plans.map(p => `<tr><td>${esc(p.plan)}</td><td>${esc(p.outcome)} <span class="muted">[${esc(p.src)}]</span></td><td>${chip(p.status)}</td></tr>`).join("")}</table></div></div></div>`;
    h += `<h2>Informacija investuotojui ir išvados</h2><ul class="b">${i.investor.map(s => `<li>${esc(s)}</li>`).join("")}</ul>`;
    h += `<h2>ESG ir rankinės patikros</h2><p>${esc(i.esg_conclusion)}</p><div class="tbl-wrap"><table><tr><th>Patikra</th><th>Būsena</th><th>Pastaba</th></tr>
      <tr><td>Sankcijos</td><td>${chip(i.checks.sanctions?.status)}</td><td>${esc(i.checks.sanctions?.note)}</td></tr>
      <tr><td>Žiniasklaida</td><td>${chip(i.checks.media?.status)}</td><td>${esc(i.checks.media?.note)}</td></tr></table></div>
      <p class="muted">Visas ESG klausimynas su įrodymais – Word / PDF faile. Sankcijų ir žiniasklaidos langelius pakeisite mygtuku „Taisyti ataskaitą“ (checks skiltis).</p>`;
    const nw = D.news.filter(n => n.issuer === id).slice(0, 25);
    h += `<h2>Naujienos</h2>${newsTable(nw)}<p class="muted"><a href="#naujienos/${id}">Visos ${esc(i.name)} naujienos →</a></p>`;
    h += `<h2>Šaltiniai</h2><ul class="b">${Object.entries(i.sources).map(([k, s]) => `<li><b>${esc(k)}</b> <a href="${esc(s.url)}" target="_blank" rel="noopener">${esc(s.title)}</a> <span class="muted">${esc(s.date)}</span></li>`).join("")}</ul>`;
    const old = i.files.filter(f => f.period !== i.period);
    if (old.length) h += `<h2>Ankstesnės ataskaitos</h2><div class="row">${old.map(f => `<a class="btn" href="${f.url}" download>${esc(f.period)} ${f.ext}</a>`).join("")}</div>`;
    setTimeout(() => {
      const sel = document.getElementById("mpick"), box = document.getElementById("mchart");
      const draw = () => { const v = i.history[sel.value] || {}; const u = (Object.values(v)[0] || {}).u;
        box.innerHTML = barChart(sortP(Object.keys(v)).map(p => ({ label: p, v: +v[p].v, src: v[p].s + (v[p].pg ? " p." + v[p].pg : "") })), u); };
      sel && (sel.onchange = draw, draw());
    });
    return h;
  }
  const sortP = ps => ps.slice().sort((a, b) => (a.slice(-4) + (a.startsWith("FY") ? "9" : a)).localeCompare(b.slice(-4) + (b.startsWith("FY") ? "9" : b)));
  function secRow(s) {
    const d = s.maturity_date ? days(s.maturity_date) : null;
    const st = s.status === "matured" ? '<span class="chip warn">Išpirkta</span>' : d !== null && d < 365 ? `<span class="chip info">Liko ${d} d.</span>` : '<span class="chip good">Aktyvi</span>';
    return `<tr><td>${esc(s.isin)}</td><td>${esc(s.name)}</td><td>${s.held === "yes" ? "Taip" : "Ne"}</td><td class="n">${s.coupon_pct ? esc(s.coupon_pct) + " %" : "–"}</td>
      <td>${esc(s.maturity_date || (s.type === "share" ? "akcija" : "neterminuota"))}${s.call_date ? `<br><span class="muted">call ${esc(s.call_date)}</span>` : ""}</td><td>${st}</td>
      <td class="muted">${s.verified === "review" ? "⚠ " : ""}${esc(s.status_note)}</td></tr>`;
  }
  function newsTable(rows) {
    if (!rows.length) return '<p class="muted">Naujienų dar nėra – jos atsiras po pirmos automatinės paieškos (mygtukas „Ieškoti naujų ataskaitų ir naujienų dabar“).</p>';
    return `<div class="tbl-wrap"><table><tr><th>Data</th><th>Emitentas</th><th>Antraštė</th><th>Šaltinis</th><th>Žyma</th></tr>${rows.map(n => `<tr><td style="white-space:nowrap">${esc(n.date)}</td><td>${esc(byId[n.issuer]?.name || n.issuer)}</td>
      <td><a href="${esc(n.url)}" target="_blank" rel="noopener">${esc(n.title)}</a>${n.topics ? ` <span class="muted">${esc(n.topics.replace(/;/g, ", "))}</span>` : ""}</td><td>${esc(n.source)} <span class="muted">${esc(n.lang)}</span></td>
      <td>${n.flag === "review" ? `<span class="chip warn" data-tip="Rasti žodžiai: ${esc(n.keywords)}">Peržiūrėti</span>` : ""}</td></tr>`).join("")}</table></div>`;
  }
  function news(sel) {
    let h = `<h1>Naujienos</h1><p class="sub">Google News ir regiono žiniasklaida emitento šalies kalba(-omis). „Peržiūrėti“ – antraštėje rasta neigiamų žodžių (tyrimas, bauda, teismas, sankcijos ir pan.). Tai pagalba, ne išvada.</p>
      <div class="filters"><select id="fi"><option value="">Visi emitentai</option>${D.issuers.map(i => `<option value="${i.id}" ${i.id === sel ? "selected" : ""}>${esc(i.name)}</option>`).join("")}</select>
      <select id="ff"><option value="">Visos</option><option value="review">Tik peržiūrai</option></select><input type="search" id="fq" placeholder="Paieška antraštėse"></div><div id="nl"></div>`;
    setTimeout(() => {
      const f = () => { const i = fi.value, fl = ff.value, q = fq.value.toLowerCase();
        nl.innerHTML = newsTable(D.news.filter(n => (!i || n.issuer === i) && (!fl || n.flag === fl) && (!q || n.title.toLowerCase().includes(q))).slice(0, 300)); };
      [fi, ff, fq].forEach(e => e.oninput = f); f();
    });
    return h;
  }
  function securities() {
    const s = D.securities.slice().sort((a, b) => (a.maturity_date || "9999").localeCompare(b.maturity_date || "9999"));
    const held = s.filter(x => x.held === "yes" && x.type === "bond" && x.status === "active" && x.maturity_date);
    const years = {}; held.forEach(x => { const y = x.maturity_date.slice(0, 4); years[y] = (years[y] || 0) + 1; });
    const matured = s.filter(x => x.status === "matured" && x.held === "yes");
    let h = `<h1>Vertybiniai popieriai</h1><p class="sub">Sąrašas – <a href="${gh("edit/main/config/securities.csv")}" target="_blank" rel="noopener">config/securities.csv</a> (taisomas GitHub'e). Kai išpirkimo data praeina, popierius automatiškai pažymimas „Išpirkta“ ir nebeįtraukiamas į ataskaitas.</p>`;
    if (matured.length) h += `<div class="note">Išpirkti, bet dar pažymėti kaip turimi: ${matured.map(x => esc(x.isin)).join(", ")}. Pakeiskite <b>held</b> į <b>no</b> arba ištrinkite eilutę.</div>`;
    h += `<div class="two"><div><h2>Turimų obligacijų išpirkimai pagal metus</h2><div class="chart">${barChart(Object.keys(years).sort().map(y => ({ label: y, v: years[y] })), "")}</div><p class="muted">Obligacijų skaičius (be neterminuotų).</p></div><div></div></div>`;
    h += `<h2>Visi</h2><div class="tbl-wrap"><table><tr><th>ISIN</th><th>Emitentas / priemonė</th><th>Turima</th><th class="n">Kuponas</th><th>Išpirkimas</th><th>Būsena</th><th>Pastaba</th></tr>${s.map(x => secRow({ ...x, name: (byId[x.issuer]?.name || x.issuer) + " – " + x.name })).join("")}</table></div>`;
    return h;
  }
  function tools() {
    const all = [...new Set(D.issuers.flatMap(i => Object.keys(i.history)))];
    let h = `<h1>Analizės įrankiai</h1><p class="sub">Palyginimai iš visų emitentų rodiklių istorijos (data/history/*.csv). Istorija pildosi su kiekviena nauja analize.</p>`;
    h += `<h2>Paieška ir analizė</h2><div class="row"><a class="btn primary" href="${ACTIONS}" target="_blank" rel="noopener">↻ Ieškoti naujų ataskaitų ir naujienų (nemokama)</a>
      <a class="btn" href="${issueUrl("Analizė: ", "Kurį emitentą analizuoti ir pagal kokią ataskaitą?\n\n(Sukurta iš dashboard)")}" target="_blank" rel="noopener">✦ Paprašyti AI analizės</a></div>
      <p class="muted">„Ieškoti“ atidaro GitHub Actions: spauskite „Run workflow“. AI analizė – žr. „Kaip naudotis“.</p>`;
    h += `<h2>Rodiklis visiems emitentams (naujausia reikšmė)</h2><div class="filters"><select id="cm">${all.map(m => `<option value="${m}">${esc(D.metrics[m] || m)}</option>`).join("")}</select></div><div id="ct"></div>`;
    setTimeout(() => {
      const f = () => { const m = cm.value;
        const rows = D.issuers.map(i => { const v = i.history[m]; if (!v) return null; const p = sortP(Object.keys(v)).pop(); return { i, p, r: v[p] }; }).filter(Boolean);
        ct.innerHTML = rows.length ? `<div class="two"><div class="tbl-wrap"><table><tr><th>Emitentas</th><th>Laikotarpis</th><th class="n">Reikšmė</th><th>Šaltinis</th></tr>${rows.map(x => `<tr><td><a href="#emitentas/${x.i.id}">${esc(x.i.name)}</a></td><td>${esc(x.p)}</td><td class="n">${fmt(x.r.v, x.r.u)} ${unitLbl(x.r.u)}</td><td class="muted">${esc(x.r.s)}${x.r.pg ? " p." + esc(x.r.pg) : ""}</td></tr>`).join("")}</table></div>
          <div class="chart">${barChart(rows.map(x => ({ label: x.i.name, v: +x.r.v, src: x.p })), rows[0].r.u)}</div></div>` : "";
      };
      cm.onchange = f; f();
    });
    return h;
  }
  function help() {
    return `<h1>Kaip naudotis</h1>
    <div class="card"><h3>Kas vyksta automatiškai (nemokamai)</h3><ul class="b">
      <li>Kiekvieną pirmadienį GitHub Actions surenka emitentų naujienas (Google News, emitento šalies kalba) ir patikrina emitentų ataskaitų puslapius – ar atsirado naujų dokumentų. Tai be AI, todėl nieko nekainuoja.</li>
      <li>Radus naują ataskaitą, emitento kortelėje atsiranda raudona žyma „Nauja ataskaita – reikia analizės“.</li>
      <li>Kai obligacijos išpirkimo data praeina, ji pažymima „Išpirkta“ ir išimama iš ataskaitų.</li>
      <li>Bet kada galite paleisti paiešką patys: mygtukas „Ieškoti naujų ataskaitų ir naujienų dabar“ → „Run workflow“.</li></ul></div>
    <div class="card" style="margin-top:12px"><h3>AI analizė (kai pasirodo nauja ataskaita)</h3><ul class="b">
      <li>Paprasčiausia: parašykite Claude projekte, pvz. „Atnaujink Ignitis apžvalgą pagal 9 mėn. ataskaitą“. Naudojamas jūsų Claude planas, papildomai nemokama.</li>
      <li>Arba mygtukas „Paprašyti naujos analizės“ – sukuria GitHub užduotį. Jei įjungsite Claude GitHub integraciją (žr. README), analizė bus atlikta automatiškai.</li></ul></div>
    <div class="card" style="margin-top:12px"><h3>Taisymas</h3><ul class="b">
      <li>„Taisyti ataskaitą“ atidaro ataskaitos tekstą GitHub'e – pakeiskite ir išsaugokite (Commit). Word, PDF ir Excel persigeneruos per kelias minutes.</li>
      <li>Atsisiųstą Word failą galite laisvai redaguoti savo kompiuteryje.</li>
      <li>Naują emitentą ar ISIN pridėkite: config/issuers.yaml ir config/securities.csv.</li></ul></div>`;
  }

  function route() {
    const [page, arg] = (location.hash.slice(1) || "apzvalga").split("/");
    document.querySelectorAll("#nav a").forEach(a => a.classList.toggle("on", a.getAttribute("href") === "#" + page || (page === "emitentas" && a.getAttribute("href") === "#apzvalga")));
    $app.innerHTML = page === "emitentas" ? issuer(arg) : page === "naujienos" ? news(arg) : page === "vp" ? securities() : page === "irankiai" ? tools() : page === "pagalba" ? help() : overview();
    scrollTo(0, 0);
  }
  addEventListener("hashchange", route);
  document.getElementById("foot").innerHTML = `Duomenys iš viešų šaltinių; skaičiai – kaip skelbia emitentai. Sugeneruota ${esc(D.generated)}.`;
  route();
})();
