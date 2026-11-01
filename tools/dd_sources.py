"""DD dokumentacijos ZIP (GitHub Actions, update.yml): data/dd/jobs/<DD id>.json -> atsisiunčia viešus šaltinius,
supakuoja į ZIP, užšifruoja viešuoju DD raktu (tools/dd_vault.py) ir prideda į DD archyvą kaip {type: zip, parent}.
Užduotis lieka faile (done = ZIP id), kad būtų matyti, kas jau padaryta. Be AI, nemokamai.

Šaltinių sąrašas ir ZIP sutampa 1:1 (savininkas 2026-11-01): kiekvienas raktas iš DD šaltinių sąrašo turi failą ZIP'e:
- {key, title, url}            - atsisiunčiama iš interneto;
- {key, title, received}       - savininko pateiktas dokumentas (D#, taip pat KIID / prospektas be viešos nuorodos): `dd_vault.py
                                  sources` jį iš karto užšifruoja atskirame ZIP {type: zip, kind: received}; dashboard'as atsisiunčiant
                                  sujungia abu į vieną „Dokumentai (ZIP)“ (savininko sprendimas 2026-11-01: visi D failai - į ZIP);
- {key, title, note}           - į ZIP nededama - ZIP'e <key>_..._NOT_INCLUDED.txt su priežastimi.
Turinio patikra (2026-11-01, AQR DD: į ZIP pateko svetainės pradinis puslapis ir prisijungimo langas vietoj KIID / straipsnio):
atsisiuntimo kodas 200 dar nereiškia, kad tai cituojamas dokumentas. `check()` atmeta svetainės pradinį puslapį, prisijungimo,
mokamo turinio, robotų blokavimo ir „nerasta“ puslapius, beveik tuščius failus ir failus be laukiamo teksto
(`expect` sąrašas; ISIN iš pavadinimo tikrinamas automatiškai). Atmesti į ZIP nededami - _saltiniai.txt nurodoma priežastis.
"""
import io
import json
import re
import sys
import time
import zipfile
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))
import archive_sources as A  # noqa: E402
import dd_vault as V  # noqa: E402

ZIPS = V.DIR / "zips.json"
ISIN = re.compile(r"\b[A-Z]{2}[A-Z0-9]{9}[0-9]\b")
BAD = [
    (r"you must (?:be )?log(?:ged)? ?in|(?:log|sign) ?in to (?:view|continue|read|access)|please (?:log|sign) ?in\b|login required", "prisijungimo puslapis"),
    (r"subscribe (?:now )?to (?:continue|read)|for subscribers only|subscriber[- ]only|already a subscriber\?|to continue reading", "mokamas turinys (paywall)"),
    (r"are you a robot|verify (?:that )?you are (?:a )?human|captcha|access denied|enable javascript and cookies to continue", "robotų blokavimas"),
    (r"\b404\b.{0,40}not found|page (?:was )?not found|page you requested (?:could not|cannot) be found", "puslapis nerastas"),
]


def text_of(data, ext):
    """Tekstas patikrai: PDF - pypdf (iki 40 psl.), HTML - be žymų ir skriptų, kita - tuščia (netikrinama)."""
    if ext == ".pdf" or data[:5] == b"%PDF-":
        try:
            from pypdf import PdfReader
            r = PdfReader(io.BytesIO(data))
            return " ".join((p.extract_text() or "") for p in r.pages[:40])
        except Exception:  # noqa: BLE001
            return None
    if ext in (".html", ".htm", ".txt", ".csv", ".json"):
        t = data.decode("utf-8", "ignore")
        if ext in (".html", ".htm"):
            t = re.sub(r"(?is)<(script|style|noscript)\b.*?</\1>", " ", t)
            t = re.sub(r"(?s)<[^>]+>", " ", t)
        return t
    return None


def check(s, data, ext):
    """Grąžina priežastį, kodėl failas NĖRA cituojamas dokumentas, arba None, jei tinka."""
    if s.get("url"):
        u = urlparse(s["url"])
        if u.path.strip("/") == "" and not u.query and not s.get("homepage_ok"):
            return "svetainės pradinis puslapis, ne dokumentas (reikia tiesioginės dokumento nuorodos)"
    t = text_of(data, ext)
    if t is None:
        return None  # xlsx, zip ir pan. - turinio netikriname
    flat = re.sub(r"\s+", " ", t).strip()
    if len(flat) < 200:
        return f"beveik tuščias ({len(flat)} simb.)"
    low = flat.lower()
    for rx, why in BAD:
        if re.search(rx, low[:20000]) and len(flat) < 15000:
            return why
    need = list(s.get("expect") or []) + ISIN.findall(s.get("title", ""))
    miss = [w for w in need if re.sub(r"\s+", " ", w).lower() not in low]
    if miss:
        return "nėra laukiamo teksto: " + ", ".join(miss)
    return None


def run():
    jobs = sorted((V.DIR / "jobs").glob("*.json")) if (V.DIR / "jobs").exists() else []
    zips = json.loads(ZIPS.read_text()) if ZIPS.exists() else {}
    for jf in jobs:
        job = json.loads(jf.read_text())
        if job.get("done"):
            continue
        buf, lines, ok, bad = io.BytesIO(), [job["name"], f"Atsisiųsta {time.strftime('%Y-%m-%d')}", ""], 0, []
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for s in job["sources"]:
                line = f"[{s['key']}] {s['title']}" + (f" - {s['url']}" if s.get("url") else "")
                if s.get("note"):  # konfidencialus gautas dokumentas: tik nuoroda, kur jį rasti
                    fn = f"{s['key']}_{A.safe(s['title'], 60)}_NOT_INCLUDED.txt"
                    z.writestr(fn, f"[{s['key']}] {s['title']}\n\n{s['note']}\n")
                    lines.append(line + f"  -> {fn} (neįtraukta: {s['note']})")
                    continue
                if s.get("received"):  # gautas dokumentas: failas yra užšifruotame gautų dokumentų ZIP (dd_vault.py sources)
                    lines.append(line + f"  -> {s['received']} (gautas dokumentas; dashboard'as prideda jį į šį ZIP atsisiuntimo metu)")
                    continue
                if any(d in s["url"] for d in A.SKIP) or not s["url"].startswith("http"):
                    lines.append(line + "  (neatsisiunčiama)")
                    continue
                try:
                    data, ctype = A.fetch(s["url"])
                    ext = A.ext_for(s["url"], ctype, data)
                    why = check(s, data, ext)
                    if why:
                        lines.append(line + f"  (NETINKA - neįtraukta: {why})"); bad.append((s["key"], why))
                        continue
                    data, ext = A.as_pdf(s["url"], data, ext)
                    why = check(s, data, ext) if ext == ".pdf" else None  # Chrome vaizdas gali skirtis nuo HTML
                    if why:
                        lines.append(line + f"  (NETINKA - neįtraukta: {why})"); bad.append((s["key"], why))
                        continue
                    fn = f"{s['key']}_{A.safe(s['title'], 60)}{ext}"
                    z.writestr(fn, data)
                    lines.append(line + f"  -> {fn}")
                    ok += 1
                except Exception as e:  # noqa: BLE001
                    lines.append(line + f"  (nepavyko: {e})"); bad.append((s["key"], str(e)))
                time.sleep(1)
            z.writestr("_saltiniai.txt", "\n".join(lines) + "\n")
        zid = V.add_bytes(buf.getvalue(), {"type": "zip", "parent": job["parent"], "file": job["name"], "docs": ok})
        job["problems"] = [{"key": k, "why": w} for k, w in bad]
        zips[zid] = job["parent"]
        # tik naujausias dokumentų ZIP prie įrašo (savininkas 2026-10-29: du „Dokumentai (ZIP)“ mygtukai)
        ev = V.EVID()
        stale = [z for z, p in zips.items() if p == job["parent"] and z != zid and z not in ev and z not in V.RECEIVED()]
        if stale:
            V.INDEX.write_text(json.dumps([x for x in json.loads(V.INDEX.read_text()) if x["id"] not in stale], indent=1) + "\n")
            for z in stale:
                zips.pop(z, None); (V.DIR / f"{z}.bin").unlink(missing_ok=True)
        job["done"] = zid
        jf.write_text(json.dumps(job, ensure_ascii=False, indent=1))
        print(f"{job['name']}: {ok}/{len(job['sources'])} dokumentų -> {zid}")
        for k, w in bad:
            print(f"  NETINKA {k}: {w}")
    ZIPS.write_text(json.dumps(zips, indent=1) + "\n")


if __name__ == "__main__":
    run()
