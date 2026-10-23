"""ESG šaltiniai: atsisiunčia emitentų metines / tvarumo ataskaitas (data/esg/docs.yaml) ir išsaugo jų tekstą puslapiais.

Kam: ESG polapio analizei reikia viso ataskaitos teksto su puslapių numeriais (tvarumo dalis dažnai yra 100+ psl. gilumoje).
Paleidžiama GitHub Actions (workflow „ESG documents“), nes ten pasiekiami visi emitentų tinklapiai. Tekstas dedamas į atskirą
šaką „esg-text“ (ne į main - kad svetainės saugykla neišpūstų). Puslapių numeriai - PDF puslapiai („pdf p.N“), kaip nuorodose #page=N.
Kiekvienam dokumentui galima nurodyti kelis adresus: imamas pirmas veikiantis; jei adresas - tinklapis, jame ieškoma PDF nuorodų
su tų metų skaičiumi ir žodžiais annual / sustainability / report. Be AI.

Šakoje: text/ - tekstas su išdėstymu (lentelėms), raw/ - skaitymo tvarka (stulpeliams).
Naudojimas: python tools/fetch_esg_docs.py [išvesties katalogas, numatytas build/esg-text]
"""
import base64
import html as _html
import io
import json
import re
import zipfile
import subprocess
import sys
import urllib.parse
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"
MAX_BYTES = 150 * 1024 * 1024
WORDS = r"annual|sustainab|integrated|report|ataskait|aruanne|parskats|pārskats|ziņojums|esg|tvarum|metin"


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read(MAX_BYTES), r.headers.get("Content-Type", ""), r.geturl()


def pdf_links(html, base, year):
    links = re.findall(r'href=["\']([^"\']+\.pdf[^"\']*)["\']', html, re.I)
    # Dragonfly tipo nuorodos (pvz. cblgroup.com/media/W1si...): failo vardas užkoduotas base64
    links += re.findall(r'href=["\']([^"\']*/media/W1[A-Za-z0-9_=-]+[^"\']*)["\']', html)
    out = []
    for h in links:
        u = urllib.parse.urljoin(base, h)
        low = urllib.parse.unquote(u).lower()
        m = re.search(r"/media/(W1[A-Za-z0-9_=-]+)", u)
        if m:
            try:
                low += " " + base64.urlsafe_b64decode(m.group(1) + "=" * (-len(m.group(1)) % 4)).decode("utf-8", "ignore").lower()
            except Exception:  # noqa: BLE001
                pass
            if ".pdf" not in low:
                continue
        if year in low and re.search(WORDS, low) and u not in out:
            out.append(u)
    # tvarumo / metinės ataskaitos pirmiau nei ketvirtinės
    out.sort(key=lambda u: (not re.search(r"sustainab|esg|tvarum|annual|integrated|metin|aastaaruanne|gada", u.lower()), re.search(r"q[1-4]|quarter|interim|ketv", u.lower()) is not None))
    return out[:6]


def get_pdf(cands, year, log):
    for url in cands:
        try:
            data, ctype, final = fetch(url)
        except Exception as e:  # noqa: BLE001
            log.append(f"  ✗ {url}: {e}")
            continue
        if data[:5] == b"%PDF-" or data[:2] == b"PK":
            return data, final
        if re.search(r"\.x?html?$", final.split("?")[0], re.I) and b"ix:header" in data[:400000]:
            return data, final  # ESEF ataskaita kaip vienas XHTML failas (pvz. filings.xbrl.org)
        if b"<html" in data[:2000].lower() or "html" in ctype:
            for u in pdf_links(data.decode("utf-8", "ignore"), final, year):
                try:
                    d2, _, f2 = fetch(u)
                    if d2[:5] == b"%PDF-":
                        log.append(f"  → rasta tinklapyje {url}")
                        return d2, f2
                except Exception as e:  # noqa: BLE001
                    log.append(f"  ✗ {u}: {e}")
        log.append(f"  ✗ {url}: ne PDF")
    return None, None


def garbled(text):
    """Ar PDF tekstas neįskaitomas (šriftai be Unicode lentelės): per mažai raidžių tarp ne tarpo simbolių."""
    chars = [c for c in text[:400000] if not c.isspace()]
    return len(chars) > 2000 and sum(c.isalpha() for c in chars) / len(chars) < 0.55


def blank_pages(pages):
    """Puslapiai be teksto sluoksnio (paveikslėliai, pvz. Inbank vadovybės ataskaita): mažiau nei 8 žodžiai."""
    return [n for n, p in enumerate(pages, 1) if len(p.split()) < 8]


def ocr_pages(pdf, lang="eng", only=None):
    """Neįskaitomo PDF puslapiai per OCR (tesseract, nemokamas). Lėta (~2-4 s puslapiui), todėl tik kai tekstas neįskaitomas.
    only - tik šie puslapių numeriai (kiti grąžinami tušti)."""
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout).group(1))
    out = []
    for i in range(1, n + 1):
        if only is not None and i not in only:
            out.append("")
            continue
        png = subprocess.run(["pdftoppm", "-r", "200", "-f", str(i), "-l", str(i), "-gray", "-png", str(pdf)], capture_output=True).stdout
        out.append(subprocess.run(["tesseract", "-", "-", "-l", lang, "--psm", "4"], input=png, capture_output=True).stdout.decode("utf-8", "ignore"))
    return out


def esef_pages(data):
    """ESEF ZIP (XHTML) arba vienas XHTML: didžiausias .xhtml/.html failas -> puslapiai pagal „page-break“ žymes (PDF nėra, todėl nuoroda be #page)."""
    if data[:2] == b"PK":
        z = zipfile.ZipFile(io.BytesIO(data))
        names = [n for n in z.namelist() if n.lower().endswith((".xhtml", ".html", ".htm"))]
        if not names:
            return [""]
        doc = z.read(max(names, key=lambda n: z.getinfo(n).file_size)).decode("utf-8", "ignore")
    else:
        doc = data.decode("utf-8", "ignore")
    parts = re.split(r"<(?:div|p|hr)[^>]*page-break-(?:after|before)[^>]*>", doc, flags=re.I) if "page-break" in doc else re.split(r'<div[^>]*class="[^"]*\bpage\b[^"]*"[^>]*>', doc)
    out = []
    for part in parts:
        t = re.sub(r"<(?:br|/p|/div|/tr|/h\d)[^>]*>", "\n", part, flags=re.I)
        t = re.sub(r"<[^>]+>", " ", t)
        t = _html.unescape(re.sub(r"[ \t\xa0]+", " ", t))
        out.append(re.sub(r"\n\s*\n+", "\n", t))
    return out


def main(out_dir):
    out = Path(out_dir)
    (out / "text").mkdir(parents=True, exist_ok=True)
    (out / "raw").mkdir(parents=True, exist_ok=True)
    docs = C.load_yaml(C.DATA / "esg" / "docs.yaml")
    man_path = out / "manifest.json"
    manifest = json.loads(man_path.read_text()) if man_path.exists() else {}
    log = []
    for issuer, years in docs.items():
        for year, items in (years or {}).items():
            for i, item in enumerate(items or []):
                cands = item.get("urls", []) if isinstance(item, dict) else [item]
                name = f"{issuer}_{year}" + (f"_{i + 1}" if i else "")
                done = manifest.get(name, {}).get("pages")
                txt_ok = (out / "text" / f"{name}.txt").exists() and (out / "text" / f"{name}.txt").stat().st_size > 2000
                esef_ok = manifest.get(name, {}).get("format") == "esef" and manifest[name].get("url") == cands[0] and txt_ok
                bad_txt = txt_ok and not manifest[name].get("ocr") and garbled((out / "text" / f"{name}.txt").read_text(encoding="utf-8", errors="ignore"))
                raw_p = out / "raw" / f"{name}.txt"
                no_text = (raw_p.exists() and not manifest[name].get("ocr") and not manifest[name].get("ocr_blank")
                           and len(blank_pages(re.split(r"\n=== pdf p\.\d+ ===\n", raw_p.read_text(encoding="utf-8", errors="ignore"))[1:])) >= 5)
                if done and not bad_txt and not no_text and (esef_ok or (out / "raw" / f"{name}.txt").exists()):
                    continue
                log.append(f"{name}:")
                data, url = get_pdf(cands, year[-4:], log)
                if not data:
                    manifest[name] = {"issuer": issuer, "year": year, "url": "", "error": "nepavyko atsisiųsti"}
                    continue
                esef = data[:5] != b"%PDF-"
                ocr = False
                if esef:
                    pages = esef_pages(data)
                else:
                    pdf = out / "pdf.tmp"
                    pdf.write_bytes(data)
                    txt = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True).stdout.decode("utf-8", "ignore")
                    pages = txt.split("\f")
                    # antra versija skaitymo tvarka (be -layout): kelių stulpelių puslapiai skaitomi stulpelis po stulpelio
                    raw = subprocess.run(["pdftotext", str(pdf), "-"], capture_output=True).stdout.decode("utf-8", "ignore").split("\f")
                    ocr = garbled(txt)
                    lang = item.get("ocr_lang", "eng") if isinstance(item, dict) else "eng"
                    blank = [] if ocr else [b for b in blank_pages(raw) if b < len(raw)]
                    if ocr:
                        log.append("  … tekstas neįskaitomas, daromas OCR")
                        pages = raw = ocr_pages(pdf, lang)
                    elif len(blank) >= 5:
                        # dalis puslapių - tik paveikslėliai: OCR tik jiems
                        log.append(f"  … {len(blank)} psl. be teksto, daromas OCR tik jiems")
                        o = ocr_pages(pdf, lang, set(blank))
                        pages = [o[i] if (i + 1) in blank and i < len(o) else p for i, p in enumerate(pages)]
                        raw = [o[i] if (i + 1) in blank and i < len(o) else p for i, p in enumerate(raw)]
                    (out / "raw" / f"{name}.txt").write_text("".join(f"\n=== pdf p.{n} ===\n{p}" for n, p in enumerate(raw, 1)), encoding="utf-8")
                    pdf.unlink()
                body = "".join(f"\n=== pdf p.{n} ===\n{p}" for n, p in enumerate(pages, 1) if p.strip() or n < len(pages))
                (out / "text" / f"{name}.txt").write_text(body, encoding="utf-8")
                title = item.get("title", "") if isinstance(item, dict) else ""
                manifest[name] = {"issuer": issuer, "year": year, "url": url, "title": title, "pages": len(pages), "kb": len(data) // 1024}
                if esef:
                    manifest[name]["format"] = "esef"
                elif ocr:
                    manifest[name]["ocr"] = True
                if not esef and not ocr:
                    manifest[name]["ocr_blank"] = len(blank)
                log.append(f"  ✓ {url} ({len(pages)} psl.)")
    man_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    print("\n".join(log))
    print(f"Iš viso dokumentų: {sum(1 for m in manifest.values() if m.get('pages'))}, nepavyko: {sum(1 for m in manifest.values() if not m.get('pages'))}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else C.ROOT / "build" / "esg-text")
