"""ESG šaltiniai: atsisiunčia emitentų metines / tvarumo ataskaitas (data/esg/docs.yaml) ir išsaugo jų tekstą puslapiais.

Kam: ESG polapio analizei reikia viso ataskaitos teksto su puslapių numeriais (tvarumo dalis dažnai yra 100+ psl. gilumoje).
Paleidžiama GitHub Actions (workflow „ESG documents“), nes ten pasiekiami visi emitentų tinklapiai. Tekstas dedamas į atskirą
šaką „esg-text“ (ne į main - kad svetainės saugykla neišpūstų). Puslapių numeriai - PDF puslapiai („pdf p.N“), kaip nuorodose #page=N.
Kiekvienam dokumentui galima nurodyti kelis adresus: imamas pirmas veikiantis; jei adresas - tinklapis, jame ieškoma PDF nuorodų
su tų metų skaičiumi ir žodžiais annual / sustainability / report. Be AI.

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


def esef_pages(data):
    """ESEF ZIP (XHTML): didžiausias .xhtml/.html failas -> puslapiai pagal „page-break“ žymes (PDF nėra, todėl nuoroda be #page)."""
    z = zipfile.ZipFile(io.BytesIO(data))
    names = [n for n in z.namelist() if n.lower().endswith((".xhtml", ".html", ".htm"))]
    if not names:
        return [""]
    doc = z.read(max(names, key=lambda n: z.getinfo(n).file_size)).decode("utf-8", "ignore")
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
    docs = C.load_yaml(C.DATA / "esg" / "docs.yaml")
    man_path = out / "manifest.json"
    manifest = json.loads(man_path.read_text()) if man_path.exists() else {}
    log = []
    for issuer, years in docs.items():
        for year, items in (years or {}).items():
            for i, item in enumerate(items or []):
                cands = item.get("urls", []) if isinstance(item, dict) else [item]
                name = f"{issuer}_{year}" + (f"_{i + 1}" if i else "")
                if manifest.get(name, {}).get("pages"):
                    continue
                log.append(f"{name}:")
                data, url = get_pdf(cands, year[-4:], log)
                if not data:
                    manifest[name] = {"issuer": issuer, "year": year, "url": "", "error": "nepavyko atsisiųsti"}
                    continue
                if data[:2] == b"PK":
                    pages = esef_pages(data)
                else:
                    pdf = out / "pdf.tmp"
                    pdf.write_bytes(data)
                    txt = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True).stdout.decode("utf-8", "ignore")
                    pages = txt.split("\f")
                    pdf.unlink()
                body = "".join(f"\n=== pdf p.{n} ===\n{p}" for n, p in enumerate(pages, 1) if p.strip() or n < len(pages))
                (out / "text" / f"{name}.txt").write_text(body, encoding="utf-8")
                title = item.get("title", "") if isinstance(item, dict) else ""
                manifest[name] = {"issuer": issuer, "year": year, "url": url, "title": title, "pages": len(pages), "kb": len(data) // 1024}
                if data[:2] == b"PK":
                    manifest[name]["format"] = "esef"
                log.append(f"  ✓ {url} ({len(pages)} psl.)")
    man_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    print("\n".join(log))
    print(f"Iš viso dokumentų: {sum(1 for m in manifest.values() if m.get('pages'))}, nepavyko: {sum(1 for m in manifest.values() if not m.get('pages'))}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else C.ROOT / "build" / "esg-text")
