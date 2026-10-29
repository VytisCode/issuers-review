"""Oficialūs fondų dokumentai (faktų lapai, KID, valdytojų puslapiai) -> šaka „fund-docs“ (PDF + tekstas, datuota kopija).
Paleidžiama rankiniu būdu (fund-docs.yml, įvestis – JSON sąrašas), nemokama, be AI.
Įvestis: [{"isin": .., "kind": "factsheet|kid|page", "url": ..}] ir/arba {"isins": [..]} – tada bandomi standartiniai adresai pagal ISIN."""
import hashlib
import json
import re
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"
PATTERNS = [("factsheet", "https://docs.extraetf.com/download/{i}/en/MR/"),
            ("factsheet", "https://www.finanzfluss.de/document/{i}/de/factsheet"),
            ("factsheet", "https://assets.finanzfluss.de/informer/documents/{i}-factsheet.pdf"),
            ("kid", "https://www.finanzfluss.de/document/{i}/de/kiid"),
            ("kid", "https://docs.extraetf.com/download/{i}/en/KID/")]


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*", "Accept-Language": "en"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read(), r.headers.get("Content-Type", ""), r.geturl()


def text_of(path):
    if path.suffix == ".pdf":
        out = []
        n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", str(path)], capture_output=True, text=True).stdout or "Pages: 0").group(1))
        for p in range(1, n + 1):
            t = subprocess.run(["pdftotext", "-layout", "-f", str(p), "-l", str(p), str(path), "-"], capture_output=True, text=True).stdout
            out.append(f"=== pdf p.{p} ===\n{t}")
        return "\n".join(out)
    if path.suffix == ".csv":
        return path.read_text(errors="ignore")
    html = path.read_text(errors="ignore")
    html = re.sub(r"(?is)<(script|style).*?</\1>", " ", html)
    return re.sub(r"[ \t]+", " ", re.sub(r"<[^>]+>", " ", html))


def run(out, spec):
    out = Path(out); day = time.strftime("%Y-%m-%d")
    items = list(spec.get("items", []))
    for i in spec.get("isins", []):
        items += [{"isin": i, "kind": k, "url": u.format(i=i)} for k, u in PATTERNS]
    man_f = out / "manifest.json"
    man = json.loads(man_f.read_text()) if man_f.exists() else []
    for it in items:
        rec = dict(it, fetched=day)
        try:
            data, ctype, final = fetch(it["url"])
            pdf = data[:5] == b"%PDF-"
            if not pdf and it["kind"] not in ("page", "data"):
                raise ValueError(f"ne PDF ({ctype}, {len(data)} B)")
            h = hashlib.sha256(data).hexdigest()
            d = out / day / it["isin"]; d.mkdir(parents=True, exist_ok=True)
            fn = d / f"{it['kind']}_{h[:10]}{'.pdf' if pdf else ('.csv' if it['kind'] == 'data' else '.html')}"
            fn.write_bytes(data)
            (fn.with_suffix(".txt")).write_text(text_of(fn))
            rec.update(ok=True, file=str(fn.relative_to(out)), sha256=h, bytes=len(data), final_url=final)
            print("OK  ", it["isin"], it["kind"], it["url"])
        except Exception as e:  # noqa: BLE001
            rec.update(ok=False, error=str(e)[:200])
            print("FAIL", it["isin"], it["kind"], it["url"], str(e)[:120])
        man.append(rec)
        time.sleep(1)
    man_f.write_text(json.dumps(man, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    run(sys.argv[1], json.loads(Path(sys.argv[2]).read_text()))
