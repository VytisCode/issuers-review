"""Šaltinių paketas: kiekvienai ataskaitai atsisiunčia jos šaltinių sąrašo dokumentus (PDF, tinklapius) į vieną ZIP.

Kam: jei emitento tinklapis ateityje nebeveiks, pagrindinės ataskaitos lieka išsaugotos.
ZIP dedamas į šios saugyklos GitHub Release „sources“ (ne į git istoriją - kad saugykla neišpūstų), nemokamai.
Sąrašas, kurie paketai jau sukurti: data/sources_archive.json. Sukurtas paketas nebeperkuriamas (archyvas = to meto dokumentai);
nauja analizė (naujas reports/<emitentas>/<laikotarpis>.yaml) = naujas paketas.
Google News nuorodos neatsisiunčiamos (tik įrašomos į sąrašą). Be AI.

Naudojimas: python tools/archive_sources.py            - sukuria trūkstamus paketus ir įkelia (GitHub Actions, reikia GH_TOKEN)
            python tools/archive_sources.py --local    - tik sukuria ZIP į build/sources/ (be įkėlimo)
"""
import re
import subprocess
import sys
import time
import urllib.request
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

TAG = "sources"
UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36 issuers-review-archive"
MAX_BYTES = 80 * 1024 * 1024
OUT = C.ROOT / "build" / "sources"
STATE = C.DATA / "sources_archive.json"
SKIP = ("news.google.com",)


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    with urllib.request.urlopen(req, timeout=90) as r:
        data = r.read(MAX_BYTES + 1)
        if len(data) > MAX_BYTES:
            raise ValueError("per didelis (> 80 MB)")
        return data, r.headers.get("Content-Type", "")


def ext_for(url, ctype, data):
    if data[:5] == b"%PDF-" or "pdf" in ctype:
        return ".pdf"
    if "html" in ctype or data[:200].lstrip().lower().startswith((b"<!doctype html", b"<html")):
        return ".html"
    m = re.search(r"\.(xlsx|xls|docx|csv|zip|json)(?:$|\?)", url.lower())
    return "." + m.group(1) if m else ".bin"


def safe(s, n=70):
    s = re.sub(r"[^\w\-. ]+", "", s, flags=re.UNICODE).strip().replace(" ", "_")
    return s[:n] or "doc"


def build_zip(issuer, rep):
    name = f"{issuer}_{rep['period']}_sources.zip"
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / name
    lines = [f"{C.issuers()[issuer]['name']} – {rep['period']} (analizė {rep['analysis_date']})", f"Atsisiųsta {C.TODAY}", ""]
    ok = 0
    seen = {}
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for key, s in rep.get("sources", {}).items():
            url = (s.get("url") or "").split("#")[0]
            title = s.get("title", key)
            row = f"[{key}] {title} | {s.get('date', '')} | {url}"
            if not url or any(x in url for x in SKIP):
                lines.append(row + " | neatsisiųsta (nuoroda)")
                continue
            if url in seen:
                lines.append(row + f" | tas pats failas kaip {seen[url]}")
                continue
            try:
                data, ctype = fetch(url)
                fn = f"{key}_{safe(title)}{ext_for(url, ctype, data)}"
                z.writestr(fn, data)
                seen[url] = fn
                ok += 1
                lines.append(row + f" | {fn} ({len(data) // 1024} KB)")
            except Exception as e:  # noqa: BLE001
                lines.append(row + f" | nepavyko: {str(e)[:120]}")
            time.sleep(1)
        z.writestr("00_saltiniai_sources.txt", "\n".join(lines) + "\n")
    return path, ok, len(rep.get("sources", {}))


def upload(path):
    if subprocess.run(["gh", "release", "view", TAG], capture_output=True).returncode != 0:
        subprocess.run(["gh", "release", "create", TAG, "--title", "Šaltinių paketai / Source packages",
         "--notes", "Ataskaitose naudotų viešų emitentų dokumentų kopijos (ZIP), sukurtos automatiškai. Copies of the public issuer documents used in the reports."],
                       check=True)
    subprocess.run(["gh", "release", "upload", TAG, str(path), "--clobber"], check=True)


def main():
    local = "--local" in sys.argv
    only = [a for a in sys.argv[1:] if not a.startswith("--")]
    state = C.load_json(STATE)
    repo = "VytisCode/issuers-review"
    for issuer in C.issuers():
        if only and issuer not in only:
            continue
        for rep in C.reports(issuer):
            name = f"{issuer}_{rep['period']}_sources.zip"
            if name in state and not local:
                continue
            path, ok, total = build_zip(issuer, rep)
            print(f"{name}: {ok}/{total} dokumentų, {path.stat().st_size // 1024} KB")
            if local or ok == 0:
                continue
            upload(path)
            state[name] = {"date": C.TODAY, "files": ok, "sources": total, "size_kb": path.stat().st_size // 1024,
                           "url": f"https://github.com/{repo}/releases/download/{TAG}/{name}"}
            C.save_json(STATE, state)


if __name__ == "__main__":
    main()
