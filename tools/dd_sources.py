"""DD dokumentacijos ZIP (GitHub Actions, update.yml): data/dd/jobs/<DD id>.json -> atsisiunčia viešus šaltinius,
supakuoja į ZIP, užšifruoja viešuoju DD raktu (tools/dd_vault.py) ir prideda į DD archyvą kaip {type: zip, parent}.
Užduotis lieka faile (done = ZIP id), kad būtų matyti, kas jau padaryta. Be AI, nemokamai.
"""
import io
import json
import sys
import time
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import archive_sources as A  # noqa: E402
import dd_vault as V  # noqa: E402

ZIPS = V.DIR / "zips.json"


def run():
    jobs = sorted((V.DIR / "jobs").glob("*.json")) if (V.DIR / "jobs").exists() else []
    zips = json.loads(ZIPS.read_text()) if ZIPS.exists() else {}
    for jf in jobs:
        job = json.loads(jf.read_text())
        if job.get("done"):
            continue
        buf, lines, ok = io.BytesIO(), [job["name"], f"Atsisiųsta {time.strftime('%Y-%m-%d')}", ""], 0
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for s in job["sources"]:
                line = f"[{s['key']}] {s['title']} - {s['url']}"
                if any(d in s["url"] for d in A.SKIP) or not s["url"].startswith("http"):
                    lines.append(line + "  (neatsisiunčiama)")
                    continue
                try:
                    data, ctype = A.fetch(s["url"])
                    data, ext = A.as_pdf(s["url"], data, A.ext_for(s["url"], ctype, data))
                    fn = f"{s['key']}_{A.safe(s['title'], 60)}{ext}"
                    z.writestr(fn, data)
                    lines.append(line + f"  -> {fn}")
                    ok += 1
                except Exception as e:  # noqa: BLE001
                    lines.append(line + f"  (nepavyko: {e})")
                time.sleep(1)
            z.writestr("_saltiniai.txt", "\n".join(lines) + "\n")
        zid = V.add_bytes(buf.getvalue(), {"type": "zip", "parent": job["parent"], "file": job["name"], "docs": ok})
        zips[zid] = job["parent"]
        # tik naujausias dokumentų ZIP prie įrašo (savininkas 2026-10-29: du „Dokumentai (ZIP)“ mygtukai)
        ev = V.EVID()
        stale = [z for z, p in zips.items() if p == job["parent"] and z != zid and z not in ev]
        if stale:
            V.INDEX.write_text(json.dumps([x for x in json.loads(V.INDEX.read_text()) if x["id"] not in stale], indent=1) + "\n")
            for z in stale:
                zips.pop(z, None); (V.DIR / f"{z}.bin").unlink(missing_ok=True)
        job["done"] = zid
        jf.write_text(json.dumps(job, ensure_ascii=False, indent=1))
        print(f"{job['name']}: {ok}/{len(job['sources'])} dokumentų -> {zid}")
    ZIPS.write_text(json.dumps(zips, indent=1) + "\n")


if __name__ == "__main__":
    run()
