"""Užšifruotas DD archyvas dashboard'e (savininko sprendimas 2026-10-23).

DD dokumentuose yra ne viešų duomenų (term sheet, brokerio laiškai), todėl į viešą repo dedami tik užšifruoti failai:
- data/dd/key.json   - viešasis RSA raktas + privatusis raktas, užšifruotas savininko slaptažodžiu (PBKDF2 + AES-GCM);
- data/dd/index.json - įrašai: atsitiktinis id, RSA užšifruotas AES raktas, užšifruoti metaduomenys;
- data/dd/<id>.bin   - užšifruotas failas (12 B IV + AES-GCM).
Naujam DD slaptažodžio nereikia (šifruojama viešuoju raktu); iššifruoja tik naršyklė su slaptažodžiu.
Slaptažodis niekur nesaugomas. Pametus - `init --force` (naujas raktas) ir visus DD užšifruoti iš naujo.

  python tools/dd_vault.py init                      # vieną kartą: sukuria raktą, atspausdina slaptažodį
  python tools/dd_vault.py add FAILAS.docx --issuer ignitis --name "Ignitis grupė" --isin XS... --lang lt --title "..."
  python tools/dd_vault.py add ... --sources SALTINIAI.json  # + vieši šaltiniai ZIP'ui ([{key,title,url}]), ZIP sukuria GitHub Actions
  python tools/dd_vault.py replace ID                # pataisytas DD (tas pats failas manifest'e) – ZIP lieka
  python tools/dd_vault.py remove ID                 # pašalina įrašą (ir jo ZIP)
  python tools/dd_vault.py rekey                     # pamiršus slaptažodį: naujas raktas, visi DD iš projekto failų iš naujo
Dokumentacijos ZIP: data/dd/jobs/<DD id>.json (tik vieši šaltinių URL) -> tools/dd_sources.py (update.yml) atsisiunčia,
supakuoja, užšifruoja viešuoju raktu ir prideda įrašą {type: zip, parent: <DD id>}.
Atviras sąrašas (failų keliai + metaduomenys) laikomas tik projekto failuose: MANIFEST (ne repo).
"""
import argparse
import base64
import datetime as dt
import json
import secrets
import sys
from pathlib import Path

from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

DIR = C.DATA / "dd"
KEY, INDEX = DIR / "key.json", DIR / "index.json"
ITER = 600_000
MANIFEST = Path("/mnt/project-files/dd/manifest.json")
OAEP = padding.OAEP(mgf=padding.MGF1(hashes.SHA256()), algorithm=hashes.SHA256(), label=None)
b64 = lambda b: base64.b64encode(b).decode()  # noqa: E731


def init(force=False, pw=None):
    if KEY.exists() and not force:
        sys.exit("data/dd/key.json jau yra (naujam raktui: --force; senų DD nebus galima atidaryti).")
    pw = pw or "-".join(secrets.token_urlsafe(6) for _ in range(4))
    priv = rsa.generate_private_key(public_exponent=65537, key_size=3072)
    salt, iv = secrets.token_bytes(16), secrets.token_bytes(12)
    kek = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(pw.encode())
    pkcs8 = priv.private_bytes(serialization.Encoding.DER, serialization.PrivateFormat.PKCS8, serialization.NoEncryption())
    spki = priv.public_key().public_bytes(serialization.Encoding.DER, serialization.PublicFormat.SubjectPublicKeyInfo)
    DIR.mkdir(parents=True, exist_ok=True)
    KEY.write_text(json.dumps({"pub": b64(spki), "salt": b64(salt), "iter": ITER, "iv": b64(iv),
                               "priv": b64(AESGCM(kek).encrypt(iv, pkcs8, None))}, indent=1))
    if force or not INDEX.exists():
        INDEX.write_text("[]\n")
    print("Slaptažodis (perduokite savininkui, niekur nesaugokite):", pw)


def add_bytes(data, meta, eid=None):
    """Užšifruoja baitus viešuoju raktu ir prideda (arba, nurodžius eid, pakeičia) įrašą index.json; grąžina id."""
    pub = serialization.load_der_public_key(base64.b64decode(json.loads(KEY.read_text())["pub"]))
    k, fiv, miv = AESGCM.generate_key(256), secrets.token_bytes(12), secrets.token_bytes(12)
    eid = eid or secrets.token_hex(8)
    meta = {**meta, "bytes": len(data)}
    (DIR / f"{eid}.bin").write_bytes(fiv + AESGCM(k).encrypt(fiv, data, None))
    idx = [x for x in json.loads(INDEX.read_text()) if x["id"] != eid]
    idx.append({"id": eid, "k": b64(pub.encrypt(k, OAEP)), "iv": b64(miv),
                "meta": b64(AESGCM(k).encrypt(miv, json.dumps(meta, ensure_ascii=False).encode(), None))})
    INDEX.write_text(json.dumps(idx, indent=1) + "\n")
    return eid


def add(path, meta, sources=None, record=True):
    eid = add_bytes(Path(path).read_bytes(), {**meta, "type": "dd", "file": Path(path).name})
    if sources:
        (DIR / "jobs").mkdir(exist_ok=True)
        (DIR / "jobs" / f"{eid}.json").write_text(json.dumps({"parent": eid, "name": f"{Path(path).stem}_documents.zip",
                                                              "sources": sources}, ensure_ascii=False, indent=1))
    if record and MANIFEST.parent.exists():
        man = json.loads(MANIFEST.read_text()) if MANIFEST.exists() else []
        man = [m for m in man if m["path"] != str(Path(path).resolve())] + [{"path": str(Path(path).resolve()), "id": eid, "meta": meta, "sources": sources or []}]
        MANIFEST.write_text(json.dumps(man, ensure_ascii=False, indent=1))
    print(f"Pridėta: {eid} ({Path(path).name})" + (" + ZIP užduotis" if sources else ""))
    return eid


def replace(eid):
    """Pataisytas DD (tas pats failas manifest'e): naujas id (kad naršyklė su sena data.js kopija neimtų naujo failo
    su senu raktu), senas įrašas pašalinamas, ZIP perkeliamas naujam id per data/dd/zips.json."""
    man = json.loads(MANIFEST.read_text())
    m = next(m for m in man if m.get("id") == eid)
    new = add_bytes(Path(m["path"]).read_bytes(), {**m["meta"], "type": "dd", "file": Path(m["path"]).name})
    INDEX.write_text(json.dumps([x for x in json.loads(INDEX.read_text()) if x["id"] != eid], indent=1) + "\n")
    (DIR / f"{eid}.bin").unlink(missing_ok=True)
    zf = DIR / "zips.json"
    zips = json.loads(zf.read_text()) if zf.exists() else {}
    zf.write_text(json.dumps({z: (new if p == eid else p) for z, p in zips.items()}, indent=1) + "\n")
    job = DIR / "jobs" / f"{eid}.json"
    if job.exists():
        j = json.loads(job.read_text()); j["parent"] = new
        (DIR / "jobs" / f"{new}.json").write_text(json.dumps(j, ensure_ascii=False, indent=1)); job.unlink()
    m["id"] = new
    MANIFEST.write_text(json.dumps(man, ensure_ascii=False, indent=1))
    print("Atnaujinta:", eid, "->", new, Path(m["path"]).name)


def remove(eid):
    """Pašalina DD įrašą ir su juo susijusį ZIP (ZIP atpažįstamas pagal data/dd/zips.json)."""
    zips = json.loads((DIR / "zips.json").read_text()) if (DIR / "zips.json").exists() else {}
    ids = {eid, *[z for z, parent in zips.items() if parent == eid]}
    idx = [x for x in json.loads(INDEX.read_text()) if x["id"] not in ids]
    INDEX.write_text(json.dumps(idx, indent=1) + "\n")
    for i in ids:
        (DIR / f"{i}.bin").unlink(missing_ok=True)
        (DIR / "jobs" / f"{i}.json").unlink(missing_ok=True)
    (DIR / "zips.json").write_text(json.dumps({z: p for z, p in zips.items() if p != eid}, indent=1) + "\n")
    if MANIFEST.exists():
        MANIFEST.write_text(json.dumps([m for m in json.loads(MANIFEST.read_text()) if m.get("id") != eid], ensure_ascii=False, indent=1))
    print("Pašalinta:", ", ".join(sorted(ids)))


def rekey(pw=None):
    man = json.loads(MANIFEST.read_text())
    for f in DIR.glob("*.bin"):
        f.unlink()
    (DIR / "zips.json").write_text("{}\n")
    init(True, pw)
    new = []
    for m in man:
        m["id"] = add(m["path"], m["meta"], m.get("sources"), record=False)
        new.append(m)
    MANIFEST.write_text(json.dumps(new, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    i = sub.add_parser("init"); i.add_argument("--force", action="store_true"); i.add_argument("--password")
    r = sub.add_parser("rekey"); r.add_argument("--password")
    rm = sub.add_parser("remove"); rm.add_argument("id")
    rp = sub.add_parser("replace"); rp.add_argument("id")
    a = sub.add_parser("add"); a.add_argument("file"); a.add_argument("--sources")
    for f in ("issuer", "name", "isin", "lang", "title", "decision", "sfdr"):
        a.add_argument("--" + f, default="")
    a.add_argument("--date", default=dt.date.today().isoformat())
    x = ap.parse_args()
    if x.cmd == "init":
        init(x.force, x.password)
    elif x.cmd == "rekey":
        rekey(x.password)
    elif x.cmd == "replace":
        replace(x.id)
    elif x.cmd == "remove":
        remove(x.id)
    else:
        add(x.file, {k: getattr(x, k) for k in ("issuer", "name", "isin", "lang", "title", "decision", "sfdr", "date")},
            json.loads(Path(x.sources).read_text()) if x.sources else None)
