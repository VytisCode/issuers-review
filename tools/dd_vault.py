"""Užšifruotas DD archyvas dashboard'e (savininko sprendimas 2026-10-23).

DD dokumentuose yra ne viešų duomenų (term sheet, brokerio laiškai), todėl į viešą repo dedami tik užšifruoti failai:
- data/dd/key.json   - viešasis RSA raktas + privatusis raktas, užšifruotas savininko slaptažodžiu (PBKDF2 + AES-GCM);
- data/dd/index.json - įrašai: atsitiktinis id, RSA užšifruotas AES raktas, užšifruoti metaduomenys;
- data/dd/<id>.bin   - užšifruotas failas (12 B IV + AES-GCM).
Naujam DD slaptažodžio nereikia (šifruojama viešuoju raktu); iššifruoja tik naršyklė su slaptažodžiu.
Slaptažodis niekur nesaugomas. Pametus - `init --force` (naujas raktas) ir visus DD užšifruoti iš naujo.

  python tools/dd_vault.py init                      # vieną kartą: sukuria raktą, atspausdina slaptažodį
  python tools/dd_vault.py add FAILAS.docx --issuer ignitis --name "Ignitis grupė" --isin XS... --lang lt --title "..."
  python tools/dd_vault.py rekey                     # pamiršus slaptažodį: naujas raktas, visi DD iš projekto failų iš naujo
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


def add(path, meta, record=True):
    pub = serialization.load_der_public_key(base64.b64decode(json.loads(KEY.read_text())["pub"]))
    data = Path(path).read_bytes()
    k, fiv, miv = AESGCM.generate_key(256), secrets.token_bytes(12), secrets.token_bytes(12)
    eid = secrets.token_hex(8)
    meta = {**meta, "file": Path(path).name, "bytes": len(data)}
    (DIR / f"{eid}.bin").write_bytes(fiv + AESGCM(k).encrypt(fiv, data, None))
    idx = json.loads(INDEX.read_text())
    idx.append({"id": eid, "k": b64(pub.encrypt(k, OAEP)), "iv": b64(miv),
                "meta": b64(AESGCM(k).encrypt(miv, json.dumps(meta, ensure_ascii=False).encode(), None))})
    INDEX.write_text(json.dumps(idx, indent=1) + "\n")
    if record and MANIFEST.parent.exists():
        man = json.loads(MANIFEST.read_text()) if MANIFEST.exists() else []
        man = [m for m in man if m["path"] != str(Path(path).resolve())] + [{"path": str(Path(path).resolve()), "meta": meta}]
        MANIFEST.write_text(json.dumps(man, ensure_ascii=False, indent=1))
    print(f"Pridėta: {eid} ({meta['file']})")


def rekey(pw=None):
    man = json.loads(MANIFEST.read_text())
    for f in DIR.glob("*.bin"):
        f.unlink()
    init(True, pw)
    for m in man:
        add(m["path"], {k: v for k, v in m["meta"].items() if k not in ("file", "bytes")}, record=False)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    i = sub.add_parser("init"); i.add_argument("--force", action="store_true"); i.add_argument("--password")
    r = sub.add_parser("rekey"); r.add_argument("--password")
    a = sub.add_parser("add"); a.add_argument("file")
    for f in ("issuer", "name", "isin", "lang", "title", "decision", "sfdr"):
        a.add_argument("--" + f, default="")
    a.add_argument("--date", default=dt.date.today().isoformat())
    x = ap.parse_args()
    if x.cmd == "init":
        init(x.force, x.password)
    elif x.cmd == "rekey":
        rekey(x.password)
    else:
        add(x.file, {k: getattr(x, k) for k in ("issuer", "name", "isin", "lang", "title", "decision", "sfdr", "date")})
