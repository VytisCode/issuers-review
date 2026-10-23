"""Naujienų antraščių vertimas (nemokamai, be Claude) -> data/news_tr.csv (tik papildoma).

Savininkas neskaito latviškai ir estiškai, todėl lv / et (ir en) antraštės verčiamos į lt, o lv / et / lt - į en
(svetainės kalbos). Naudojamas viešas Google Translate adresas be rakto; nepavykus antraštė lieka originalo kalba.
Neigiamų žodžių patikra (flag=review) daroma originalo tekste, vertimas tik rodymui.

Naudojimas: python tools/translate_news.py   (update.yml, po fetch_news.py)
"""
import json
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common as C  # noqa: E402

OUT = C.DATA / "news_tr.csv"
FIELDS = ["id", "to", "text"]
TARGETS = {"lt": ["lv", "et", "en"], "en": ["lv", "et", "lt"]}  # į kurią kalbą -> iš kurių
MAX = 400  # per vieną paleidimą


def translate(text, src, to):
    url = "https://translate.googleapis.com/translate_a/single?" + urllib.parse.urlencode({"client": "gtx", "sl": src, "tl": to, "dt": "t", "q": text})
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (issuers-review)"})
    with urllib.request.urlopen(req, timeout=20) as r:
        data = json.loads(r.read().decode("utf-8"))
    return "".join(part[0] for part in data[0] if part and part[0]).strip()


def main():
    rows = C.read_csv(OUT) if OUT.exists() else []
    have = {(r["id"], r["to"]) for r in rows}
    news = sorted(C.read_csv(C.DATA / "news.csv"), key=lambda r: r["date"], reverse=True)
    todo = [(n, to) for n in news for to, srcs in TARGETS.items() if n["lang"] in srcs and (n["id"], to) not in have]
    done = fails = 0
    for n, to in todo[:MAX]:
        try:
            txt = translate(n["title"], n["lang"], to)
        except Exception as e:  # tinklas, limitai - bandoma kitą kartą
            fails += 1
            print(f"  {n['id']} -> {to}: {e}")
            if fails >= 5:
                break
            continue
        if txt:
            rows.append({"id": n["id"], "to": to, "text": txt})
            done += 1
        time.sleep(0.3)
    C.write_csv(OUT, rows, FIELDS)
    print(f"news_tr.csv: +{done} vertimų, liko {max(0, len(todo) - done)}, klaidų {fails}")


if __name__ == "__main__":
    main()
