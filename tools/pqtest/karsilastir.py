"""PQTest çıktısındaki Power Query tablosunu Python benzetimiyle (tools/pq_veri.py) karşılaştırır.

Kullanım: python tools/pqtest/karsilastir.py <pqtest çıktısı.json> <Ayarlar.xlsx> [--kapanis yyyy-mm] [--mudurluk ad] [--hizli]
"""

import json
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pq_veri  # noqa: E402


def bul(nesne, onek='[["Anahtar"'):
    """Çıktı JSON'unda sorgunun döndürdüğü tablo metnini ('[["Anahtar", ...') arar."""
    if isinstance(nesne, str) and nesne.startswith(onek):
        return nesne
    if isinstance(nesne, dict):
        nesne = list(nesne.values())
    if isinstance(nesne, list):
        for x in nesne:
            sonuc = bul(x, onek)
            if sonuc:
                return sonuc
    return None


def esit(a, b):
    if a in (None, "") and b in (None, ""):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 1e-6 * max(1.0, abs(b))
    return str(a) == str(b)


def main():
    ham = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8-sig"))
    metin = bul(ham)
    if metin is None:
        print(json.dumps(ham, ensure_ascii=False, indent=1)[:6000])
        raise SystemExit("Sorgu sonucu bulunamadı (hata yukarıda)")
    tablo = json.loads(metin)
    basliklar, pq = tablo[0], tablo[1:]
    if basliklar != pq_veri.COLUMNS:
        raise SystemExit(f"Sütunlar farklı: {basliklar}")
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("cikti")
    ap.add_argument("ayar")
    ap.add_argument("--kapanis", default="")
    ap.add_argument("--mudurluk", default="")
    ap.add_argument("--hizli", action="store_true")
    a = ap.parse_args()
    kapanis = tuple(map(int, a.kapanis.split("-"))) if a.kapanis else None
    beklenen = pq_veri.build(a.ayar, kapanis, a.mudurluk or None, a.hizli)
    print("PQ   :", len(pq), dict(Counter(r[1] for r in pq)))
    print("Beklenen:", len(beklenen), dict(Counter(r[1] for r in beklenen)))
    pq_k = {r[0]: r for r in pq}
    bek_k = {r[0]: r for r in beklenen}
    hata = 0
    for k in sorted(set(pq_k) ^ set(bek_k)):
        hata += 1
        if hata <= 15:
            print("Yalnız", "PQ" if k in pq_k else "beklenen", ":", k)
    farkli = 0
    for k in set(pq_k) & set(bek_k):
        a, b = pq_k[k], bek_k[k]
        # Kimlik satırındaki yenileme zamanı (DateTime.LocalNow) karşılaştırılmaz
        alanlar = [basliklar[i] for i in range(len(a)) if not esit(a[i], b[i]) and not (k == "Kimlik" and basliklar[i] == "Metin")]
        if alanlar:
            farkli += 1
            if farkli <= 15:
                print("Fark", k, [(f, a[basliklar.index(f)], b[basliklar.index(f)]) for f in alanlar[:4]])
    print(f"Anahtar farkı {hata}, değer farkı {farkli}")
    raise SystemExit(1 if hata or farkli or len(pq) != len(beklenen) else 0)


if __name__ == "__main__":
    main()
