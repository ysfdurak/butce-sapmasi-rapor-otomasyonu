"""PQTest çıktısındaki Power Query tablosunu Python benzetimiyle (tools/pq_veri.py) karşılaştırır.

Kullanım: python tools/pqtest/karsilastir.py <pqtest çıktısı.json> <Ayarlar.xlsx> [kapanış yyyy-mm]
"""

import json
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pq_veri  # noqa: E402


def bul(nesne):
    """Çıktı JSON'unda sorgunun döndürdüğü tablo metnini ('[["Anahtar", ...') arar."""
    if isinstance(nesne, str) and nesne.startswith('[["Anahtar"'):
        return nesne
    if isinstance(nesne, dict):
        nesne = list(nesne.values())
    if isinstance(nesne, list):
        for x in nesne:
            sonuc = bul(x)
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
    kapanis = tuple(map(int, sys.argv[3].split("-"))) if len(sys.argv) > 3 and sys.argv[3] else None
    beklenen = pq_veri.build(sys.argv[2], kapanis)
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
        alanlar = [basliklar[i] for i in range(len(a)) if not esit(a[i], b[i])]
        if alanlar:
            farkli += 1
            if farkli <= 15:
                print("Fark", k, [(f, a[basliklar.index(f)], b[basliklar.index(f)]) for f in alanlar[:4]])
    print(f"Anahtar farkı {hata}, değer farkı {farkli}")
    raise SystemExit(1 if hata or farkli or len(pq) != len(beklenen) else 0)


if __name__ == "__main__":
    main()
