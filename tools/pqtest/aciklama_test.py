"""Butce_Aciklama sorgusunun gerçek M motoru testi: üç yenileme adımı Python benzetimiyle karşılaştırılır.

  1. Önceki tablo yok: dönen dosyalardaki açıklamalar "Yeni açıklama" olarak gelir.
  2. Önceki tablo = 1. adımın sonucu + kullanıcı işlemleri (onay, ret, bu döneme aktar, metni değişen ve kaynağı
     kaldırılan satır): durumlar korunur, işlem geçmişine yazılır.
  3. Önceki tablo = 2. adımın sonucu, değişiklik yok: yeni işlem yazılmamalı.

Excel.CurrentWorkbook() Excel dışında olmadığından M tablosu olarak verilir (tools/pqtest/aciklama_calistir.ps1).

Kullanım:
  python tools/pqtest/aciklama_test.py hazirla ADIM ÇALIŞMA_KLASÖRÜ
  python tools/pqtest/aciklama_test.py karsilastir ADIM PQ_ÇIKTISI.json DÖNEN_KLASÖR ÇALIŞMA_KLASÖRÜ
"""

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import pq_aciklama  # noqa: E402
from karsilastir import bul, esit  # noqa: E402

AYAR = ROOT / "python" / "ornek" / "Ayarlar.xlsx"
DONEM, AYLAR, SIMDI = "2026-10", "2026:1,2,3,4,5,6,7,8,9,10", "2026-11-13 10:00"
KULLANICI = {1: "Test Kişi", 2: "F. Onaylayan", 3: "F. Onaylayan"}


def m(v):
    """Python değeri → M sabiti."""
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (int, float)):
        return repr(float(v)) if not float(v).is_integer() else str(int(v))
    s = str(v).replace("#(", "#(#)(").replace('"', '""').replace("\r", "").replace("\n", "#(lf)")
    return f'"{s}"'


def tablo(cols, rows):
    satirlar = ", ".join("{" + ", ".join(m(r.get(c)) for c in cols) + "}" for r in rows)
    return "#table({" + ", ".join(m(c) for c in cols) + "}, {" + satirlar + "})"


def kitap(onceki, kullanici):
    adlar = [("Kullanici", kullanici), ("Rapor_Donemi", DONEM), ("Rapor_Aylari", AYLAR)]
    satirlar = [f'{{"{ad}", #table({{"Column1"}}, {{{{{m(v)}}}}})}}' for ad, v in adlar]
    if onceki is not None:
        satirlar.insert(0, '{"Butce_Aciklama", ' + tablo(pq_aciklama.COLUMNS, onceki) + "}")
    return '#table({"Name", "Content"}, {' + ", ".join(satirlar) + "})"


def kullanici_islemleri(rows):
    """2. adım: Açıklama Onayı sayfasında kullanıcının yaptıkları."""
    out = []
    for r in rows:
        r = dict(r)
        if r["Anahtar"].startswith("Yerleştirme|1|2026-10"):
            r.update({"Durum": "Onaylandı", "Onay notu": "Uygun"})
        elif r["Anahtar"].startswith("Toplama|1|2026-10"):
            r.update({"Durum": "Reddedildi", "Onay notu": "Kaynak belirtin"})
        elif r["Anahtar"].startswith("Giriş|3|2026-09"):
            r.update({"Durum": "Bu döneme aktar"})
        elif r["Anahtar"].startswith("Tartım|1|2026-10"):
            # önceki onaylı metin; dosyadaki metin farklı → yeniden onay gerekir
            r.update({"Açıklama": "Eski metin", "Durum": "Onaylandı", "Kayıtlı durum": "Onaylandı"})
        out.append(r)
    out.append({c: None for c in pq_aciklama.COLUMNS} | {
        "Anahtar": f"Paketleme|1|{DONEM}|{AYLAR}", "Müdürlük": "Sevkiyat", "Bölüm": "Paketleme", "Gider kodu": 1,
        "Açıklama": "Kaynağı kaldırılan açıklama", "Kapanış": DONEM, "Aylar": AYLAR, "Kaynak dosya": "Eski.xlsx",
        "Kaynakta": "Var", "Durum": "Onaylandı", "Kayıtlı durum": "Onaylandı",
        "İşlem geçmişi": "2026-11-01 09:00 · X · Onaylandı"})
    return out


def oku_pq(path):
    ham = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    metin = bul(ham, '[["Anahtar"')
    if metin is None:
        print(json.dumps(ham, ensure_ascii=False, indent=1)[:6000])
        raise SystemExit("Sorgu sonucu bulunamadı (hata yukarıda)")
    t = json.loads(metin)
    if t[0] != pq_aciklama.COLUMNS:
        raise SystemExit(f"Sütunlar farklı: {t[0]}")
    return [dict(zip(t[0], r)) for r in t[1:]]


def hazirla(adim, is_klasoru, onceki=None):
    is_klasoru = Path(is_klasoru)
    is_klasoru.mkdir(parents=True, exist_ok=True)
    (is_klasoru / f"kitap{adim}.pq").write_text(kitap(onceki, KULLANICI[adim]), encoding="utf-8")
    (is_klasoru / f"onceki{adim}.json").write_text(json.dumps(onceki or [], ensure_ascii=False), encoding="utf-8")


def karsilastir(adim, pq_json, klasor, is_klasoru):
    is_klasoru = Path(is_klasoru)
    onceki = json.loads((is_klasoru / f"onceki{adim}.json").read_text(encoding="utf-8"))
    pq = oku_pq(pq_json)
    beklenen = pq_aciklama.build(AYAR, klasor, onceki, KULLANICI[adim], DONEM, AYLAR, SIMDI)
    hata = 0 if len(pq) == len(beklenen) else 1
    print(f"Adım {adim}: PQ {len(pq)} satır, beklenen {len(beklenen)}")
    for i, (a, b) in enumerate(zip(pq, beklenen)):
        fark = [c for c in pq_aciklama.COLUMNS if not esit(a[c], b[c])]
        if fark:
            hata += 1
            print("  Fark", i, a["Anahtar"], [(c, a[c], b[c]) for c in fark[:4]])
    for r in pq:
        print(f"  {r['Anahtar'][:48]:48} {str(r['Durum']):16} {str(r['Son işlem'])[:44]}")
    if adim == 3:  # değişiklik yoksa yeni işlem yazılmaz
        onceki_gecmis = {r["Anahtar"]: r["İşlem geçmişi"] for r in onceki}
        degisen = [r["Anahtar"] for r in pq if r["İşlem geçmişi"] != onceki_gecmis.get(r["Anahtar"])
                   and not r["Anahtar"].startswith("Dosya|")]
        if degisen:
            print("  Değişiklik yokken işlem yazıldı:", degisen)
            hata += 1
    if adim < 3:
        hazirla(adim + 1, is_klasoru, kullanici_islemleri(pq) if adim == 1 else pq)
    print(f"Adım {adim} fark: {hata}")
    raise SystemExit(1 if hata else 0)


if __name__ == "__main__":
    komut, adim = sys.argv[1], int(sys.argv[2])
    if komut == "hazirla":
        hazirla(adim, sys.argv[3])
    else:
        karsilastir(adim, *sys.argv[3:6])
