"""Tek gider kalemi için hesap izi: Python ile Excel (Power Query şablonu) yan yana.

Aynı sentetik örnek, aynı kapanış, aynı kalem ve bölüm için iki sürümün verdiği tutarları satır satır koyar:
karşılaştırma bütçesi, gerçekleşen, seçili bütçe, fark, fark %, yön, kaynak dosya ve ay ay kaynak değerleri.
Şablon LibreOffice ile yeniden hesaplanır (Windows Excel'de açmak gerekmez). Fark varsa çıkış kodu 1'dir.

Kullanım:
    python tools/pqtest/kalem_izi.py                 # ekrana yazar
    python tools/pqtest/kalem_izi.py --yaz           # docs/KALEM_IZI.md dosyasını da günceller
    python tools/pqtest/kalem_izi.py --kod 7 --bolum Depolama   (bölüm adı örnekteki gibi)
"""

import argparse
import sys
import tempfile
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import sablon_dogrula as sd  # noqa: E402  (yolları ve bağımlılıkları kurar)

import kalem  # noqa: E402
import pq_sablon  # noqa: E402
import pq_veri  # noqa: E402
from model import Dataset  # noqa: E402

AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]


def tl(v):
    if v in (None, ""):
        return "—"
    return f"{Decimal(str(v)):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def yuzde(v):
    return "—" if v in (None, "") else f"%{float(v) * 100:.2f}".replace(".", ",")


def ayni(a, b):
    if a in (None, "") and b in (None, ""):
        return True
    if isinstance(a, str) or isinstance(b, str):
        return str(a).strip() == str(b).strip()
    return abs(float(a) - float(b)) < 0.005


def calistir(kod, bolum):
    data = Dataset(sd.AYAR, (2026, 10))
    r = kalem.detail(data, {"filters": {}}, kod, department=bolum)
    with tempfile.TemporaryDirectory() as d:
        wb = pq_sablon.build(pq_veri.build(sd.AYAR))
        wb["Kalem Ayrıntısı"]["B3"], wb["Kalem Ayrıntısı"]["B4"] = kod, bolum
        ws = sd.hesapla(Path(d), wb)["Kalem Ayrıntısı"]

    ozet = [ws.cell(8, c).value for c in range(1, 9)]
    satir = next(ws.cell(sd.pq_sablon.KA_ILK + j, 3).value for j in range(60)
                 if ws.cell(pq_sablon.KA_ILK + j, 1).value == bolum)
    p_bolum = r["departments"][0]
    adim = {t["step"].split(" · ")[0]: t for t in r["trace"]}
    karsi = [
        ("Gider kalemi", r["name"], ws["B5"].value, str),
        ("Karşılaştırma bütçesi", r["basis"].split(" (")[0], str(ws["B6"].value or "").split(" (")[0], str),
        ("Gerçekleşen (TL)", r["actual"], ozet[1], tl),
        ("Seçili bütçe (TL)", r["budget"], ozet[0], tl),
        ("Fark = gerçekleşen − bütçe (TL)", r["variance"], ozet[2], tl),
        ("Fark %", r["rate"], ozet[3], yuzde),
        ("Yön", r["direction"], ozet[4], str),
        ("Hacim etkisi (TL)", r["volume"], ozet[5], tl),
        ("Birim maliyet etkisi (TL)", r["price"], ozet[6], tl),
        ("Gerçekleşen kaynak dosyası", p_bolum["actualFile"], satir, str),
    ]
    # Ay ay kaynak değerleri: Python kaynak hücrelerinden, şablon aylık tablodan
    py_ay = defaultdict(dict)
    for s in r["sources"]:
        if s[4] in ("Gerçekleşen", "Bütçe"):
            py_ay[s[5]][s[4]] = (py_ay[s[5]].get(s[4]) or 0) + float(s[9] or 0)
    a0 = pq_sablon.KA_AY
    aylik = []
    for m in range(1, 13):
        ad = ws.cell(a0 + m, 1).value
        if not ad:
            continue
        donem = f"2026-{m:02d}"
        p = py_ay.get(donem, {})
        aylik.append((AYLAR[m - 1], p.get("Gerçekleşen"), ws.cell(a0 + m, 2).value, p.get("Bütçe"),
                      ws.cell(a0 + m, 3).value))
    return r, adim, karsi, aylik


def rapor(kod, bolum):
    r, adim, karsi, aylik = calistir(kod, bolum)
    hata = 0
    out = [f"### Kalem {kod} · {r['name']} · {bolum} · Ocak–Ekim 2026 (Ekim 2026 kapanışı)", "",
           "| Adım | Python | Excel (Power Query) | Aynı mı? |", "| --- | ---: | ---: | :---: |"]
    for ad, p, e, f in karsi:
        ok = ayni(p, e)
        hata += not ok
        out.append(f"| {ad} | {f(p)} | {f(e)} | {'✓' if ok else '✗'} |")
    out += ["", "Ay ay kaynak değerleri (Python: kaynak dosyadaki hücreler; Excel: şablonun aylık tablosu):", "",
            "| Ay | Gerçekleşen · Python | Gerçekleşen · Excel | Bütçe · Python | Bütçe · Excel | Aynı mı? |",
            "| --- | ---: | ---: | ---: | ---: | :---: |"]
    for ay, gp, ge, bp, be in aylik:
        ok = ayni(gp, ge) and ayni(bp, be)
        hata += not ok
        out.append(f"| {ay} | {tl(gp)} | {tl(ge)} | {tl(bp)} | {tl(be)} | {'✓' if ok else '✗'} |")
    toplam_g = sum(x[1] or 0 for x in aylik)
    toplam_b = sum(x[3] or 0 for x in aylik)
    ok = ayni(toplam_g, r["actual"]) and ayni(toplam_b, r["budget"])
    hata += not ok
    out += ["", f"Ayların toplamı gerçekleşen {tl(toplam_g)} · bütçe {tl(toplam_b)} TL; özet satırıyla "
            f"{'aynı ✓' if ok else 'FARKLI ✗'}.", "",
            "Python hesap izi (uygulamadaki Kalem ayrıntısı › Hesap izi tablosu):", "",
            "| Adım | Tutar | Kaynak dosya · dönem | Nasıl |", "| --- | ---: | --- | --- |"]
    for t in r["trace"]:
        v = yuzde(t["value"]) if t["step"] == "Fark %" else tl(t["value"]) + " TL"
        out.append(f"| {t['step']} | {v} | {t['source']} | {t['how']} |")
    return "\n".join(out), hata


BASLIK = """# Kalem hesap izi: Python ve Excel yan yana

Bu sayfa `tools/pqtest/kalem_izi.py --yaz` ile üretilir; elle düzenlemeyin. Aynı sentetik örnekte tek gider kalemi
için iki sürümün verdiği tutarları satır satır koyar. Excel tarafı Power Query şablonudur
(`powerquery/Butce_Rapor_Sablonu.xlsx`, Kalem Ayrıntısı sayfası); LibreOffice ile yeniden hesaplanır.
GitHub Actions her gönderimde aynı karşılaştırmayı çalıştırır, fark çıkarsa iş düşer.

**Nasıl okunur?** Bir yönetici "bu kalemde hangi bütçeye göre, ne kadar aştık?" diye sorduğunda cevap şu sırayla
izlenir: gerçekleşen (hangi dosya, hangi kapanış), seçili bütçe (güncel mi ilk mi, hangi dosya), fark ve yüzde.
Ay ay tablo, özet rakamın kaynak dosyadaki aylık değerlerin toplamı olduğunu gösterir.

**Kapsam dışı:** Makro (VBA) sürümü aynı beklenen değerleri `MakroTest.SentetikTest` ile kontrol eder; bu ortamda
Windows Excel olmadığı için makro çıktısı bu tabloya konmadı. Windows'ta makro testini çalıştırıp sonucu
`makro/Dogrulama_Durumu.json` dosyasına yazmak kullanıcıya kalan iştir.

**İlk bütçe bazı:** Karşılaştırma bütçesi "İlk bütçe" seçildiğinde Python hesap izi bütçe hücrelerini yılın ilk
kapanış dosyasından okur (Kasım 2026 düzeltmesi; önceden kapanış dosyasından okunuyor ve mutabakat "tutmuyor"
diyordu). Test: `python/tests/test_kalem.py › LineTraceBasisTests`.
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--kod", type=int)
    ap.add_argument("--bolum")
    ap.add_argument("--yaz", action="store_true")
    a = ap.parse_args()
    ornekler = [(a.kod, a.bolum)] if a.kod else [(3, "Giriş"), (1, "Sayım")]
    parcalar, hata = [], 0
    for kod, bolum in ornekler:
        metin, h = rapor(kod, bolum)
        print(metin, "\n")
        parcalar.append(metin)
        hata += h
    print("Toplam fark:", hata)
    if a.yaz:
        (sd.ROOT / "docs" / "KALEM_IZI.md").write_text(BASLIK + "\n" + "\n\n".join(parcalar) + "\n", encoding="utf-8")
    raise SystemExit(1 if hata else 0)


if __name__ == "__main__":
    main()
