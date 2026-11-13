"""Power Query şablonunun formüllerini Python motoruyla karşılaştırır (LibreOffice ile yeniden hesaplayarak).

Her senaryoda: şablon örnek veriyle (tools/pq_veri.py) doldurulur, Kontrol Paneli seçimleri yazılır, LibreOffice
ile hesaplanır ve aynı seçimle `butce.run` çıktısıyla sayfa sayfa karşılaştırılır.

Ayrıca --aciklama KLASÖR verilirse (tools/pqtest/aciklama_ornek.py çıktısı) Açıklama Onayı sayfası açıklama
sorgusunun benzetim çıktısıyla doldurulur ve kontrol sütunu ile rapora yansıyan açıklamalar denetlenir.

Kullanım: python tools/pqtest/sablon_dogrula.py [--aciklama KLASÖR]   (soffice PATH'te olmalı)
"""

import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "python"))

from bootstrap import dependencies  # noqa: E402

dependencies()
from openpyxl import load_workbook  # noqa: E402

import butce  # noqa: E402
import pq_sablon  # noqa: E402
import pq_veri  # noqa: E402

AYAR = ROOT / "python" / "ornek" / "Ayarlar.xlsx"
SENARYOLAR = [
    # ad, kapsam, ay seçimi, elle aylar, Python aylar, Python çıktı dosyası
    ("Direktörlük · yıl başından", "Tümü", "Yıl başından kapanışa", [], "1,2,3,4,5,6,7,8,9,10", "Genel.xlsx"),
    ("Direktörlük · kapanış ayı", "Tümü", "Yalnız kapanış ayı", [], "10", "Genel.xlsx"),
    ("Depolama · Ocak, Mart, Mayıs", "Depolama", "Elle seçim", [1, 3, 5], "1,3,5", "Mudurluk_02_Depolama_*.xlsx"),
    # Müdürlük dosyası: sorgu yalnız Depolama verisini yükler (MudurlukDosyasi), kapsam kendiliğinden Depolama
    ("Müdürlük dosyası · Depolama", "Tümü", "Yıl başından kapanışa", [], "1,2,3,4,5,6,7,8,9,10",
     "Mudurluk_02_Depolama_*.xlsx", "Depolama"),
]


def satirlar(ws, ilk, sutunlar):
    sonuc = []
    for r in range(ilk, ws.max_row + 1):
        v = [ws.cell(r, c).value for c in sutunlar]
        if v[0] in (None, ""):
            break
        sonuc.append(v)
    return sonuc


def esit(a, b):
    if a in (None, "") and b in (None, ""):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= max(1e-6, 1e-9 * abs(b))
    return str(a) == str(b)


def karsilastir(ad, t, p):
    fark = [(i, a, b) for i, (a, b) in enumerate(zip(t, p)) if not all(esit(x, y) for x, y in zip(a, b))]
    hata = len(fark) + (len(t) != len(p))
    print(f"  {ad}: şablon {len(t)}, Python {len(p)} satır, fark {len(fark)}")
    for i, a, b in fark[:3]:
        print("    Ş", a, "\n    P", b)
    return hata


def senaryo(gecici, veri, ad, kapsam, mod, aylar, py_aylar, py_dosya, mudurluk=None):
    print("==", ad)
    wb = pq_sablon.build(pq_veri.build(AYAR, mudurluk=mudurluk) if mudurluk else veri)
    panel = wb["Kontrol Paneli"]
    panel["C6"], panel["C8"] = kapsam, mod
    for m in aylar:
        panel.cell(10, 2 + m, 1)
    tpl = hesapla(gecici, wb)

    cikti = gecici / "python"
    shutil.rmtree(cikti, ignore_errors=True)
    cikti.mkdir()
    butce.run(SimpleNamespace(ayar=AYAR, donem="2026-10", kapanis="2026-10", aylar=py_aylar,
                              cikti=cikti / "Genel.xlsx"))
    py = load_workbook(next(cikti.glob(py_dosya)), data_only=True)

    hata = 0
    for sayfa, son in (("Genel Özet", 9), ("Birim Maliyet", 15), ("Önemli Sapmalar", 11), ("Yıl Sonu Tahmini", 12),
                       ("Hacim ve Birim Maliyet", 12)):
        # Python "eşiği aşan sapma yok" gibi tek hücrelik bilgi satırlarını atla
        p = [r for r in satirlar(py[sayfa], 2, range(1, son)) if any(v not in (None, "") for v in r[1:])]
        t = satirlar(tpl[sayfa], 2, range(1, son))
        if mudurluk and sayfa == "Birim Maliyet":  # direktörlük ağırlığı müdürlük dosyasında bilinçli olarak boş
            if any(r[12] not in (None, "") for r in t):
                print("  Birim Maliyet: müdürlük dosyasında direktörlük ağırlığı boş olmalı")
                hata += 1
            t, p = [r[:12] + r[13:] for r in t], [r[:12] + r[13:] for r in p]
        hata += karsilastir(sayfa, t, p)
    tg, pg = tpl["Giderler"], py["Giderler"]
    bolumler = {b[0] for b in satirlar(tpl["Genel Özet"], 2, [1])}

    def bloklar(ws):  # 4. satırdaki bölüm adı → bloğun ilk sütunu
        return {ws.cell(4, c).value: c for c in range(4, ws.max_column + 1) if ws.cell(4, c).value in bolumler}

    tb, pb = bloklar(tg), bloklar(pg)
    if list(tb) != list(pb):
        print("  Giderler blokları farklı:", list(tb), list(pb))
        hata += 1
    for ad, c in tb.items():
        if ad in pb:
            hata += karsilastir(f"Giderler · {ad}", satirlar(tg, 6, [3] + list(range(c, c + 4))),
                                satirlar(pg, 6, [3] + list(range(pb[ad], pb[ad] + 4))))
    # Uyarılar (sıra bağımsız) ve revizyon sayısı
    tu = Counter(tuple(str(tpl["Uyarılar"].cell(r, c).value or "") for c in (2, 3, 5))
                 for r in range(2, tpl["Uyarılar"].max_row + 1) if tpl["Uyarılar"].cell(r, 5).value)
    kr = py["Kontrol ve Revizyon"]
    bas = next(r for r in range(1, kr.max_row + 1) if kr.cell(r, 1).value == "Uyarılar") + 2
    pu = Counter(tuple(str(kr.cell(r, c).value or "") for c in (1, 2, 4)) for r in range(bas, kr.max_row + 1)
                 if kr.cell(r, 4).value)
    print(f"  Uyarılar: şablon {sum(tu.values())}, Python {sum(pu.values())}")
    if tu != pu:
        print("    yalnız şablon", list((tu - pu).items())[:3], "\n    yalnız Python", list((pu - tu).items())[:3])
        hata += 1
    return hata


def guncel_mi():
    """Depodaki Butce_Rapor_Sablonu.xlsx, pq_sablon.py'nin ürettiğiyle hücre hücre aynı mı?"""
    with tempfile.TemporaryDirectory() as d:
        yeni = Path(d) / "yeni.xlsx"
        pq_sablon.build().save(yeni)
        a, b = load_workbook(ROOT / "powerquery" / "Butce_Rapor_Sablonu.xlsx"), load_workbook(yeni)
    if a.sheetnames != b.sheetnames:
        print("Şablon güncel değil: sayfalar farklı")
        return 1
    for ws in a:
        wb2 = b[ws.title]
        for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, wb2.max_row),
                                max_col=max(ws.max_column, wb2.max_column)):
            for c in row:
                if c.value != wb2[c.coordinate].value:
                    print(f"Şablon güncel değil: {ws.title}!{c.coordinate}. python tools/pq_sablon.py ile yeniden üretin.")
                    return 1
    print("Depodaki şablon güncel")
    return 0


def hesapla(gecici, wb):
    kaynak = gecici / "sablon.xlsx"
    wb.save(kaynak)
    hesap = gecici / "hesap"
    shutil.rmtree(hesap, ignore_errors=True)
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                    "--outdir", str(hesap), str(kaynak)], check=True, capture_output=True, timeout=300)
    return load_workbook(hesap / "sablon.xlsx", data_only=True)


def onay_senaryosu(gecici, veri, klasor):
    """Açıklama Onayı: kontrol sütunu ve onaylı açıklamanın rapora yansıması."""
    import pq_aciklama
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import aciklama_test as t

    print("== Açıklama Onayı")
    r1 = pq_aciklama.build(AYAR, klasor, [], "Test", t.DONEM, t.AYLAR, t.SIMDI)
    rows = pq_aciklama.build(AYAR, klasor, t.kullanici_islemleri(r1), "F. Onaylayan", t.DONEM, t.AYLAR, t.SIMDI)
    for r in rows:  # tutarı sonradan değişmiş gibi
        if r["Anahtar"].startswith("Giriş|1|2026-10"):
            r["Durum"] = "Onaylandı"
        if r["Anahtar"].startswith("Sayım|3|2026-10"):
            r["Gerçekleşen (dosyada)"] += 1000
    wb = pq_sablon.build(veri)
    ws = wb["Açıklama Onayı"]
    for c, name in enumerate(pq_aciklama.COLUMNS, 1):
        ws.cell(1, c, name)
    for i, r in enumerate(rows, 2):
        for c, name in enumerate(pq_aciklama.COLUMNS, 1):
            ws.cell(i, c, r[name])
    tpl = hesapla(gecici, wb)
    ws, kalem, onemli = tpl["Açıklama Onayı"], tpl["_Kalem"], tpl["Önemli Sapmalar"]
    kontrol = {ws.cell(r, 1).value: ws.cell(r, 27).value for r in range(2, len(rows) + 2)}
    metin = {kalem.cell(r, 8).value: kalem.cell(r, 27).value or "" for r in range(2, kalem.max_row + 1) if kalem.cell(r, 8).value}
    durum = {(onemli.cell(r, 2).value, onemli.cell(r, 3).value): onemli.cell(r, 11).value for r in range(2, 502)
             if onemli.cell(r, 2).value}
    k = f"|{t.DONEM}|{t.AYLAR}"
    beklenen_kontrol = {
        "Giriş|1" + k: "Güncel", "Yerleştirme|1" + k: "Güncel", "Toplama|1" + k: "Güncel", "Tartım|1" + k: "Güncel",
        "Sayım|3" + k: "Tutar değişti", "Giriş|99" + k: "Bu bölümde bu kod yok",
        "Giriş|3" + k: "Güncel (önceki dönemden aktarıldı)", "Giriş|3|2026-09|2026:1,2,3,4,5,6,7,8,9": "Önceki kapanış (2026-09)",
        "Paketleme|1" + k: "Tutar değişti",
        "Dosya|Notlar.xlsx": "Dosya okunamadı: Açıklamalar sayfası yok; dosya rapor şablonundan hazırlanmamış",
    }
    beklenen_metin = {
        "Giriş|1": "Kadro planı Mart'a kaydı; boş pozisyonlar Kasım'da doluyor.",
        "Yerleştirme|1": "Yeni WMS ile yerleştirme ekibi küçüldü.",
        "Tartım|1": "[Onay bekliyor] Vardiya sayısı 3'ten 2'ye indi; tasarruf kalıcı.",
        "Toplama|1": "", "Sayım|3": "", "Paketleme|1": "",
        "Giriş|3": "Eylül: denetim öncesi fazla mesai.",
    }
    hata = 0
    for a, b in beklenen_kontrol.items():
        if kontrol.get(a) != b:
            print(f"  Kontrol {a}: {kontrol.get(a)!r} ≠ {b!r}")
            hata += 1
    for a, b in beklenen_metin.items():
        if metin.get(a) != b:
            print(f"  Açıklama {a}: {metin.get(a)!r} ≠ {b!r}")
            hata += 1
    for (bolum, kod), bek in (("Giriş", 1), "Açıklandı"), (("Tartım", 1), "Onay bekliyor"), (("Toplama", 1), "Açıklama gerekli"):
        if (bolum, kod) in durum and durum[bolum, kod] != bek:
            print(f"  Önemli Sapmalar {bolum} {kod}: {durum[bolum, kod]!r} ≠ {bek!r}")
            hata += 1
    print(f"  Açıklama Onayı: {len(beklenen_kontrol)} kontrol, {len(beklenen_metin)} rapor açıklaması, fark {hata}")
    return hata


def revize_trend_senaryosu(gecici):
    """Revize bütçe seçimi, Bütçe Sürümleri ve Trend sayfaları (Python motorunda karşılığı yok: veriden beklenen)."""
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import revize_ornek

    print("== Revize bütçe ve trend")
    hata = 0
    revize_ornek.main(gecici / "revize")
    veri = pq_veri.build(gecici / "revize" / "Ayarlar.xlsx")
    c = pq_veri.COLUMNS
    gi, bi, ri = c.index("G1"), c.index("B1"), c.index("R1")
    toplam = {r[4]: r for r in veri if r[1] == "Toplam" and r[8] == 2026}

    def katı(r, i, aylar=range(10)):
        v = [r[i + m] for m in aylar]
        return None if any(x is None for x in v) else sum(v)

    for secim in ("Bütçe", "Revize bütçe"):
        wb = pq_sablon.build(veri)
        wb["Kontrol Paneli"]["C7"] = secim
        tpl = hesapla(gecici, wb)
        genel = {r[0]: r for r in satirlar(tpl["Genel Özet"], 2, range(1, 6)) if r[1] == "Bölüm"}
        sur = {r[0]: r for r in satirlar(tpl["Bütçe Sürümleri"], 2, range(1, 10)) if r[1] == "Bölüm"}
        for ad, r in toplam.items():
            bek = katı(r, ri if secim == "Revize bütçe" else bi)
            for adi, deger, beklenen in (("Genel Özet bütçe", genel[ad][4], bek),
                                         ("Sürümler bütçe", sur[ad][3], katı(r, bi)),
                                         ("Sürümler revize", sur[ad][4], katı(r, ri))):
                if not esit(deger, beklenen):
                    print(f"  {secim} · {ad} · {adi}: {deger} ≠ {beklenen}")
                    hata += 1
        if secim == "Revize bütçe" and not (genel["Giriş"][4] in (None, "") and sur["Giriş"][4] in (None, "")):
            print("  Giriş revize bütçesi Şubat'ı eksik: boş olmalı")
            hata += 1
        print(f"  {secim}: {len(toplam)} bölüm, Genel Özet ve Bütçe Sürümleri karşılaştırıldı")

    # Trend: örnek veri, karşılaştırma bütçesi = bütçe
    veri = pq_veri.build(AYAR)
    tpl = hesapla(gecici, pq_sablon.build(veri))
    esik = 1_000_000 / 12
    beklenen = []
    for r in veri:
        if r[1] != "Gider" or r[8] != 2026:
            continue
        fark = [r[gi + m] - r[bi + m] if r[gi + m] is not None and r[bi + m] is not None else None for m in range(10)]
        seri = 0
        for f in reversed(fark):
            if f is None or f < esik:
                break
            seri += 1
        if seri >= 3:
            beklenen.append((r[3], r[4], r[6], r[5], seri, sum(1 for f in fark if f is not None and f >= esik),
                             sum(fark[10 - seri:]), fark[9]))
    beklenen.sort(key=lambda x: (-x[4], -x[6]))
    tr = tpl["Trend"]
    liste = [tuple(tr.cell(r, col).value for col in range(17, 25)) for r in range(20, 520) if tr.cell(r, 26).value not in (None, "")]
    if len(liste) != len(beklenen) or any(not all(esit(a, b) for a, b in zip(x, y)) for x, y in zip(liste, beklenen)):
        print(f"  Trend listesi farklı: şablon {len(liste)}, beklenen {len(beklenen)}")
        for x, y in list(zip(liste, beklenen))[:3]:
            print("    Ş", x, "\n    B", y)
        hata += 1
    kapsam = [tr.cell(8, 3 + m).value for m in range(12)]
    bek = [sum(t[gi + m] - t[bi + m] for t in toplam.values()) for m in range(10)]
    if not all(esit(a, b) for a, b in zip(kapsam[:10], bek)) or any(v not in (None, "") for v in kapsam[10:]):
        print("  Trend kapsam toplamı farklı:", kapsam[:3], bek[:3])
        hata += 1
    print(f"  Trend: {len(beklenen)} kalem üst üste ≥3 ay aleyhte, fark {hata}")
    return hata


def main():
    veri = pq_veri.build(AYAR)
    hata = guncel_mi()
    klasor = sys.argv[sys.argv.index("--aciklama") + 1] if "--aciklama" in sys.argv else None
    with tempfile.TemporaryDirectory() as d:
        hata += revize_trend_senaryosu(Path(d))
        if klasor:
            hata += onay_senaryosu(Path(d), veri, klasor)
        for s in SENARYOLAR:
            hata += senaryo(Path(d), veri, *s)
    print("Toplam fark:", hata)
    raise SystemExit(1 if hata else 0)


if __name__ == "__main__":
    main()
