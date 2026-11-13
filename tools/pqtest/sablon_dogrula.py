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


def butce_trend_senaryosu(gecici):
    """Bütçe tabloları, bütçe değişimi, ilk/güncel bütçe ve trend (Python motorunda karşılığı yok: veriden beklenen)."""
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import butce_degisim_ornek

    print("== Bütçe tabloları, bütçe değişimi ve trend")
    hata = 0
    butce_degisim_ornek.main(gecici / "bd")
    veri = pq_veri.build(gecici / "bd" / "Ayarlar.xlsx")
    c = pq_veri.COLUMNS
    gi, bi, ri = c.index("G1"), c.index("B1"), c.index("R1")
    toplam = {r[4]: r for r in veri if r[1] == "Toplam" and r[8] == 2026}
    onceki = {r[4]: r for r in veri if r[1] == "OncekiToplam"}
    duzen = [r for r in veri if r[1] == "Duzen"]

    def uyeler(d):  # düzen satırının bölümleri
        tur = d[7]
        return [b for b in toplam if (tur == "Bölüm" and b == d[4])
                or (tur == "Müdürlük toplamı" and toplam[b][3] == d[3])
                or (tur == "Direktörlük toplamı" and toplam[b][2] == d[2]) or tur == "Genel toplam"]

    def topla(satirlar, i, aylar=range(12)):
        out = []
        for m in aylar:
            v = [x[i + m] for x in satirlar]
            out.append(None if not v or any(y is None for y in v) else sum(v))
        return out

    def kontrol(ad, ş, b):
        nonlocal hata
        if len(ş) != len(b) or not all(esit(x, y) for x, y in zip(ş, b)):
            print(f"  {ad}: {ş[:4]} ≠ {b[:4]}")
            hata += 1

    for secim in ("Güncel bütçe", "İlk bütçe (yıl başı)"):
        wb = pq_sablon.build(veri)
        wb["Kontrol Paneli"]["C7"] = secim
        tpl = hesapla(gecici, wb)
        genel = {r[0]: r for r in satirlar(tpl["Genel Özet"], 2, range(1, 6)) if r[1] == "Bölüm"}
        sur = {r[0]: r for r in satirlar(tpl["Bütçe Sürümleri"], 2, range(1, 10)) if r[1] == "Bölüm"}
        for ad, r in toplam.items():
            g, i = sum(topla([r], bi, range(10))), sum(topla([r], ri, range(10)))
            kontrol(f"{secim} · {ad} · Genel Özet bütçe", [genel[ad][4]], [i if secim.startswith("İlk") else g])
            kontrol(f"{ad} · Bütçe Sürümleri", [sur[ad][3], sur[ad][4]], [g, i])
        if secim.startswith("İlk"):
            continue
        # Bütçe - Bölüm / Müdürlük / Direktörlük
        for sayfa, tur, ilk_ay in (("Bütçe - Bölüm", "Bölüm", 3), ("Bütçe - Müdürlük", "Müdürlük toplamı", 3),
                                   ("Bütçe - Direktörlük", "Direktörlük toplamı", 2)):
            ws = tpl[sayfa]
            bek = [d for d in duzen if d[7] == tur]
            got = satirlar(ws, 5, range(1, ilk_ay + 15))
            if len(got) != len(bek):
                print(f"  {sayfa}: {len(got)} satır, beklenen {len(bek)}")
                hata += 1
            for d, row in zip(bek, got):
                aylar = topla([toplam[b] for b in uyeler(d)], bi)
                ilk = topla([toplam[b] for b in uyeler(d)], ri)
                yil = sum(aylar) if all(v is not None for v in aylar) else None
                ilk_yil = sum(ilk) if all(v is not None for v in ilk) else None
                kontrol(f"{sayfa} · {d[5]}", [row[0]] + row[ilk_ay - 1:ilk_ay + 14],
                        [d[5]] + aylar + [yil, ilk_yil, None if yil is None or ilk_yil is None else yil - ilk_yil])
        # Bütçe Değişimi: seviye satırları ve kalemler
        ws = tpl["Bütçe Değişimi"]
        bek = []
        for d in duzen:
            u = uyeler(d)
            if not u or any(b not in onceki for b in u):
                continue
            g, p = topla([toplam[b] for b in u], bi), topla([onceki[b] for b in u], bi)
            fark = [None if x is None or y is None else x - y for x, y in zip(g, p)]
            if any(f is not None and abs(f) > 0.01 for f in fark):
                bek.append([d[5], d[7]] + fark)
        got = [r for r in satirlar(ws, 5, range(1, 15))]
        if len(got) != len(bek):
            print(f"  Bütçe Değişimi: {len(got)} satır, beklenen {len(bek)}")
            hata += 1
        for x, y in zip(got, bek):
            kontrol(f"Bütçe Değişimi · {y[0]}", [x[0], x[1]] + [v if v not in (None, "") and abs(v) > 1e-9 else 0 for v in x[2:]],
                    [y[0], y[1]] + [v if v is not None else 0 for v in y[2:]])
        kalem_bek = []
        gider = {(r[4], r[6]): r for r in veri if r[1] == "Gider" and r[8] == 2026}
        for r in veri:
            if r[1] != "OncekiGider" or (r[4], r[6]) not in gider:
                continue
            g, p = gider[r[4], r[6]], r
            ay = sum(1 for m in range(12) if g[bi + m] is not None and p[bi + m] is not None and abs(g[bi + m] - p[bi + m]) > 0.01)
            if ay:
                kalem_bek.append((r[4], r[6], ay, sum(g[bi:bi + 12]) - sum(p[bi:bi + 12])))
        sira = {name: i for i, name in enumerate(toplam)}
        kalem_bek.sort(key=lambda x: (sira[x[0]], x[1]))
        kalem_got = [(ws.cell(r, 22).value, ws.cell(r, 23).value, ws.cell(r, 25).value, ws.cell(r, 28).value)
                     for r in range(5, 505) if ws.cell(r, 49).value not in (None, "")]
        kontrol("Bütçesi değişen kalemler", [v for t in kalem_got for v in t], [v for t in kalem_bek for v in t])
        print(f"  Bütçe tabloları 3 sayfa, değişim {len(bek)} satır / {len(kalem_bek)} kalem")
    print(f"  İlk/güncel bütçe seçimi ve Bütçe Sürümleri: {len(toplam)} bölüm")

    # Trend: örnek veri, karşılaştırma bütçesi = bütçe
    veri = pq_veri.build(AYAR)
    toplam = {r[4]: r for r in veri if r[1] == "Toplam" and r[8] == 2026}
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
        hata += butce_trend_senaryosu(Path(d))
        if klasor:
            hata += onay_senaryosu(Path(d), veri, klasor)
        for s in SENARYOLAR:
            hata += senaryo(Path(d), veri, *s)
    print("Toplam fark:", hata)
    raise SystemExit(1 if hata else 0)


if __name__ == "__main__":
    main()
