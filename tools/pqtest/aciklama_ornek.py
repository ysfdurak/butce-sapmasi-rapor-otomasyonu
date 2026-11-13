"""Örnek "müdürlükten dönen" açıklama dosyaları üretir (Butce_Aciklama sorgusunun testi ve tanıtımı için).

Her dosya, rapor şablonunun müdürlük verisiyle (MudurlukDosyasi) doldurulmuş kopyasıdır; Açıklamalar sayfasına
örnek açıklamalar yazılır ve LibreOffice ile hesaplanır (Power Query formüllerin kayıtlı değerlerini okur).

Kullanım: python tools/pqtest/aciklama_ornek.py ÇIKTI_KLASÖRÜ   (soffice PATH'te olmalı)
"""

import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "python"))

AYAR = ROOT / "python" / "ornek" / "Ayarlar.xlsx"

# (dosya adı, müdürlük, kapanış, [(bölüm, gider kodu, açıklama, yazan, tarih)])
DOSYALAR = [
    ("Mal Kabul_2026-10.xlsx", "Mal Kabul", (2026, 10), [
        ("Giriş", 1, "Kadro planı Mart'a kaydı; boş pozisyonlar Kasım'da doluyor.", "A. Yılmaz", datetime(2026, 11, 5)),
        ("Tartım", 1, "Vardiya sayısı 3'ten 2'ye indi; tasarruf kalıcı.", "A. Yılmaz", datetime(2026, 11, 5)),
        ("Giriş", 99, "Geçersiz kod denemesi.", "A. Yılmaz", datetime(2026, 11, 5)),
    ]),
    ("Depolama_2026-10.xlsx", "Depolama", (2026, 10), [
        ("Yerleştirme", 1, "Yeni WMS ile yerleştirme ekibi küçüldü.", "B. Kaya", datetime(2026, 11, 6)),
        ("Sayım", 3, "Yıl sonu sayımı için fazla mesai öne çekildi.", "B. Kaya", datetime(2026, 11, 6)),
    ]),
    ("Sevkiyat_2026-10.xlsx", "Sevkiyat", (2026, 10), [
        ("Toplama", 1, "Sezonluk personel bütçelenenden az alındı.", "C. Demir", datetime(2026, 11, 7)),
    ]),
    ("Operasyon Destek_2026-10.xlsx", "Operasyon Destek", (2026, 10), []),
    ("Mal Kabul_2026-09.xlsx", "Mal Kabul", (2026, 9), [
        ("Giriş", 3, "Eylül: denetim öncesi fazla mesai.", "A. Yılmaz", datetime(2026, 10, 4)),
    ]),
]


def main(out):
    from bootstrap import dependencies

    dependencies()
    from openpyxl import Workbook

    import pq_sablon
    import pq_veri

    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as d:
        d = Path(d)
        for name, manager, closing, notes in DOSYALAR:
            wb = pq_sablon.build(pq_veri.build(AYAR, closing, manager))
            ws = wb["Açıklamalar"]
            for r, (dept, code, text, writer, when) in enumerate(notes, 2):
                ws[f"A{r}"], ws[f"B{r}"], ws[f"D{r}"], ws[f"E{r}"], ws[f"F{r}"] = dept, code, text, writer, when
                ws[f"F{r}"].number_format = "yyyy-mm-dd"
            source = d / name
            wb.save(source)
            subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                            "--outdir", str(out), str(source)], check=True, capture_output=True, timeout=600)
            print("üretildi:", out / name)
        bozuk = Workbook()
        bozuk.active.title = "Notlar"
        bozuk.active["A1"] = "Bu dosya rapor şablonundan değil."
        bozuk.save(out / "Notlar.xlsx")
        shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    main(sys.argv[1])
