"""Konum satırları silinmiş Ayarlar ile örnek klasör (V5.5 otomatik konum testi).

Örnek klasörün kopyasında Ayarlar › Dosya ve Konum Ayarları'ndan Maliyet, Kontrol ve Endeks konum satırları
silinir (sayfa adı ve tolerans kalır). Python, makro ve Power Query konumları dosyalardan bulmalı ve aynı
sonucu vermelidir.

Kullanım: python tools/pqtest/konumsuz_ornek.py HEDEF_KLASÖR
"""

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

KAYNAK = ROOT / "python" / "ornek"


def main(hedef):
    from bootstrap import dependencies

    dependencies()
    from openpyxl import load_workbook

    hedef = Path(hedef)
    shutil.rmtree(hedef, ignore_errors=True)
    shutil.copytree(KAYNAK, hedef, ignore=shutil.ignore_patterns("pq_aciklama"))
    book = load_workbook(hedef / "Ayarlar.xlsx")
    sheet = book["Dosya ve Konum Ayarları"]
    silinen = []
    for r in range(sheet.max_row, 4, -1):
        alan, kod = sheet.cell(r, 1).value, sheet.cell(r, 2).value
        if alan in ("Maliyet", "Kontrol", "Endeks") and kod not in ("sayfa", "tolerans"):
            silinen.append(f"{alan}|{kod}")
            sheet.delete_rows(r)
    book.save(hedef / "Ayarlar.xlsx")
    print(f"{len(silinen)} konum satırı silindi: {hedef / 'Ayarlar.xlsx'}")


if __name__ == "__main__":
    main(sys.argv[1])
