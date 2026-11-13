"""Revize bütçe test verisi: örnek klasörün kopyasında Ekim kapanış dosyalarına "Revize Bütçe" bloğu ekler.

Blok, Bütçe başlığıyla aynı satırda, bütçe bloğunun iki sütun sağında başlar; altında 2026'nın 12 ayı ve her gider
için bütçenin %4 fazlası yazılır. Giriş bölümünde Şubat'ın 1. kalemi boş bırakılır (katı toplam testi).

Kullanım: python tools/pqtest/revize_ornek.py HEDEF_KLASÖR
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

    import butce

    hedef = Path(hedef)
    shutil.rmtree(hedef, ignore_errors=True)
    shutil.copytree(KAYNAK, hedef, ignore=shutil.ignore_patterns("pq_aciklama"))
    cfg = butce.read_settings(hedef / "Ayarlar.xlsx")["maliyet"]
    b_col, b_row, _, _ = butce.excel_bounds(cfg["butce_etiketi"], "Bütçe", allow_cell=True)
    first, size = cfg["ilk_gider_satiri"], butce.item_count(cfg)
    rows = list(range(first, first + size))
    start = b_col + 13  # bütçe bloğu (12 ay) + bir boş sütun
    for path in sorted((hedef / "maliyet").glob("Ekim 2026*.xlsx")):
        book = load_workbook(path)
        for ws in book:
            if ws.cell(b_row, b_col).value is None:
                continue
            ws.cell(b_row, start, "Revize Bütçe")
            for m in range(12):
                ws.cell(b_row + 1, start + m, ws.cell(b_row + 1, b_col + m).value)
                for r in rows:
                    v = ws.cell(r, b_col + m).value
                    blank = ws.title == "Giriş" and m == 1 and r == first
                    ws.cell(r, start + m, None if blank or v is None else round(v * 1.04, 2))
        book.save(path)
        print("revize eklendi:", path.name)


if __name__ == "__main__":
    main(sys.argv[1])
