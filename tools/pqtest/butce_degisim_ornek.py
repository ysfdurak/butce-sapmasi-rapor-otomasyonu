"""Bütçe değişimi test verisi: örnek klasörün kopyasında Ekim kapanış dosyalarında bazı bütçeleri değiştirir.

Gerçek süreçteki gibi revize bütçe ayrı blokta değil, aynı bütçe hücrelerinde farklı değerle gelir:
  - Giriş, kalem 5: Kasım ve Aralık bütçesi %10 artar
  - Toplama, kalem 1: Ekim–Aralık bütçesine 250.000 TL eklenir
  - Sayım, kalem 3: Mart bütçesinden 100.000 TL düşülür (geçmiş ay revizyonu)
Eylül dosyası değişmez; böylece ilk bütçe (Eylül) ile güncel bütçe (Ekim) ve önceki ay bütçesi farklılaşır.

Kullanım: python tools/pqtest/butce_degisim_ornek.py HEDEF_KLASÖR
"""

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "python"))

KAYNAK = ROOT / "python" / "ornek"
DEGISIKLIK = {  # sayfa: [(kalem kodu, aylar, işlem)]
    "Giriş": [(5, [11, 12], lambda v: round(v * 1.10, 2))],
    "Toplama": [(1, [10, 11, 12], lambda v: round(v + 250_000, 2))],
    "Sayım": [(3, [3], lambda v: round(v - 100_000, 2))],
}


def main(hedef):
    from bootstrap import dependencies

    dependencies()
    from openpyxl import load_workbook

    import butce

    hedef = Path(hedef)
    shutil.rmtree(hedef, ignore_errors=True)
    shutil.copytree(KAYNAK, hedef, ignore=shutil.ignore_patterns("pq_aciklama"))
    cfg = butce.read_settings(hedef / "Ayarlar.xlsx")["maliyet"]
    b_col, _, _, _ = butce.excel_bounds(cfg["butce_etiketi"], "Bütçe", allow_cell=True)
    first = cfg["ilk_gider_satiri"]
    for path in sorted((hedef / "maliyet").glob("Ekim 2026*.xlsx")):
        book = load_workbook(path)
        changed = False
        for ws in book:
            for code, months, op in DEGISIKLIK.get(ws.title, []):
                for m in months:
                    cell = ws.cell(first + code - 1, b_col + m - 1)
                    cell.value = op(cell.value)
                    changed = True
        if changed:
            book.save(path)
            print("bütçe değiştirildi:", path.name)


if __name__ == "__main__":
    main(sys.argv[1])
