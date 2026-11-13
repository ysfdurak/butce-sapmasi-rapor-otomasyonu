"""Power Query yolu: veri sözleşmesi (tools/pq_veri.py) ve formül şablonu (tools/pq_sablon.py).

Gerçek M motoruyla karşılaştırma ve LibreOffice ile formül hesabı ayrı CI işlerindedir (tools/pqtest).
"""

import sys
import unittest
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))

import pq_sablon  # noqa: E402
import pq_veri  # noqa: E402

SETTINGS = ROOT / "python/ornek/Ayarlar.xlsx"


class PowerQueryDataTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.rows = pq_veri.build(SETTINGS)

    def test_blocks_and_unique_keys(self):
        counts = Counter(r[1] for r in self.rows)
        self.assertEqual(counts["Bolum"], 12)
        self.assertEqual(counts["Revizyon"], 26)
        self.assertEqual(counts["Uyari"], 96)  # ardışık dosya eksik: Ocak–Ağustos × 12 bölüm
        self.assertEqual(len({r[0] for r in self.rows}), len(self.rows))
        self.assertTrue(all(len(r) == len(pq_veri.COLUMNS) for r in self.rows))

    def test_strict_month_total(self):
        index = {r[0]: r for r in self.rows}
        g = pq_veri.COLUMNS.index("G1")
        items = [r for r in self.rows if r[1] == "Gider" and r[4] == "Giriş" and r[8] == 2026]
        total = index["Toplam|Giriş|2026"]
        self.assertAlmostEqual(total[g], sum(r[g] for r in items), places=2)
        self.assertIsNone(pq_veri.strict_sum([1, None]))  # bir kalem boşsa ay toplamı boş

    def test_template_builds_empty_and_filled(self):
        empty = pq_sablon.build()
        self.assertIsNone(empty["Veri"]["A1"].value)  # sorgu A1'e yüklenir
        self.assertIn("Kontrol Paneli", empty.sheetnames)
        filled = pq_sablon.build(self.rows)
        self.assertEqual(filled["Veri"].max_row, len(self.rows) + 1)


if __name__ == "__main__":
    unittest.main()
