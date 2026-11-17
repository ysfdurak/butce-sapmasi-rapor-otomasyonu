"""surum.py: ilk/güncel/önceki bütçe, bütçe değişimi ve trend serisi (küçük yapay veriyle)."""

import sys
import unittest
from decimal import Decimal as D
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import surum  # noqa: E402

CONFIG = {"bolumler": {"A": {"mudurluk": "M", "direktorluk": "Dir"}}, "kontrol": {"tolerans": 0.01}}
LAYOUT = [("A", "Bölüm", ["A"]), ("M", "Müdürlük toplamı", ["A"]), ("Dir", "Direktörlük toplamı", ["A"])]
RULES = {"onemlilik_tutari": D(1200)}


def snap(budget, actual, source):
    items = {1: ("Ücret", actual, budget)}
    return (items, None, {}, source)


def months(value, upto=12, change=None):
    out = {(2026, m): D(value) for m in range(1, upto + 1)}
    out.update({(2026, m): D(v) for m, v in (change or {}).items()})
    return out


class SurumTests(unittest.TestCase):
    def setUp(self):
        actual = months(100, 10, {8: 200, 9: 250, 10: 300})  # son üç ay bütçenin ≥ 100 üstünde
        self.snapshots = {
            ((2026, 9), "A"): snap(months(100), months(100, 9), "Eylül.xlsx"),
            ((2026, 10), "A"): snap(months(100, change={11: 150}), actual, "Ekim.xlsx"),
        }

    def test_first_budget_replaces_only_closing_budget(self):
        first = surum.with_first_budget(CONFIG, self.snapshots, (2026, 10))
        self.assertEqual(first[(2026, 10), "A"][0][1][2][(2026, 11)], D(100))
        self.assertEqual(self.snapshots[(2026, 10), "A"][0][1][2][(2026, 11)], D(150))
        self.assertIs(first[(2026, 9), "A"], self.snapshots[(2026, 9), "A"])

    def test_change_rows_against_previous_month_file(self):
        levels, items = surum.change_rows(CONFIG, self.snapshots, (2026, 10), LAYOUT, ["A"], 0.01)
        self.assertEqual([r[0] for r in levels], ["A", "M", "Dir"])
        self.assertEqual(levels[0][2 + 10], D(50))  # Kasım
        self.assertEqual(items[0][4], 1)  # değişen ay
        self.assertEqual(items[0][7], D(50))
        self.assertEqual(surum.change_rows(CONFIG, self.snapshots, (2026, 1), LAYOUT, ["A"], 0.01), ([], []))

    def test_trend_streak_ends_at_closing_month(self):
        levels, total, items = surum.trend_rows(CONFIG, self.snapshots, (2026, 10), LAYOUT, ["A"], 2026, RULES)
        self.assertEqual(items[0][4:7], [3, 3, D(450)])
        self.assertIsNone(total[10])  # kapanıştan sonraki ay
        self.assertEqual(surum.trend_rows(CONFIG, self.snapshots, (2026, 10), LAYOUT, ["A"], 2025, RULES)[2], [])

    def test_versions_and_comparison_setting(self):
        rows = surum.version_rows(CONFIG, self.snapshots, (2026, 10), LAYOUT, list(range(1, 11)), 2026)
        self.assertEqual(rows[0][3:6], [D(1000), D(1000), D(0)])
        self.assertEqual(surum.comparison({"karsilastirma_butcesi": "İlk bütçe"}), surum.FIRST)
        self.assertEqual(surum.comparison({}), surum.CURRENT)


if __name__ == "__main__":
    unittest.main()
