"""surum.py: ilk/güncel/önceki bütçe, bütçe değişimi ve trend serisi (küçük yapay veriyle)."""

import sys
import unittest
from decimal import Decimal as D
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import surum  # noqa: E402
import butce  # noqa: E402

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


class DashboardBasisTests(unittest.TestCase):
    """Pano, Excel raporlarıyla aynı karşılaştırma bütçesini kullanır (Ayarlar › karsilastirma_butcesi)."""

    def test_dashboard_uses_first_budget_when_selected(self):
        from pathlib import Path
        from decimal import Decimal as D
        import analiz
        from model import Dataset

        base = Path(__file__).resolve().parents[1] / "ornek/Ayarlar.xlsx"
        current = Dataset(base, (2026, 10))
        self.assertEqual(current.meta()["budgetBasis"], "current")
        real_settings, real_first = analiz.settings, surum.with_first_budget

        def doubled(config, snapshots, closing):
            # Örnek veride ilk ve güncel bütçe aynı; bağlantıyı görmek için ilk bütçeyi iki katına çıkar.
            out = real_first(config, snapshots, closing)
            return {
                k: ({c: (lab, a, {p: (v * 2 if v is not None else None) for p, v in b.items()}) for c, (lab, a, b) in s[0].items()}, *s[1:])
                if k[0] == closing else s
                for k, s in out.items()
            }

        analiz.settings = lambda c: {**real_settings(c), "karsilastirma_butcesi": "İlk bütçe"}
        surum.with_first_budget = doubled
        try:
            first = Dataset(base, (2026, 10))
        finally:
            analiz.settings, surum.with_first_budget = real_settings, real_first
        self.assertEqual(first.meta()["budgetBasis"], "first")
        a = current.query({})["general"]["cards"]["budget"]
        b = first.query({})["general"]["cards"]["budget"]
        self.assertEqual(b, a * 2)
        self.assertEqual(len(first.revisions), len(current.revisions))  # revizyonlar kaynak dosyalardan


class DashboardEqualsReportTests(unittest.TestCase):
    """Pano = rapor: panodaki toplam gerçekleşen ve bütçe, Excel rapor motorunun (butce.run ile aynı adımlar)
    satırlarının toplamına eşit olmalı. İki bütçe temelinde de denenir; ilk bütçe, farkı görmek için iki katına
    çıkarılır. Birinden biri temeli yok sayarsa (Kasım 2026'da panoda olan hata) test düşer."""

    def _report_totals(self, base, closing, months):
        import analiz

        config = butce.read_settings(base)
        snapshots = butce.read_snapshots(base.parent / config["maliyet_klasoru"], config, closing, [])
        if surum.comparison(analiz.settings(config)) == surum.FIRST:
            snapshots = surum.with_first_budget(config, snapshots, closing)
        detail, _, _ = butce.make_report(config, closing, closing, months, snapshots, {}, {}, {}, [])
        actual = sum((r[5] or 0) for r in detail)
        planned = sum((r[7] or 0) for r in detail)
        return actual / 10**6, planned / 10**6

    def _check(self, first):
        from pathlib import Path
        import analiz
        from model import Dataset

        base = Path(__file__).resolve().parents[1] / "ornek/Ayarlar.xlsx"
        real_settings, real_first = analiz.settings, surum.with_first_budget

        def doubled(config, snapshots, closing):
            out = real_first(config, snapshots, closing)
            return {
                k: ({c: (lab, a, {p: (v * 2 if v is not None else None) for p, v in b.items()}) for c, (lab, a, b) in s[0].items()}, *s[1:])
                if k[0] == closing else s
                for k, s in out.items()
            }

        if first:
            analiz.settings = lambda c: {**real_settings(c), "karsilastirma_butcesi": "İlk bütçe"}
            surum.with_first_budget = doubled
        try:
            cards = Dataset(base, (2026, 10)).query({})["general"]["cards"]
            actual, planned = self._report_totals(base, (2026, 10), list(range(1, 11)))
        finally:
            analiz.settings, surum.with_first_budget = real_settings, real_first
        self.assertEqual(cards["actual"], actual)
        self.assertEqual(cards["budget"], planned)
        return cards

    def test_current_budget(self):
        self._check(False)

    def test_first_budget(self):
        current = self._check(False)
        first = self._check(True)
        self.assertEqual(first["budget"], current["budget"] * 2, "ilk bütçe gerçekten kullanıldı")
