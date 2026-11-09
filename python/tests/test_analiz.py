"""Önemli sapma, yıl sonu tahmini, hacim/birim maliyet, yönetici özeti ve kaynak sağlamlığı testleri."""

import shutil
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path
from types import SimpleNamespace

from openpyxl import load_workbook

import analiz
import butce

BASE = Path(__file__).resolve().parents[1]
SETTINGS = BASE / "ornek/Ayarlar.xlsx"
RULES = analiz.settings({})


def detail_row(name, code, actual, budget, manager="M"):
    return [name, manager, code, f"Kalem {code}", "G", actual, None, budget, None, None]


class TextNumberTests(unittest.TestCase):
    def test_turkish_and_dot_decimal_text(self):
        cases = {
            "1.234,56": D("1234.56"),
            "1234,5": D("1234.5"),
            "1.234": D("1234"),
            "12.5": D("12.5"),
            "-2.000.000": D("-2000000"),
            "(1.500,25)": D("-1500.25"),
            " 7 500 TL ": D("7500"),
            "₺10": D("10"),
        }
        for text, expected in cases.items():
            self.assertEqual(butce.text_number(text), expected, text)

    def test_non_numbers_stay_empty(self):
        for text in ("", "abc", "1,2,3", "12-3", "--", "1.23.4", None, 5):
            self.assertIsNone(butce.text_number(text), text)


class MaterialityTests(unittest.TestCase):
    def test_amount_and_rate_must_both_pass(self):
        rows = analiz.material_rows(
            [
                detail_row("A", 1, D(2_000_000), D(1_000_000)),  # +1M, %100 → önemli, aleyhte
                detail_row("A", 2, D(10_900_000), D(10_000_000)),  # +900K → tutar altında
                detail_row("A", 3, D(51_000_000), D(50_000_000)),  # +1M ama %2 → oran altında
                detail_row("B", 1, D(0), D(3_000_000)),  # −3M, −%100 → önemli, lehte
                detail_row("B", 2, D(1_500_000), D(0)),  # bütçe 0 → oran tanımsız, yalnız tutar
                detail_row("B", 3, None, D(5_000_000)),  # eksik veri → değerlendirilmez
            ],
            RULES,
        )
        self.assertEqual([(r[1], r[2]) for r in rows], [("B", 1), ("B", 2), ("A", 1)])
        self.assertEqual([r[9] for r in rows], ["Lehte", "Aleyhte", "Aleyhte"])
        self.assertIsNone(rows[1][8])

    def test_direction_and_threshold_text(self):
        self.assertEqual(analiz.direction(D(1)), "Aleyhte")
        self.assertEqual(analiz.direction(D(-1)), "Lehte")
        self.assertEqual(analiz.direction(D(0)), "Bütçede")
        self.assertIsNone(analiz.direction(None))
        self.assertEqual(analiz.threshold_text(RULES), "≥ 1.000.000 TL ve ≥ %10")
        custom = analiz.settings({"rapor": {"onemlilik_tutari": 250000, "onemlilik_orani": 0.075}})
        self.assertEqual(analiz.threshold_text(custom), "≥ 250.000 TL ve ≥ %7,5")


class SyntheticAnalysisTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config = butce.read_settings(SETTINGS)
        cls.warnings = []
        cls.snapshots = butce.read_snapshots(
            SETTINGS.parent / cls.config["maliyet_klasoru"], cls.config, (2026, 10), cls.warnings
        )
        cls.layout = list(butce.hierarchy(cls.config["bolumler"]))
        from birim import read_budget_counts

        cls.counts = read_budget_counts(SETTINGS.parent, cls.config, cls.snapshots, (2026, 10))

    def test_forecast_identities(self):
        rows = analiz.forecast_rows(self.config, self.snapshots, (2026, 10), self.layout)
        total = rows[-1]
        self.assertEqual(total[1], "Direktörlük toplamı")
        self.assertEqual(total[2], D("3449383390.31"))  # Ocak–Ekim gerçekleşen
        self.assertEqual(total[3], D("3440938980.13"))  # Ocak–Ekim bütçe
        self.assertEqual(total[5], D("4108586425.59"))  # yıllık bütçe
        self.assertEqual(total[6], D("4117030835.77"))  # gerçekleşen + kalan bütçe
        self.assertEqual(total[7], total[4])  # kalan aylar bütçede kalınca yıl sonu farkı = bugünkü fark
        self.assertEqual(total[8], "Aleyhte")
        self.assertEqual(total[9], total[2] * 12 / 10)
        self.assertEqual(sum(r[2] for r in rows if r[1] == "Bölüm"), total[2])

    def test_volume_and_unit_cost_effects_sum_to_variance(self):
        rows = analiz.volume_price_rows(
            self.config, self.snapshots, (2026, 10), list(range(1, 11)), self.counts, self.layout
        )
        departments = [r for r in rows if r[1] == "Bölüm"]
        self.assertEqual(len(departments), 12)
        for row in rows:
            self.assertIsNotNone(row[9], row[0])
            self.assertLess(abs(row[9] + row[10] - row[4]), D("0.000001"), row[0])
        giris = departments[0]
        self.assertEqual((giris[0], giris[5], giris[6]), ("Giriş", D(5070772), D(5660000)))

    def test_material_rows_on_synthetic_data(self):
        detail, _, _ = butce.make_report(
            self.config, (2026, 10), (2026, 10), list(range(1, 11)), self.snapshots, {}, {}, {}, []
        )
        rows = analiz.material_rows(detail, RULES)
        self.assertEqual(len(rows), 150)
        self.assertEqual(sum(r[9] == "Aleyhte" for r in rows), 75)
        self.assertEqual((rows[0][1], rows[0][2], rows[0][7]), ("Toplama", 1, D("-20764946.98")))

    def test_year_boundary_revision_pairs(self):
        self.assertEqual(
            [(a, b, [k for k, _, _ in c]) for a, b, c in butce.revision_pairs(2027, 1)],
            [((2026, 12), (2027, 1), ["Gerçekleşen"])],
        )
        pairs = list(butce.revision_pairs(2026, 3))
        self.assertEqual([(a, b) for a, b, _ in pairs], [((2026, 1), (2026, 2)), ((2026, 2), (2026, 3))])
        self.assertTrue(butce.in_closing_scope((2026, 12), (2027, 1)))
        self.assertFalse(butce.in_closing_scope((2026, 12), (2027, 2)))
        self.assertFalse(butce.in_closing_scope((2026, 11), (2027, 1)))


class ReportOutputTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.folder = Path(tempfile.mkdtemp())
        cls.paths = butce.run(
            SimpleNamespace(
                ayar=SETTINGS,
                donem="2026-10",
                kapanis="2026-10",
                aylar="1,2,3,4,5,6,7,8,9,10",
                cikti=cls.folder / "Genel.xlsx",
            )
        )

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.folder, ignore_errors=True)

    def test_new_sheets_and_executive_first(self):
        book = load_workbook(self.paths[0])
        self.assertEqual(book.sheetnames[0], "Yönetici Özeti")
        self.assertEqual(book.active.title, "Yönetici Özeti")
        self.assertEqual(sum(bool(ws.sheet_view.tabSelected) for ws in book.worksheets), 1)
        for name in (
            "Önemli Sapmalar",
            "Yıl Sonu Tahmini",
            "Hacim ve Birim Maliyet",
            "Genel Özet",
            "Giderler",
        ):
            self.assertIn(name, book.sheetnames)
        summary = {
            r[0]: r[1] for r in book["Yönetici Özeti"].iter_rows(min_row=5, max_row=16, values_only=True)
        }
        self.assertAlmostEqual(summary["Gerçekleşen − bütçe (TL)"], 8444410.18, places=2)
        self.assertEqual(summary["Yön"], "Aleyhte")
        self.assertEqual(summary["Önemli sapma"], "150 kalem (75 aleyhte)")
        self.assertEqual(book["Yönetici Özeti"].page_setup.orientation, "landscape")

    def test_manager_report_requires_explanation_for_material_items(self):
        book = load_workbook(next(p for p in self.paths if "Depolama" in p.name))
        sheet = book["Önemli Sapmalar"]
        formula = sheet.cell(2, 11).value
        self.assertTrue(formula.startswith("=IF(LEN(TRIM(Giderler!"), formula)
        address = formula.split("Giderler!")[1].split(")")[0]
        self.assertEqual(book["Giderler"][address].fill.fgColor.rgb[-6:], "F8CBAD")
        identity, members, columns, rows = butce.explanation_layout(book, "Müdürlük")
        self.assertEqual(len(members), 3)


class SettingsTests(unittest.TestCase):
    def copy_settings(self, edit):
        folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, folder, True)
        path = folder / "Ayarlar.xlsx"
        shutil.copy(SETTINGS, path)
        book = load_workbook(path)
        edit(book["Dosya ve Konum Ayarları"])
        book.save(path)
        return path

    def test_report_settings_and_validation(self):
        def put(rows):
            def edit(sheet):
                for section, code, value in rows:
                    row = next(
                        (r for r in range(5, sheet.max_row + 1) if sheet.cell(r, 2).value == code), None
                    )
                    if row is None:
                        sheet.append([section, code, value])
                    else:
                        sheet.cell(row, 3).value = value

            return edit

        self.assertEqual(analiz.settings(butce.read_settings(SETTINGS))["onemlilik_tutari"], D(1000000))
        path = self.copy_settings(
            put([("Rapor", "onemlilik_tutari", 250000), ("Rapor", "onemlilik_orani", 0.05)])
        )
        rules = analiz.settings(butce.read_settings(path))
        self.assertEqual((rules["onemlilik_tutari"], rules["onemlilik_orani"]), (D(250000), D("0.05")))
        for bad in (
            ("Rapor", "onemlilik_orani", 10),
            ("Rapor", "onemlilik_tutari", -1),
            ("Rapor", "bilinmeyen", 1),
        ):
            with self.assertRaises(ValueError):
                butce.read_settings(self.copy_settings(put([bad])))

    def test_legacy_layout_keys_optional_with_positioned_layout(self):
        legacy = {
            "yil_satiri",
            "ay_satiri",
            "indeks_sutunu",
            "ilk_donem_sutunu",
            "ciro_satiri",
            "ilk_blok_satiri",
            "ilk_ay_sutunu",
            "ilk_veri_satiri",
            "gerceklesen_araligi",
            "butce_araligi",
            "ciro_araligi",
        }

        def drop(sheet):
            for r in range(sheet.max_row, 4, -1):
                if sheet.cell(r, 2).value in legacy:
                    sheet.delete_rows(r)

        config = butce.read_settings(self.copy_settings(drop))
        self.assertNotIn("yil_satiri", config["maliyet"])

        def drop_positioned(sheet):
            for r in range(sheet.max_row, 4, -1):
                if sheet.cell(r, 2).value in {"tarih_baslangici", "yil_satiri"}:
                    sheet.delete_rows(r)

        with self.assertRaisesRegex(ValueError, "yil_satiri"):
            butce.read_settings(self.copy_settings(drop_positioned))


class SourceRobustnessTests(unittest.TestCase):
    def copy_sources(self):
        folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, folder, True)
        shutil.copytree(BASE / "ornek", folder / "ornek")
        return folder / "ornek"

    def test_text_numbers_are_converted_with_cell_addresses(self):
        root = self.copy_sources()
        path = root / "maliyet/Ekim 2026 - Dosya 1.xlsx"
        book = load_workbook(path)
        sheet = book.worksheets[2]
        original = sheet["BF65"].value
        sheet["BF65"] = f"{original:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        sheet["BG65"] = "yok"
        book.save(path)
        config = butce.read_settings(root / "Ayarlar.xlsx")
        warnings = []
        snaps = butce.read_snapshots(root / "maliyet", config, (2026, 10), warnings)
        messages = [w[3] for w in warnings if w[2] == path.name]
        self.assertTrue(any("sayıya çevrildi" in m and "BF65" in m for m in messages), messages)
        self.assertTrue(any("boş sayıldı" in m and "BG65" in m for m in messages), messages)
        name = next(n for (p, n) in snaps if p == (2026, 10) and snaps[(p, n)][3] == path.name)
        self.assertIsNotNone(snaps[(2026, 10), name])

    def test_uncached_formula_is_reported(self):
        root = self.copy_sources()
        path = root / "maliyet/Ekim 2026 - Dosya 2.xlsx"
        book = load_workbook(path)
        book.worksheets[2]["BF66"] = "=1+1"  # openpyxl hesaplanmış sonucu kaydetmez
        book.save(path)
        config = butce.read_settings(root / "Ayarlar.xlsx")
        warnings = []
        butce.read_snapshots(root / "maliyet", config, (2026, 10), warnings)
        self.assertTrue(any("formül hücresinin sonucu" in w[3] and "BF66" in w[3] for w in warnings))

    def test_shifted_layout_suggests_new_cells(self):
        root = self.copy_sources()
        path = root / "maliyet/Ekim 2026 - Dosya 3.xlsx"
        book = load_workbook(path)
        sheet = book.worksheets[2]
        sheet.insert_rows(1, 2)
        book.save(path)
        config = butce.read_settings(root / "Ayarlar.xlsx")
        warnings = []
        butce.read_snapshots(root / "maliyet", config, (2026, 10), warnings)
        hint = [w[3] for w in warnings if w[2] == path.name and "kaymış" in w[3]]
        self.assertTrue(hint, [w[3] for w in warnings if w[2] == path.name])
        self.assertIn("BF61", hint[0])
        self.assertIn("EZ60", hint[0])


class AuditTrailTests(unittest.TestCase):
    def test_notes_record_who_and_when(self):
        from model import Dataset
        from notes import Notes

        folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, folder, True)
        data = Dataset(SETTINGS, (2026, 10))
        notes = Notes(folder / "kayit.sqlite3")
        notes.save(data, [10], [{"department": "Giriş", "code": 1, "note": "Zam etkisi"}])
        record = notes.list(data, [10])["Giriş", 1]
        self.assertTrue(record["lastAction"].startswith("Kaydedildi"))
        self.assertTrue(record["actor"])
        self.assertTrue(record["actedAt"])


if __name__ == "__main__":
    unittest.main()
