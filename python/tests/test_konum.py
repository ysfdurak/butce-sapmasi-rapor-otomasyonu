"""konum.py: konum ayarları olmadan veya kaymış kaynakta konumların dosyadan bulunması."""

import shutil
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import butce  # noqa: E402
import konum  # noqa: E402

ORNEK = Path(__file__).resolve().parents[1] / "ornek"
POSITIONS = {"Maliyet", "Kontrol", "Endeks"}


def stripped_settings(folder):
    """Örnek klasörün kopyası; Ayarlar'da Maliyet/Kontrol/Endeks konum satırları silinmiş."""
    for name in ("Kontrol.xlsx", "Endeks_2026_09.xlsx", "Endeks_2026_10.xlsx"):
        shutil.copy(ORNEK / name, folder / name)
    shutil.copytree(ORNEK / "maliyet", folder / "maliyet")
    book = load_workbook(ORNEK / "Ayarlar.xlsx")
    sheet = book["Dosya ve Konum Ayarları"]
    for r in range(sheet.max_row, 4, -1):
        if sheet.cell(r, 1).value in POSITIONS and sheet.cell(r, 2).value not in ("sayfa", "tolerans"):
            sheet.delete_rows(r)
    book.save(folder / "Ayarlar.xlsx")
    return folder / "Ayarlar.xlsx"


class KonumTests(unittest.TestCase):
    def setUp(self):
        self.folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.folder)
        self.full = butce.read_settings(ORNEK / "Ayarlar.xlsx")

    def test_sample_settings_are_not_overridden(self):
        self.assertEqual(self.full["otomatik_konum"], [])

    def test_detected_positions_equal_settings(self):
        config = butce.read_settings(stripped_settings(self.folder))
        for section in ("maliyet", "kontrol"):
            for code, value in config[section].items():
                self.assertEqual(value, self.full[section][code], f"{section}.{code}")
        self.assertEqual(config["endeks"]["endeks_araligi"], self.full["endeks"]["endeks_araligi"])
        self.assertEqual([n[0] for n in config["otomatik_konum"]], ["Maliyet", "Kontrol", "Endeks"])
        self.assertTrue(all(w[3].startswith("Konum otomatik bulundu") for w in butce.auto_layout_warnings(config)))

    def test_every_sample_sheet_is_detected(self):
        cfg = self.full["maliyet"]
        expected = {k: cfg[k] for k in konum.COST_KEYS}
        for path in sorted((ORNEK / "maliyet").glob("*.xlsx")):
            book = load_workbook(path, data_only=True, read_only=True)
            for sheet in book:
                if butce.key(sheet.title) in {butce.key(n) for n in self.full["bolumler"]}:
                    self.assertEqual(konum.cost_layout(sheet), expected, f"{path.name} › {sheet.title}")
            book.close()

    def test_shifted_sheet_falls_back_and_warns(self):
        path = ORNEK / "maliyet" / "Ekim 2026 - Dosya 1.xlsx"
        book = load_workbook(path, data_only=True)
        sheet = next(s for s in book if butce.key(s.title) in {butce.key(n) for n in self.full["bolumler"]})
        before = butce.read_cost(sheet, self.full["maliyet"], (2026, 10))
        sheet.insert_rows(1, 2)
        sheet.insert_cols(1)
        issues = butce.SourceIssues()
        after = butce.read_cost(sheet, self.full["maliyet"], (2026, 10), issues)
        self.assertEqual(before, after)
        notes = issues.warnings(sheet.title)
        self.assertTrue(any("otomatik bulundu" in n and "BG61" in n for n in notes), notes)

    def test_missing_layout_is_reported(self):
        settings = stripped_settings(self.folder)
        for path in (self.folder / "maliyet").glob("*.xlsx"):
            path.unlink()
        with self.assertRaisesRegex(ValueError, "Konum dosyadan bulunamadı.*tarih_baslangici"):
            butce.read_settings(settings)


if __name__ == "__main__":
    unittest.main()
