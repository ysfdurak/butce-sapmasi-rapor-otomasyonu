"""Source titles, all mapping rows, persisted locations and safe Excel updates."""

from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch
import copy, shutil, time, unittest
from openpyxl import load_workbook
import butce, configuration
from model import Dataset
from app import Application

BASE = Path(__file__).resolve().parents[1]
EXAMPLE = BASE / "ornek"
STAFF = ["Kadrolu çalışan sayısı", "Taşeron çalışan sayısı", "Fiili çalışan sayısı", "Planlı çalışan sayısı"]


class SheetLabels:
    def __init__(self, sheet, column=57, labels=STAFF, different=False, blank=False):
        self.sheet = sheet
        self.column = column
        self.labels = labels
        self.different = different
        self.blank = blank

    def cell(self, row, col):
        original = self.sheet.cell(row, col)
        if col == self.column and 60 <= row <= 64:
            value = self.labels[row - 60] if row < 64 else "İşlenen koli adedi"
            if self.different and self.sheet.title == "Giriş" and row == 60:
                value = "Operatör sayısı"
            if self.blank and row == 60:
                value = None
            return SimpleNamespace(value=value, coordinate=original.coordinate)
        return original

    def __getattr__(self, name):
        return getattr(self.sheet, name)


class BookLabels:
    def __init__(self, path, **kwargs):
        self.book = load_workbook(path, data_only=True)
        self.kwargs = kwargs

    def __getitem__(self, name):
        return SheetLabels(self.book[name], **self.kwargs)

    def close(self):
        self.book.close()


class SourceTitleTests(unittest.TestCase):
    def test_actual_titles_all_four_and_transaction(self):
        with patch("model.load_workbook", side_effect=lambda path, **kwargs: BookLabels(path)):
            data = Dataset(EXAMPLE / "Ayarlar.xlsx", (2026, 10))
        q = data.query({})
        self.assertEqual([s["name"] for s in q["staff"]], STAFF)
        self.assertEqual(q["transactions"]["name"], "İşlenen koli adedi")
        self.assertEqual(q["general"]["cards"]["actual"], butce.Decimal("3449.38339031"))

    def test_different_department_titles_are_scoped(self):
        with patch(
            "model.load_workbook", side_effect=lambda path, **kwargs: BookLabels(path, different=True)
        ):
            data = Dataset(EXAMPLE / "Ayarlar.xlsx", (2026, 10))
        self.assertEqual(
            data.query({"filters": {"department": ["Giriş"]}})["staff"][0]["name"], "Operatör sayısı"
        )
        all_data = data.query({})
        self.assertEqual(len(all_data["staff"][0]["titleVariants"]), 2)
        self.assertTrue(any("adları farklı" in message for message in all_data["messages"]))

    def test_missing_title_is_visible(self):
        with patch("model.load_workbook", side_effect=lambda path, **kwargs: BookLabels(path, blank=True)):
            data = Dataset(EXAMPLE / "Ayarlar.xlsx", (2026, 10))
        self.assertIn("Başlığı eksik", data.query({})["staff"][0]["name"])
        self.assertTrue(any("Başlığı eksik" in row[3] for row in data.warnings))

    def test_optional_title_column_is_parametric(self):
        original = butce.read_settings
        cfg = original(EXAMPLE / "Ayarlar.xlsx")
        cfg["maliyet"]["gosterge_adi_sutunu"] = 2
        with (
            patch("butce.read_settings", return_value=cfg),
            patch("model.load_workbook", side_effect=lambda path, **kwargs: BookLabels(path, column=2)),
        ):
            data = Dataset(EXAMPLE / "Ayarlar.xlsx", (2026, 10))
        self.assertEqual(data.query({})["staff"][2]["name"], STAFF[2])


class ConfigurationTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.root = Path(self.temp.name)
        shutil.copytree(EXAMPLE, self.root / "ornek")
        self.settings = self.root / "ornek/Ayarlar.xlsx"

    def tearDown(self):
        self.temp.cleanup()

    def preview(self):
        return configuration.preview(self.settings)

    def test_preview_contains_every_mapping_field(self):
        p = self.preview()
        self.assertEqual(len(p["mappings"]), 12)
        self.assertEqual(
            p["mappingCounts"],
            {"departments": 12, "managers": 4, "groups": 6, "expenseMappings": 2, "excludedSheets": 2},
        )
        self.assertEqual(
            set(p["mappings"][0]),
            {
                "department",
                "mudurluk",
                "direktorluk",
                "endeks",
                "lokasyon",
                "kontrol_bolumu",
                "mudurluk_alti",
            },
        )
        self.assertEqual(len(p["expenseGroups"]), 6)
        self.assertEqual(len(p["expenseMappings"]), 2)
        self.assertEqual(len(p["excludedSheets"]), 2)

    def test_locations_persist_and_backup_preserves_all_mappings(self):
        p = self.preview()
        before = self.settings.read_bytes()
        old = butce.read_settings(self.settings)
        data, backup = configuration.update_locations(
            self.settings, p["settings"], p["fingerprint"], p["locations"], [2026, 10], Dataset
        )
        self.assertEqual(Path(backup).read_bytes(), before)
        cfg = butce.read_settings(self.settings)
        for key in ["bolumler", "gider_gruplari", "gider_eslestirme", "ortak_sayfalar"]:
            self.assertEqual(cfg[key], old[key])
        self.assertTrue(all(Path(cfg[v]).is_absolute() for v in configuration.KEYS.values()))
        self.assertEqual(data.settings, self.settings.resolve())
        self.assertEqual(
            data.identity, configuration.hashlib.sha256(str(self.settings.resolve()).encode()).hexdigest()
        )

    def test_invalid_source_does_not_modify_settings(self):
        p = self.preview()
        before = self.settings.read_bytes()
        bad = dict(p["locations"], costFolder=str(self.root / "missing"))
        with self.assertRaises(ValueError):
            configuration.update_locations(
                self.settings, p["settings"], p["fingerprint"], bad, [2026, 10], Dataset
            )
        self.assertEqual(self.settings.read_bytes(), before)
        self.assertFalse((self.settings.parent / "Ayar_Yedekleri").exists())

    def test_failed_candidate_does_not_modify_settings(self):
        p = self.preview()
        before = self.settings.read_bytes()

        def reject(*args):
            raise ValueError("Yeni kaynaklar uyumsuz")

        with self.assertRaises(ValueError):
            configuration.update_locations(
                self.settings, p["settings"], p["fingerprint"], p["locations"], [2026, 10], reject
            )
        self.assertEqual(self.settings.read_bytes(), before)
        self.assertFalse(list(self.settings.parent.glob(".butce_ayar_*.xlsx")))

    def test_external_change_does_not_get_overwritten(self):
        p = self.preview()
        b = load_workbook(self.settings)
        b["Bölüm Eşleştirme"]["B5"] = "Değişen müdürlük"
        b.save(self.settings)
        b.close()
        changed = self.settings.read_bytes()
        with self.assertRaises(ValueError):
            configuration.update_locations(
                self.settings, p["settings"], p["fingerprint"], p["locations"], [2026, 10], Dataset
            )
        self.assertEqual(self.settings.read_bytes(), changed)

    def test_new_department_and_manager_included_after_reload(self):
        b = load_workbook(self.settings)
        b["Bölüm Eşleştirme"].append(
            [
                "Yeni Bölüm",
                "Yeni Müdürlük",
                "Lojistik Direktörlüğü",
                "Endeks Giriş",
                "Yeni Lokasyon",
                "Yeni Kontrol",
                "Yeni Alt",
            ]
        )
        b.save(self.settings)
        b.close()
        source = load_workbook(self.settings.parent / "maliyet/Ekim 2026 - Dosya 1.xlsx")
        for name in list(source.sheetnames):
            if name != "Giriş":
                del source[name]
        source["Giriş"].title = "Yeni Bölüm"
        source.save(self.settings.parent / "maliyet/Ekim 2026 - Ek Bölüm.xlsx")
        source.close()
        data = Dataset(self.settings, (2026, 10))
        meta = data.meta()
        q = data.query({})
        self.assertEqual(len(meta["options"]["department"]), 13)
        self.assertEqual(len(meta["options"]["manager"]), 5)
        self.assertIn("Yeni Bölüm", q["selection"]["departments"])
        self.assertIn("Yeni Lokasyon", meta["options"]["location"])
        self.assertIn("Yeni Alt", meta["options"]["submanager"])
        self.assertIn("Yeni Müdürlük Toplamı", [row["name"] for row in q["departments"]["expense"]])

    def test_unmapped_source_is_warned_without_guessed_manager(self):
        source = load_workbook(self.settings.parent / "maliyet/Ekim 2026 - Dosya 1.xlsx")
        for name in list(source.sheetnames):
            if name != "Giriş":
                del source[name]
        source["Giriş"].title = "Eşleşmeyen Bölüm"
        source.save(self.settings.parent / "maliyet/Ekim 2026 - Ek Bölüm.xlsx")
        source.close()
        data = Dataset(self.settings, (2026, 10))
        self.assertNotIn("Eşleşmeyen Bölüm", data.departments)
        self.assertTrue(any(r[1] == "Eşleşmeyen Bölüm" and "eşleştirmesi" in r[3] for r in data.warnings))

    def test_deleted_and_renamed_mapping_reflects_in_filters(self):
        b = load_workbook(self.settings)
        s = b["Bölüm Eşleştirme"]
        s.delete_rows(5)
        for row in s.iter_rows(min_row=5):
            if row[1].value == "Mal Kabul":
                row[1].value = "Yeni kabul müdürlüğü"
        b.save(self.settings)
        b.close()
        data = Dataset(self.settings, (2026, 10))
        meta = data.meta()
        self.assertNotIn("Giriş", meta["options"]["department"])
        self.assertNotIn("Mal Kabul", meta["options"]["manager"])
        self.assertIn("Yeni kabul müdürlüğü", meta["options"]["manager"])
        self.assertEqual(len(data.departments), 11)

    def test_configure_job_remembers_settings_path_on_restart(self):
        state = self.root / "state"
        app = Application(self.settings, (2026, 10), state)
        try:
            p = self.preview()
            app.start_job(
                "configure",
                {
                    "settings": str(self.settings),
                    "expectedSettings": p["settings"],
                    "fingerprint": p["fingerprint"],
                    "locations": p["locations"],
                    "closing": [2026, 10],
                },
            )
            limit = time.monotonic() + 20
            while app.job["state"] == "running" and time.monotonic() < limit:
                time.sleep(0.02)
            self.assertEqual(app.job["state"], "done", app.job)
        finally:
            app.executor.shutdown()
        restarted = Application(state_dir=state)
        try:
            self.assertEqual(restarted.data.settings, self.settings.resolve())
            self.assertEqual(restarted.meta()["configuration"]["locations"], p["locations"])
        finally:
            restarted.executor.shutdown()

    def test_all_selection_includes_additions_and_deleted_names_are_removed(self):
        old = {"department": ["A", "B"], "manager": ["M"], "year": [2026], "month": [1, 2]}
        new = {"department": ["B", "C"], "manager": ["N"], "year": [2026], "month": [1, 2]}
        state = {
            "filters": {"department": ["A", "B"], "manager": ["M"], "year": [2026], "month": [1, 2]},
            "bmLocal": ["A", "B"],
            "opexLocal": ["B"],
        }
        actual = configuration.reconcile_selection(state, old, new)
        self.assertIsNone(actual["filters"]["department"])
        self.assertIsNone(actual["bmLocal"])
        self.assertIsNone(actual["filters"]["manager"])
        self.assertEqual(actual["opexLocal"], ["B"])
        self.assertEqual(actual["filters"]["year"], [2026])
        self.assertEqual(actual["filters"]["month"], [1, 2])


if __name__ == "__main__":
    unittest.main()
