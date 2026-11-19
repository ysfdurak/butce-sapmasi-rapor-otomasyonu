"""Kalem ayrıntısı: gider satırı bazında bütçe, gerçekleşen, hacim / birim maliyet etkisi ve kaynak hücreleri."""

import json
import tempfile
import threading
import unittest
import urllib.request
from decimal import Decimal as D
from pathlib import Path

import analiz
import birim
import kalem
from app import Application, Server
from model import Dataset

BASE = Path(__file__).resolve().parents[1]
SETTINGS = BASE / "ornek/Ayarlar.xlsx"
KURUS = D("0.01")


class LineDetailTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data = Dataset(SETTINGS, (2026, 10))
        cls.state = {"filters": {}}

    def test_effects_add_up_to_variance(self):
        r = kalem.detail(self.data, self.state, 1)
        self.assertEqual(r["name"], "Ücret giderleri")
        self.assertEqual(r["actual"], D("845437105.52"))
        self.assertEqual(r["budget"], D("843367397.13"))
        self.assertEqual(r["variance"], D("2069708.39"))
        self.assertEqual(r["direction"], "Aleyhte")
        self.assertEqual(r["residual"], D("0.00"))
        for row in r["departments"]:
            self.assertLess(abs(row["volume"] + row["price"] - row["variance"]), KURUS, row["department"])

    def test_lines_sum_to_department_volume_price_sheet(self):
        """Bir bölümün bütün kalemleri toplandığında Hacim ve Birim Maliyet sayfasının bölüm satırı çıkar."""
        names, _, months = self.data.selection(self.state)
        counts = birim.read_budget_counts(self.data.root, self.data.config, self.data.snapshots, self.data.closing)
        layout = [(n, "Bölüm", [n]) for n in names]
        sheet = {
            row[0]: row
            for row in analiz.volume_price_rows(
                self.data.config, self.data.snapshots, self.data.closing, months, counts, layout
            )
        }
        for name in names:
            effects = [kalem._effects(self.data, name, months, code) for code in self.data.catalog]
            volume = sum((e["volume"] for e in effects), D(0))
            price = sum((e["price"] for e in effects), D(0))
            self.assertLess(abs(volume - sheet[name][9]), KURUS, name)
            self.assertLess(abs(price - sheet[name][10]), KURUS, name)

    def test_source_cells_reconcile_and_include_counts(self):
        r = kalem.detail(self.data, self.state, 3, department="Giriş")
        self.assertEqual(r["scope"]["departments"], ["Giriş"])
        self.assertTrue(all(x["ok"] for x in r["reconciliation"]))
        kinds = {row[4] for row in r["sources"]}
        self.assertEqual(
            kinds, {"Gerçekleşen", "Bütçe", "İşlem adedi (gerçekleşen)", "İşlem adedi (bütçe)"}
        )
        # Her ay için bir gerçekleşen, bir bütçe ve iki adet hücresi (eşleştirmesiz kalem).
        self.assertEqual(len(r["sources"]), 10 * 4)
        self.assertTrue(all(row[8] for row in r["sources"]))

    def test_multi_year_keeps_actual_only(self):
        state = {"filters": {"year": [2025, 2026]}}
        r = kalem.detail(self.data, state, 1)
        self.assertFalse(r["budgetReady"])
        self.assertIsNotNone(r["actual"])
        self.assertIsNone(r["budget"])
        self.assertIsNone(r["volume"])
        self.assertEqual([x["kind"] for x in r["reconciliation"]], ["Gerçekleşen"])
        self.assertTrue(r["reconciliation"][0]["ok"])

    def test_missing_planned_count_leaves_effects_empty(self):
        name = self.data.departments[0]
        key = (name, 2026, 3)
        saved = self.data.records[key]["count_budget"]
        try:
            self.data.records[key]["count_budget"] = D(0)
            e = kalem._effects(self.data, name, list(range(1, 11)), 1)
            self.assertIsNone(e["volume"])
            self.assertIn("Bütçe işlem adedi", e["reason"])
            self.assertIsNotNone(e["variance"])
        finally:
            self.data.records[key]["count_budget"] = saved

    def test_invalid_requests(self):
        with self.assertRaises(ValueError):
            kalem.detail(self.data, self.state, 999)
        with self.assertRaises(ValueError):
            kalem.detail(self.data, self.state, "1")
        with self.assertRaises(ValueError):
            kalem.detail(self.data, {"filters": {"department": ["Giriş"]}}, 1, department="Tartım")


class LineDetailHTTPTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.folder = tempfile.TemporaryDirectory()
        cls.app = Application(SETTINGS, (2026, 10), cls.folder.name)
        cls.server = Server(("127.0.0.1", 0), cls.app)
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.app.executor.shutdown()
        cls.folder.cleanup()

    def post(self, path, args):
        req = urllib.request.Request(
            self.server.origin + path,
            data=json.dumps(args).encode(),
            headers={"Cookie": "butce_session=" + self.app.token, "Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.load(r)

    def test_note_flows_into_line_detail(self):
        state = {"filters": {}}
        self.post("/api/note", {"state": state, "department": "Giriş", "code": 1, "note": "Ücret zammı."})
        r = self.post("/api/line-detail", {"state": state, "code": 1, "department": "Giriş"})
        self.assertEqual(r["departments"][0]["note"], "Ücret zammı.")
        self.assertAlmostEqual(r["volume"] + r["price"], r["variance"], places=2)


if __name__ == "__main__":
    unittest.main()


class LineTraceBasisTests(unittest.TestCase):
    """Hesap izi seçili bütçeyi ve geldiği dosyayı/dönemi söyler; ilk bütçede kaynak hücreler ilk kapanış
    dosyasından okunur ve hesaplanan bütçeyle tutar (dış inceleme, Kasım 2026)."""

    def test_trace_current_basis(self):
        r = kalem.detail(Dataset(SETTINGS, (2026, 10)), {"filters": {}}, 3, department="Giriş")
        self.assertEqual(r["basis"], "Güncel bütçe")
        steps = {t["step"]: t for t in r["trace"]}
        self.assertEqual(list(steps), ["Gerçekleşen", "Seçili bütçe · Güncel bütçe", "Fark", "Fark %"])
        self.assertIn("Ekim 2026", steps["Seçili bütçe · Güncel bütçe"]["source"])
        self.assertIn("kapanış 2026-10", steps["Seçili bütçe · Güncel bütçe"]["source"])
        self.assertEqual(steps["Fark"]["value"], r["actual"] - r["budget"])
        self.assertEqual(steps["Fark %"]["value"], r["rate"])
        self.assertTrue(all(x["ok"] for x in r["reconciliation"]))

    def test_first_basis_reads_budget_cells_from_first_closing_file(self):
        import shutil
        from openpyxl import load_workbook

        import butce

        with tempfile.TemporaryDirectory() as d:
            root = Path(d) / "ornek"
            shutil.copytree(BASE / "ornek", root)
            data0 = Dataset(root / "Ayarlar.xlsx", (2026, 10))
            cfg = data0.config["maliyet"]
            path = root / data0.config["maliyet_klasoru"] / "Eylül 2026 - Dosya 1.xlsx"
            wb = load_workbook(path)
            ws = wb["Giriş"]
            col = butce.positioned_columns(ws, cfg, (2026, 9))[1][(2026, 1)]
            cell = ws.cell(cfg["ilk_gider_satiri"] + 2, col)  # kalem 3, Ocak bütçesi
            cell.value = float(cell.value) + 1_000_000
            wb.save(path)

            real = analiz.settings
            analiz.settings = lambda c: {**real(c), "karsilastirma_butcesi": "İlk bütçe"}
            try:
                first = Dataset(root / "Ayarlar.xlsx", (2026, 10))
                r1 = kalem.detail(first, {"filters": {}}, 3, department="Giriş")
            finally:
                analiz.settings = real
            r0 = kalem.detail(Dataset(root / "Ayarlar.xlsx", (2026, 10)), {"filters": {}}, 3, department="Giriş")

        self.assertEqual(r1["basis"], "İlk bütçe (yıl başı)")
        self.assertEqual(r1["budget"] - r0["budget"], D("1000000"))
        self.assertEqual(r1["actual"], r0["actual"])
        step = next(t for t in r1["trace"] if t["step"].startswith("Seçili bütçe"))
        self.assertIn("Eylül 2026 - Dosya 1.xlsx (kapanış 2026-09)", step["source"])
        budget_cells = [s for s in r1["sources"] if s[4] == "Bütçe"]
        self.assertTrue(budget_cells and all("Eylül 2026" in s[6] for s in budget_cells))
        actual_cells = [s for s in r1["sources"] if s[4] == "Gerçekleşen"]
        self.assertTrue(all("Ekim 2026" in s[6] for s in actual_cells))
        self.assertTrue(all(x["ok"] for x in r1["reconciliation"]), r1["reconciliation"])
