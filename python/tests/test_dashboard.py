"""Run: python -m unittest discover -s tests -v (bundled wheels via run_tests.py)."""

import unittest, copy, tempfile, threading, json, time, urllib.request, urllib.error
from pathlib import Path
from decimal import Decimal as D
from openpyxl import load_workbook
from model import Dataset
from notes import Notes
from app import Application, Server
import jobs, butce

BASE = Path(__file__).resolve().parents[1]


class CalculationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data = Dataset(BASE / "ornek/Ayarlar.xlsx", (2026, 10))

    def test_reference_totals(self):
        q = self.data.query({})
        self.assertEqual(q["general"]["cards"]["actual"], D("3449.38339031"))
        self.assertEqual(q["general"]["cards"]["budget"], D("3440.93898013"))
        self.assertEqual(q["general"]["cards"]["gap"], D("-8.44441018"))
        self.assertEqual(q["transactions"]["cards"]["actual"], D(69216031))
        self.assertEqual(q["transactions"]["cards"]["budget"], D(67920000))
        self.assertEqual(len(q["audit"]["revisions"]), 26)
        self.assertTrue(all(r[7] == "Doğrulandı" for r in q["audit"]["checks"]))

    def test_arbitrary_months_same_period(self):
        q = self.data.query({"filters": {"month": [1, 3, 5]}})
        self.assertEqual(q["general"]["cards"]["actual"], D("913.18612667"))
        self.assertEqual(q["general"]["cards"]["budget"], D("1094.63598961"))
        expected = sum(
            self.data.records[n, 2025, m]["actual"] for n in self.data.departments for m in [1, 3, 5]
        )
        self.assertEqual(self.data.aggregate(self.data.departments, [2025], [1, 3, 5]), expected)

    def test_multi_year_budget_gate(self):
        q = self.data.query({"filters": {"year": [2025, 2026]}})
        self.assertIsNone(q["general"]["cards"]["budget"])
        self.assertIsNone(q["transactions"]["cards"]["budget"])
        self.assertIsNone(q["detail"]["total"]["budget"])
        self.assertEqual(len(q["priority"]["rows"]), 0)
        active = next(v for v in q["departments"]["expense"][-1]["values"] if v["year"] == 2026)
        self.assertEqual(active["budget"], D("3440.93898013"))

    def test_historical_year_no_budget(self):
        q = self.data.query({"filters": {"year": [2025]}})
        self.assertIsNone(q["general"]["cards"]["budget"])
        self.assertEqual(q["general"]["cards"]["actual"], D("3004.13075488"))
        self.assertTrue(all("budget" not in v for r in q["unit"]["rows"] for v in r["values"]))

    def test_future_month_missing_is_not_zero(self):
        q = self.data.query({"filters": {"month": [11]}})
        self.assertIsNone(q["general"]["cards"]["actual"])
        self.assertIsNone(q["general"]["cards"]["budget"])
        self.assertTrue(q["messages"])
        self.assertGreater(q["priority"]["incomplete"], 0)

    def test_empty_selections(self):
        for key in ["department", "year", "month"]:
            q = self.data.query({"filters": {key: []}})
            self.assertIsNone(q["general"]["cards"]["actual"])
            self.assertEqual(q["unit"]["rate"], None)

    def test_staff_counts_do_not_enter_financials(self):
        d = copy.deepcopy(self.data)
        for rec in d.records.values():
            rec["staff"] = [D(10**15)] * 4
            rec["count"] = D(10**15)
        self.assertEqual(d.query({})["general"]["cards"]["actual"], D("3449.38339031"))

    def test_staff_budget_completion(self):
        for n in self.data.departments:
            for m in range(1, 13):
                r = self.data.records[n, 2026, m]
                self.assertEqual(r["staff_budget"][2], r["staff_budget"][1])
                self.assertEqual(r["staff_budget"][3], r["staff_budget"][1])

    def test_transaction_previous_year_values(self):
        q = self.data.query({"filters": {"year": [2025, 2026]}})
        a = next(r for r in q["transactions"]["rows"] if r["year"] == 2026)["values"][-1]
        b = next(r for r in q["transactions"]["rows"] if r["year"] == 2025)["values"][-1]
        self.assertNotEqual(a["actual"], b["actual"])
        self.assertEqual(a["annual"]["value"], a["actual"] - b["actual"])

    def test_detail_groups_and_hover(self):
        q = self.data.query({"filters": {"year": [2024, 2025, 2026]}})
        self.assertEqual(sum(g["actual"] for g in q["detail"]["groups"]), q["detail"]["total"]["actual"])
        for g in q["detail"]["groups"]:
            self.assertEqual(sum(i["actual"] for i in g["items"]), g["actual"])
            self.assertEqual([v["year"] for v in g["history"]], [2024, 2025, 2026])

    def test_unit_weighted_reference(self):
        q = self.data.query({})
        self.assertAlmostEqual(q["unit"]["rate"], D("-0.0152069853055713"), places=12)
        department = [r for r in q["unit"]["rows"] if r["kind"] == "department"]
        self.assertEqual(sum(r["values"][0]["weight"] for r in department), D(1))
        for manager in self.data.meta()["options"]["manager"]:
            subset = [r for r in department if self.data.config["bolumler"][r["name"]]["mudurluk"] == manager]
            self.assertAlmostEqual(sum(r["values"][0]["managerWeight"] for r in subset), D(1), places=20)

    def test_local_scopes_independent(self):
        q = self.data.query({"bmLocal": ["Giriş"], "opexLocal": ["Tartım"]})
        self.assertEqual(q["unit"]["eligible"], ["Giriş"])
        self.assertEqual(q["departments"]["opexEligible"], ["Tartım"])
        self.assertEqual(q["unit"]["rows"][0]["values"][0]["weight"], 1)
        self.assertEqual(q["general"]["cards"]["actual"], D("3449.38339031"))
        self.assertAlmostEqual(q["unit"]["rate"], D("-.00934420966969"), places=12)

    def test_missing_previous_cost_excludes_and_renormalizes(self):
        d = copy.deepcopy(self.data)
        for n in ["Tartım", "İade Kabul"]:
            d.records[n, 2025, 1]["actual"] = None
        q = d.query({})
        self.assertNotIn("Tartım", q["unit"]["eligible"])
        self.assertNotIn("Tartım", q["departments"]["opexEligible"])
        g = next(r for r in q["unit"]["rows"] if r["name"] == "Giriş")["values"][0]
        self.assertEqual(g["managerWeight"], 1)
        self.assertAlmostEqual(
            sum(r["values"][0]["weight"] for r in q["unit"]["rows"] if r["kind"] == "department"),
            D(1),
            places=20,
        )

    def test_zero_count_excludes_unit_only(self):
        d = copy.deepcopy(self.data)
        d.records["Giriş", 2026, 1]["count"] = D(0)
        q = d.query({})
        self.assertNotIn("Giriş", q["unit"]["eligible"])
        self.assertIn("Giriş", q["departments"]["opexEligible"])
        self.assertEqual(q["general"]["cards"]["actual"], D("3449.38339031"))

    def test_index_arithmetic_mean_of_selected_monthly_ratios(self):
        d = copy.deepcopy(self.data)
        alias = butce.key(d.config["bolumler"]["Giriş"]["endeks"])
        for m, factor in [(1, 2), (3, 3), (5, 4)]:
            d.indices[alias, (2025, m)] = D(100)
            d.indices[alias, (2026, m)] = D(100) * factor
        q = d.query({"filters": {"month": [1, 3, 5]}, "bmLocal": ["Giriş"]})
        row = q["unit"]["rows"][0]["values"][0]
        self.assertEqual(row["inflation"], 2)
        self.assertEqual(row["adjusted"], row["prior"] * 3)

    def test_revenue_is_used_once(self):
        q = self.data.query({})
        totalrow = q["departments"]["opex"][-1]["values"][0]
        revenue = sum(self.data.revenue[2026, m] for m in range(1, 11))
        self.assertEqual(totalrow["actual"], D("3449383390.31") / revenue)

    def test_priority_reference(self):
        p = self.data.query({})["priority"]
        self.assertEqual(len(p["rows"]), 234)
        self.assertEqual(p["gross"], D("230.81928351"))
        self.assertEqual(p["netDepartments"], 9)
        self.assertEqual(p["rows"][0]["department"], "Yerleştirme")
        self.assertEqual(p["rows"][0]["item"], "Ücret giderleri")

    def test_year_rollover_in_memory(self):
        d = copy.deepcopy(self.data)
        d.budget_year = 2027
        d.closing = (2027, 10)
        d.years = [y + 1 for y in d.years]
        d.records = {(n, y + 1, m): r for (n, y, m), r in d.records.items()}
        d.revenue = {(y + 1, m): v for (y, m), v in d.revenue.items()}
        d.revenue_budget = {(y + 1, m): v for (y, m), v in d.revenue_budget.items()}
        d.indices = {(n, (y + 1, m)): v for (n, (y, m)), v in d.indices.items()}

        def span(values):
            return {(y + 1, m): v for (y, m), v in values.items()}

        d.snapshots = {
            ((cy + 1, cm), n): (
                {c: (label, span(a), span(b)) for c, (label, a, b) in snap[0].items()},
                span(snap[1]),
                span(snap[2]),
                snap[3],
            )
            for ((cy, cm), n), snap in d.snapshots.items()
        }
        d.controls = {k: [span(r) for r in rows] for k, rows in d.controls.items()}
        q = d.query({})
        self.assertEqual(q["selection"]["years"], [2027])
        self.assertEqual(q["general"]["cards"]["budget"], D("3440.93898013"))
        self.assertAlmostEqual(q["unit"]["rate"], D("-.0152069853055713"), places=12)

    def test_invalid_filter_rejected(self):
        with self.assertRaises(ValueError):
            self.data.query({"filters": {"department": ["Bilinmeyen bölüm"]}})

    def test_control_mismatch_has_visible_warning(self):
        d = copy.deepcopy(self.data)
        key = next(k for k, rows in d.controls.items() if rows[0].get((2026, 1)) is not None)
        d.controls[key][0][2026, 1] += D(100)
        q = d.query({})
        self.assertTrue(any("doğrulanamadı" in m for m in q["messages"]))
        self.assertEqual(q["general"]["cards"]["actual"], D("3449.38339031"))


class PersistenceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data = (
            CalculationTests.data
            if hasattr(CalculationTests, "data")
            else Dataset(BASE / "ornek/Ayarlar.xlsx", (2026, 10))
        )

    def test_notes_exact_months_and_stale_amounts(self):
        with tempfile.TemporaryDirectory() as folder:
            notes = Notes(Path(folder) / "state.sqlite3")
            entry = {"department": "Giriş", "code": 1, "note": "Ücret farkı incelendi."}
            notes.save(self.data, [1, 3, 5], [entry])
            self.assertEqual(Notes(notes.path).list(self.data, [1, 3, 5])["Giriş", 1]["note"], entry["note"])
            self.assertEqual(notes.list(self.data, [1, 2, 3]), {})
            d = copy.deepcopy(self.data)
            r = d.records["Giriş", 2026, 1]["items"][1]
            d.records["Giriş", 2026, 1]["items"][1] = (r[0] + 1, r[1])
            self.assertTrue(notes.list(d, [1, 3, 5])["Giriş", 1]["stale"])

    def test_note_import_atomic_and_period_validation(self):
        paths = list((BASE / "ornek/aciklama").glob("*Mudurluk*.xlsx"))
        if not paths:
            paths = [p for p in (BASE / "ornek/aciklama").glob("*.xlsx") if "Direktorluk" not in p.name]
        with tempfile.TemporaryDirectory() as folder:
            notes = Notes(Path(folder) / "state.sqlite3")
            count = notes.import_files(self.data, list(range(1, 11)), paths)
            self.assertEqual(count, 24)
            with self.assertRaises(ValueError):
                notes.import_files(self.data, [2], paths)
            self.assertEqual(len(notes.list(self.data, list(range(1, 11)))), 24)

    def test_existing_report_pipeline_generates_five_files(self):
        with tempfile.TemporaryDirectory() as folder:
            result = jobs.report(
                self.data, 2026, [1, 3, 5], folder, {("Giriş", 1): {"note": "=Formül değil, açıklama metni"}}
            )
            self.assertEqual(len(result["files"]), 5)
            self.assertTrue(all(Path(p).parent == Path(result["folder"]) for p in result["files"]))
            b = load_workbook(result["files"][0], data_only=True)
            self.assertIn("Genel Özet", b.sheetnames)
            self.assertIn("Kontrol ve Revizyon", b.sheetnames)
            self.assertEqual(b["_Rapor Kimliği"]["E1"].value, "aylar:01,03,05")
            self.assertEqual(b["_Rapor Kimliği"]["D1"].value, "2026-10")
            self.assertIn("Alt Gider Açıklamaları", b.sheetnames)
            self.assertIn("Giriş: =Formül değil", b["Alt Gider Açıklamaları"]["D2"].value)
            b.close()
            for path in result["files"][1:]:
                b = load_workbook(path, data_only=False)
                _, members, columns, items = butce.explanation_layout(b, "Müdürlük")
                if "Giriş" in members:
                    key = next(item for item in items if item[0] == 1)
                    cell = b["Giderler"].cell(items[key], columns["Giriş"] + 4)
                    self.assertEqual(cell.data_type, "s")
                    self.assertTrue(cell.value.startswith("="))
                b.close()


class HTTPTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.folder = tempfile.TemporaryDirectory()
        cls.app = Application(BASE / "ornek/Ayarlar.xlsx", (2026, 10), cls.folder.name)
        cls.server = Server(("127.0.0.1", 0), cls.app)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.app.executor.shutdown()
        cls.folder.cleanup()

    def request(self, path, args=None, authorized=True, origin=None):
        headers = {"Cookie": "butce_session=" + self.app.token} if authorized else {}
        if args is not None:
            headers["Content-Type"] = "application/json"
        if origin:
            headers["Origin"] = origin
        req = urllib.request.Request(
            self.server.origin + path,
            data=json.dumps(args).encode() if args is not None else None,
            headers=headers,
        )
        return urllib.request.urlopen(req, timeout=15)

    def test_cookie_auth(self):
        with self.assertRaises(urllib.error.HTTPError) as error:
            self.request("/api/meta", authorized=False)
        self.assertEqual(error.exception.code, 403)
        with self.request("/api/meta") as r:
            self.assertEqual(json.load(r)["budgetYear"], 2026)

    def test_cross_origin_rejected(self):
        with self.assertRaises(urllib.error.HTTPError) as error:
            self.request("/api/query", {}, origin="https://example.invalid")
        self.assertEqual(error.exception.code, 403)

    def test_no_filesystem_traversal(self):
        with self.assertRaises(urllib.error.HTTPError) as error:
            self.request("/../ornek/Ayarlar.xlsx")
        self.assertEqual(error.exception.code, 404)

    def test_query_and_saved_filters(self):
        state = {"filters": {"month": [1, 3, 5]}, "bmLocal": ["Giriş"]}
        with self.request("/api/query", state) as r:
            q = json.load(r)
        self.assertEqual(q["general"]["cards"]["actual"], 913.18612667)
        with self.request("/api/save-selection", state):
            pass
        with self.request("/api/meta") as r:
            self.assertEqual(json.load(r)["savedFilters"], state)

    def test_failed_load_preserves_dataset(self):
        old = self.app.data
        self.app.start_job("load", {"settings": "__missing__.xlsx", "closing": [2027, 1]})
        limit = time.monotonic() + 10
        while self.app.job["state"] == "running" and time.monotonic() < limit:
            time.sleep(0.02)
        self.assertEqual(self.app.job["state"], "error")
        self.assertIs(self.app.data, old)


if __name__ == "__main__":
    unittest.main()
