"""Local workbench: evidence, preflight, mapping review and immutable run inputs."""

from pathlib import Path
from datetime import datetime
from decimal import Decimal
from difflib import SequenceMatcher
import json, hashlib, shutil, uuid, os
from openpyxl import load_workbook
import butce, configuration


def encode(v):
    if isinstance(v, (Decimal, Path)):
        return str(v)
    raise TypeError(type(v).__name__)


def write_json(path, value):
    path = Path(path)
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(value, ensure_ascii=False, indent=2, default=encode), encoding="utf-8")
    os.replace(temp, path)


def inventory(data):
    paths = [
        ("settings", data.settings),
        ("control", data.root / data.config["kontrol_dosyasi"]),
        ("index", data.root / data.config["endeks_dosyasi"]),
    ]
    for path in sorted((data.root / data.config["maliyet_klasoru"]).glob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        try:
            p = butce.filename_period(path)
        except ValueError:
            continue
        if butce.in_closing_scope(p, data.closing):
            paths.append(("cost", path))
    return [
        {
            "role": role,
            "path": str(path.resolve()),
            "name": path.name,
            "sha256": configuration.digest(path),
            "bytes": path.stat().st_size,
        }
        for role, path in paths
    ]


def mapping_diff(old, new):
    old = old or {}
    new = dict(new)
    added = sorted(set(new) - set(old))
    removed = sorted(set(old) - set(new))
    changed = [
        {"department": n, "before": old[n], "after": new[n]} for n in new if n in old and old[n] != new[n]
    ]
    suggestions = []
    for n in added:
        scores = sorted(
            ((SequenceMatcher(None, butce.key(n), butce.key(p)).ratio(), p) for p in removed), reverse=True
        )
        if scores and scores[0][0] >= 0.45:
            suggestions.append({"old": scores[0][1], "new": n, "similarity": round(scores[0][0], 2)})
    return {"added": added, "removed": removed, "changed": changed, "suggestions": suggestions}


def quality(data, state):
    result = data.query(state)
    issues = []
    for p, n, file, msg in data.warnings:
        path = Path(file) if file else data.settings
        if file and not path.is_absolute():
            path = data.root / data.config["maliyet_klasoru"] / path
        issues.append(
            {
                "kind": "Kaynak",
                "department": n,
                "period": p,
                "file": str(path),
                "sheet": n,
                "cell": "",
                "message": msg,
                "action": "Kaynaklar ve eşleştirmeleri inceleyin; ayar Excel’ini düzelterek yenileyin.",
            }
        )
    for row in result["audit"]["checks"]:
        if row[7] != "Doğrulandı":
            issues.append(
                {
                    "kind": "Dip toplam",
                    "department": row[0],
                    "period": row[3],
                    "file": str(data.root / data.config["kontrol_dosyasi"]),
                    "sheet": data.config["kontrol"]["sayfa"],
                    "cell": data.config["kontrol"].get("ciro_baslangici", ""),
                    "message": str(row[7]) + f" · Fark TL: {row[6]}",
                    "action": "Birleşen bölüm eşleştirmesini ve kaynak/kontrol tutarlarını karşılaştırın.",
                }
            )
    names, years, months = data.selection(state)
    exclusions = []
    for n in names:
        if n not in result["unit"]["eligible"]:
            exclusions.append(
                [
                    n,
                    "Birim maliyet",
                    "Yerel seçim dışında veya endeks/önceki yıl/güncel gider/pozitif işlem adedi eksik.",
                ]
            )
        if n not in result["departments"]["opexEligible"]:
            exclusions.append(
                [n, "OPEX / ciro", "Yerel seçim dışında veya önceki yıl gideri/pozitif şirket cirosu eksik."]
            )
    return {
        "issues": issues,
        "checks": result["audit"]["checks"],
        "exclusions": exclusions,
        "revisions": result["audit"]["revisions"],
        "ready": not issues,
        "scope": {"departments": names, "years": years, "months": months},
        "expectedDepartments": len(data.departments),
        "receivedDepartments": len(data.sources),
    }


def trace(data, state, department=None, code=None):
    if inventory(data) != data.fingerprint:
        raise ValueError("Kaynak dosyaları değişmiş; hesap izini açmadan önce kaynakları yenileyin.")
    names, years, months = data.selection(state)
    if department:
        if department not in names:
            raise ValueError("Bölüm seçim kapsamında değil.")
        names = [department]
    if code is not None and code not in data.catalog:
        raise ValueError("Geçersiz gider kodu.")
    rows = []
    books = {}
    missing = []
    try:
        for n in names:
            snap = data.snapshots.get((data.closing, n))
            if not snap:
                missing.append(n)
                continue
            path = (data.root / data.config["maliyet_klasoru"] / snap[3]).resolve()
            if path not in books:
                books[path] = load_workbook(path, data_only=True)
            ws = books[path][n]
            a, b = butce.positioned_columns(ws, data.config["maliyet"], data.closing)
            for raw in range(1, data.config["maliyet"].get("gider_kalem_sayisi", 27) + 1):
                target = data.config.get("gider_eslestirme", {}).get(raw, (raw, ""))[0]
                if code is not None and target != code:
                    continue
                r = data.config["maliyet"]["ilk_gider_satiri"] + raw - 1
                for year in years:
                    for m in months:
                        for kind, cols in [("Gerçekleşen", a), ("Bütçe", b)]:
                            if kind == "Bütçe" and (years != [data.budget_year] or year != data.budget_year):
                                continue
                            c = cols.get((year, m))
                            cell = ws.cell(r, c) if c else None
                            v = butce.number(cell.value) if cell else None
                            if kind == "Gerçekleşen" and (year, m) > data.closing:
                                v = None
                            rows.append(
                                [
                                    n,
                                    data.config["bolumler"][n]["mudurluk"],
                                    target,
                                    data.catalog.get(target, ("",))[0],
                                    kind,
                                    butce.label((year, m)),
                                    str(path),
                                    n,
                                    cell.coordinate if cell else "",
                                    v,
                                ]
                            )
    finally:
        for book in books.values():
            book.close()
    if inventory(data) != data.fingerprint:
        raise ValueError("Kaynaklar hesap izi okunurken değişti; yeniden yenileyin.")
    actual = data.aggregate(names, years, months, code=code)
    budget = (
        data.aggregate(names, years, months, "budget", code=code) if years == [data.budget_year] else None
    )
    weighted = {
        (row["name"], value["year"]): value
        for row in data.query(state)["unit"]["rows"]
        if row["kind"] == "department"
        for value in row["values"]
    }
    unit = []
    for n in names:
        for y in years:
            i = data._unit_info(n, y, months)
            factors = []
            alias = data.config["bolumler"][n].get("endeks", "")
            for m in months:
                curr = data.indices.get((butce.key(alias), (y, m)))
                prior = data.indices.get((butce.key(alias), (y - 1, m)))
                factors.append([m, curr, prior, butce.ratio(curr, prior)])
            unit.append(
                {
                    "department": n,
                    "year": y,
                    **i,
                    "indices": factors,
                    "directorWeight": weighted.get((n, y), {}).get("weight"),
                    "managerWeight": weighted.get((n, y), {}).get("managerWeight"),
                }
            )
    return {
        "rows": rows,
        "missingDepartments": missing,
        "actualTL": actual,
        "budgetTL": budget,
        "gapTL": butce.delta(budget, actual),
        "rate": butce.ratio(butce.delta(budget, actual), budget),
        "unit": unit,
        "formula": "Gerçekleşen = seçili bölüm/ayların gider satırları toplamı (ilk 5 adet satırı hariç). Sapma = bütçe − gerçekleşen. BM = gider / işlem adedi. Endeks = seçili ayların yıllık endeks oranlarının aritmetik ortalaması. Reel iyileşme = (endekslenmiş önceki BM − güncel BM) / endekslenmiş önceki BM. Müdürlük/direktörlük sonucu = uygun bölümlerin güncel gider ağırlığı × iyileşme toplamı.",
    }


class Workspace:
    def __init__(self, app):
        self.app = app
        self.folder = app.state_dir / "calismalar"
        self.folder.mkdir(exist_ok=True)

    def review(self):
        data = self.app.data
        key = "mapping:" + data.identity
        old = self.app.preference(key, {})
        diff = mapping_diff(old.get("mappings"), data.config["bolumler"])
        return {
            **diff,
            "baseline": old,
            "fingerprint": configuration.digest(data.settings),
            "settings": str(data.settings),
        }

    def accept(self, args):
        data = self.app.data
        if args.get("fingerprint") != configuration.digest(data.settings):
            raise ValueError("Ayar değişti; kaynakları yenileyip tekrar inceleyin.")
        # Comparing to the actually loaded config avoids approving edited-but-unloaded mappings.
        if butce.read_settings(data.settings)["bolumler"] != data.config["bolumler"]:
            raise ValueError("Eşleştirmeler değişmiş; önce kaynakları yenileyin.")
        key = "mapping:" + data.identity
        old = self.app.preference(key, {})
        diff = self.review()
        ids = dict(old.get("ids", {}))
        renames = args.get("renames", [])
        if not isinstance(renames, list):
            raise ValueError("İsim değişiklikleri liste olmalı.")
        seen = set()
        targets = set()
        for pair in renames:
            source, target = pair["old"], pair["new"]
            if (
                source not in diff["removed"]
                or target not in diff["added"]
                or source in seen
                or target in targets
            ):
                raise ValueError("İsim değişikliği yalnız tekil eski/yeni bölüm çiftleri için yapılabilir.")
            seen.add(source)
            targets.add(target)
            ids[target] = ids.get(source, str(uuid.uuid4()))
        with self.app.notes.connect() as c:
            for pair in renames:
                if c.execute(
                    "SELECT 1 FROM notes WHERE dataset=? AND department=?", (data.identity, pair["new"])
                ).fetchone():
                    raise ValueError("Yeni bölümde açıklama var; otomatik taşınmadı.")
                if c.execute(
                    "SELECT 1 FROM note_review WHERE dataset=? AND department=?", (data.identity, pair["new"])
                ).fetchone():
                    raise ValueError("Yeni bölümde yorum inceleme kaydı var; taşıma durduruldu.")
                c.execute(
                    "UPDATE note_review SET department=? WHERE dataset=? AND department=?",
                    (pair["new"], data.identity, pair["old"]),
                )
                c.execute(
                    "UPDATE notes SET department=? WHERE dataset=? AND department=?",
                    (pair["new"], data.identity, pair["old"]),
                )
            for n in data.departments:
                ids.setdefault(n, str(uuid.uuid4()))
            value = {
                "mappings": data.config["bolumler"],
                "ids": ids,
                "acceptedAt": datetime.now().isoformat(timespec="seconds"),
            }
            c.execute(
                "INSERT OR REPLACE INTO preferences VALUES (?,?)",
                (key, json.dumps(value, ensure_ascii=False)),
            )
        return {
            "ok": True,
            "message": "Eşleştirme incelemesi kaydedildi. Onaylı isim değişikliklerinde açıklamalar taşındı; tutar değişmişse eski açıklama güncel sayılmaz.",
        }

    def history(self):
        rows = []
        for p in sorted(self.folder.glob("*/manifest.json"), reverse=True):
            try:
                r = json.loads(p.read_text(encoding="utf-8"))
                rows.append({**r, "id": p.parent.name})
            except (ValueError, OSError):
                continue
        return rows

    def capture(self, data, kind, args, checkpoint=lambda *a: None, notes=None):
        files = inventory(data)
        run_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f") + "_" + uuid.uuid4().hex[:6]
        folder = self.folder / run_id
        folder.mkdir()
        inputs = folder / "inputs"
        inputs.mkdir()
        (inputs / "maliyet").mkdir()
        baseline = next(
            (r for r in self.history() if r.get("dataset") == data.identity and r["id"] != run_id), None
        )
        before = {r["path"]: r["sha256"] for r in baseline.get("sources", [])} if baseline else {}
        manifest = {
            "id": run_id,
            "dataset": data.identity,
            "kind": kind,
            "state": "running",
            "started": datetime.now().isoformat(timespec="seconds"),
            "closing": list(data.closing),
            "args": args,
            "sources": files,
            "mappings": data.config["bolumler"],
            "changedFiles": [r["path"] for r in files if before.get(r["path"]) != r["sha256"]],
            "removedFiles": sorted(set(before) - {r["path"] for r in files}),
            "comparison": "Bir önceki aynı ayar yolu çalışması",
            "notes": [{"department": n, "code": code, **value} for (n, code), value in (notes or {}).items()],
            "engine": {
                name: configuration.digest(Path(__file__).parent / name)
                for name in (
                    "model.py",
                    "butce.py",
                    "birim.py",
                    "personel.py",
                    "jobs.py",
                    "notes.py",
                    "note_review.py",
                    "analiz.py",
                )
            },
        }
        write_json(folder / "manifest.json", manifest)
        try:
            for i, row in enumerate(files):
                checkpoint(f'Kaynak kopyası {i+1}/{len(files)}: {row["name"]}')
                dest = inputs / ("maliyet/" + row["name"] if row["role"] == "cost" else row["role"] + ".xlsx")
                shutil.copy2(row["path"], dest)
                if (
                    configuration.digest(dest) != row["sha256"]
                    or configuration.digest(row["path"]) != row["sha256"]
                ):
                    raise ValueError("Kaynak kopyalanırken değişti: " + row["name"])
            book = load_workbook(inputs / "settings.xlsx")
            try:
                for row in book["Dosya ve Konum Ayarları"].iter_rows(min_row=5, max_col=3):
                    key = row[1].value
                    if row[0].value == "Dosyalar" and key in configuration.KEYS.values():
                        row[2].value = {
                            "maliyet_klasoru": "maliyet",
                            "kontrol_dosyasi": "control.xlsx",
                            "endeks_dosyasi": "index.xlsx",
                        }[key]
                book.save(inputs / "Ayarlar.xlsx")
            finally:
                book.close()
            manifest["normalizedSettingsHash"] = configuration.digest(inputs / "Ayarlar.xlsx")
            write_json(folder / "manifest.json", manifest)
            checkpoint("Girdi kopyaları doğrulandı")
            return run_id, inputs / "Ayarlar.xlsx"
        except Exception as e:
            self.finish(run_id, "cancelled" if isinstance(e, Cancelled) else "error", {"error": str(e)})
            raise

    def finish(self, run_id, state, result):
        path = self.folder / run_id / "manifest.json"
        value = json.loads(path.read_text(encoding="utf-8"))
        value.update(state=state, finished=datetime.now().isoformat(timespec="seconds"), result=result)
        write_json(path, value)
        if state == "done":
            self.prune(keep=run_id)

    def prune(self, keep=None):
        """Girdi arşivinde yalnız son N çalışmayı tutar (Rapor › arsiv_saklama_adedi, 0 = sınırsız).

        Kaynak kopyaları gizli mali veri içerir; sınırsız birikmesini önler. VBA: MakroMerkez.ArchivePrune.
        """
        import analiz

        limit = analiz.settings(self.app.data.config)["arsiv_saklama_adedi"]
        if limit <= 0:
            return []
        runs = sorted((p for p in self.folder.iterdir() if (p / "manifest.json").is_file()), reverse=True)
        removed = []
        for folder in runs[limit:]:
            if folder.name == keep:
                continue
            shutil.rmtree(folder, ignore_errors=True)
            removed.append(folder.name)
        return removed

    def replay_settings(self, run_id):
        if (
            not isinstance(run_id, str)
            or not run_id
            or Path(run_id).name != run_id
            or "/" in run_id
            or "\\" in run_id
        ):
            raise ValueError("Geçersiz çalışma kimliği.")
        folder = self.folder / run_id
        manifest = json.loads((folder / "manifest.json").read_text(encoding="utf-8"))
        for name, checksum in manifest.get("engine", {}).items():
            if (
                name
                not in (
                    "model.py",
                    "butce.py",
                    "birim.py",
                    "personel.py",
                    "jobs.py",
                    "notes.py",
                    "note_review.py",
                    "analiz.py",
                )
                or configuration.digest(Path(__file__).parent / name) != checksum
            ):
                raise ValueError(
                    "Hesaplama sürümü farklı. Eski girdi kopyaları korunuyor; eski sürümle yeniden üretin."
                )
        if configuration.digest(folder / "inputs/Ayarlar.xlsx") != manifest.get("normalizedSettingsHash"):
            raise ValueError("Saklanan ayar kopyası değişmiş.")
        expected = {r["name"] for r in manifest["sources"] if r["role"] == "cost"}
        if {p.name for p in (folder / "inputs/maliyet").glob("*.xlsx")} != expected:
            raise ValueError("Arşivin kaynak dosya listesi değişmiş.")
        for row in manifest["sources"]:
            path = (
                folder
                / "inputs"
                / ("maliyet/" + row["name"] if row["role"] == "cost" else row["role"] + ".xlsx")
            )
            if configuration.digest(path) != row["sha256"]:
                raise ValueError("Saklanan girdi kopyası değişmiş: " + row["name"])
        return folder / "inputs/Ayarlar.xlsx", manifest


class Cancelled(Exception):
    pass
