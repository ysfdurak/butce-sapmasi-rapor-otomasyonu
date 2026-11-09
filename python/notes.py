"""Local explanation storage, isolated by dataset, closing and exact selected months."""

import sqlite3
from datetime import datetime
from pathlib import Path
from decimal import Decimal
from openpyxl import load_workbook
import butce
from note_review import Review, PENDING, audit, AUDIT_TABLE


class Notes:
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as c:
            c.execute(
                "CREATE TABLE IF NOT EXISTS notes (dataset TEXT, closing TEXT, months TEXT, department TEXT, code INTEGER, note TEXT, actual TEXT, budget TEXT, source TEXT, updated TEXT, PRIMARY KEY(dataset,closing,months,department,code))"
            )
            c.execute(AUDIT_TABLE)
        self.review = Review(self)

    def connect(self):
        return sqlite3.connect(self.path, timeout=10)

    def key(self, data, months):
        return (data.identity, butce.label(data.closing), butce.month_signature(sorted(months)))

    def list(self, data, months):
        with self.connect() as c:
            rows = c.execute(
                "SELECT department,code,note,actual,budget,source,updated FROM notes WHERE dataset=? AND closing=? AND months=?",
                self.key(data, months),
            ).fetchall()
        result = {}
        for name, code, note, actual, budget, source, updated in rows:
            a = data.aggregate([name], [data.budget_year], months, "actual", code=code)
            b = data.aggregate([name], [data.budget_year], months, "budget", code=code)
            stale = a is None or b is None or Decimal(actual) != a or Decimal(budget) != b
            result[name, code] = {
                "note": note if not stale else "",
                "stale": stale,
                "source": source,
                "updated": updated,
            }
        for row in self.review.rows(data, months):
            key = row["department"], row["code"]
            if row["status"] == "approved" and not row["protected"] and key in result:
                result[key].update(
                    status="approved",
                    previousClosing=row["previousClosing"],
                    previousMonths=row["previousMonths"],
                )
        for key, value in self.review.effective(data, months).items():
            if key not in result or result[key].get("stale"):
                result[key] = value
        trail = self.trail(data, months)
        for key, value in result.items():
            last = trail.get(key)
            if last:
                value.update(lastAction=last["action"], actor=last["user"], actedAt=last["at"])
        return result

    def save(self, data, months, rows, source="Uygulamada girildi"):
        staged = []
        if not months:
            raise ValueError("Açıklama için en az bir ay seçin.")
        for row in rows:
            name = row["department"]
            code = row["code"]
            note = row["note"]
            if (
                name not in data.departments
                or code not in data.catalog
                or not isinstance(note, str)
                or len(note) > 32767
            ):
                raise ValueError("Açıklama bölümü, gider kodu veya metni geçersiz.")
            a = data.aggregate([name], [data.budget_year], months, "actual", code=code)
            b = data.aggregate([name], [data.budget_year], months, "budget", code=code)
            if a is None or b is None:
                raise ValueError("Tutarları eksik kayda açıklama eklenemez.")
            staged.append(
                (
                    *self.key(data, months),
                    name,
                    code,
                    note,
                    str(a),
                    str(b),
                    source,
                    datetime.now().isoformat(timespec="seconds"),
                )
            )
        with self.connect() as c:
            c.executemany("INSERT OR REPLACE INTO notes VALUES (?,?,?,?,?,?,?,?,?,?)", staged)
            for row in staged:
                audit(c, row[:5], "Kaydedildi: " + source, row[5])
        return len(staged)

    def trail(self, data, months):
        """(bölüm, gider kodu) → son işlem, işlemi yapan ve zamanı (onay izi)."""
        with self.connect() as c:
            rows = c.execute(
                "SELECT department,code,action,user,at FROM note_audit WHERE dataset=? AND closing=? AND months=?"
                " ORDER BY at,rowid",
                self.key(data, months),
            ).fetchall()
        return {
            (name, code): {"action": action, "user": user, "at": at} for name, code, action, user, at in rows
        }

    def import_files(self, data, months, paths):
        staged = {}
        sources = []
        for path in paths:
            book = load_workbook(path, data_only=False)
            try:
                if "_Rapor Kimliği" not in book:
                    raise ValueError("Rapor kimliği yok: " + Path(path).name)
                kind = book["_Rapor Kimliği"]["B1"].value
                identity, members, columns, items = butce.explanation_layout(book, kind)
                if str(identity[0])[:4] != str(data.budget_year) or identity[1:] != (
                    butce.label(data.closing),
                    butce.month_signature(months),
                ):
                    raise ValueError("Açıklamanın yılı, kapanışı veya seçili ayları farklı.")
                if kind not in ("Müdürlük", "Direktörlük"):
                    raise ValueError("Müdürlük veya birleşik direktörlük raporu seçin.")
                merged = {}
                if kind == "Direktörlük":
                    if "Alt Gider Açıklamaları" not in book:
                        raise ValueError("Birleşik açıklama sayfası bulunamadı.")
                    for vals in book["Alt Gider Açıklamaları"].iter_rows(min_row=2, max_col=4):
                        code = vals[1].value
                        if not isinstance(code, int):
                            continue
                        if vals[3].data_type == "f":
                            raise ValueError("Açıklama formül olamaz.")
                        current = None
                        parts = {}
                        for line in str(vals[3].value or "").splitlines():
                            match = next((name for name in members if line.startswith(name + ":")), None)
                            if match:
                                current = match
                                parts.setdefault(match, []).append(line[len(match) + 1 :].lstrip())
                            elif current:
                                parts[current].append(line)
                        for name, lines in parts.items():
                            merged[name, code] = "\n".join(lines)
                sheet = book["Giderler"]
                for name in members:
                    if name not in data.departments:
                        raise ValueError("Eşleşmeyen bölüm: " + name)
                    meta = data.config["bolumler"][name]
                    if members[name] != (meta["mudurluk"], meta["direktorluk"]):
                        raise ValueError("Rapor organizasyon eşleştirmesi farklı: " + name)
                    col = columns[name]
                    for (code, group, label), r in items.items():
                        if code not in data.catalog or data.catalog[code] != (label, group):
                            raise ValueError("Alt gider adı/grubu farklı: " + str(code))
                        for offset, key in [(0, "actual"), (1, "budget")]:
                            expected = data.aggregate([name], [data.budget_year], months, key, code=code)
                            cell = sheet.cell(r, col + offset)
                            actual = butce.number(cell.value)
                            # Excel exports may contain harmless binary tails (e.g. 0.00000001 TL).
                            tolerance = Decimal(str(data.config["kontrol"]["tolerans"]))
                            if (
                                cell.data_type == "f"
                                or expected is None
                                or actual is None
                                or abs(actual - expected) > tolerance
                            ):
                                raise ValueError(f"Rapor tutarı farklı/eksik: {name} / {label}")
                        if kind == "Müdürlük":
                            cell = sheet.cell(r, col + 4)
                            if cell.data_type == "f":
                                raise ValueError("Açıklama formül olamaz.")
                            note = str(cell.value or "").strip()
                        else:
                            note = merged.get((name, code), "").strip()
                        if note.startswith("[Kontrol bekliyor"):
                            raise ValueError(
                                "Kontrol bekleyen yorum onaylı açıklama olarak yüklenemez; açıklama incelemesini kullanın."
                            )
                        if note:
                            record = {"department": name, "code": code, "note": note}
                            if (name, code) in staged and staged[name, code] != record:
                                raise ValueError("Aynı bölüm/gidere farklı açıklamalar verildi.")
                            staged[name, code] = record
                sources.append(Path(path).name)
            finally:
                book.close()
        return self.save(data, months, list(staged.values()), "; ".join(sources))
