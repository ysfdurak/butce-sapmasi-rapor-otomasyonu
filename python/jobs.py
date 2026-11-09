"""Existing report engine exposed without Excel automation or PowerShell."""

from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from datetime import datetime
import shutil
import butce
from note_review import report_text, PENDING, literal


def report(data, year, months, output, notes=None, checkpoint=None):
    checkpoint = checkpoint or (lambda *a: None)
    if not months:
        raise ValueError("En az bir ay seçin.")
    if year not in data.years:
        raise ValueError("Rapor yılı kaynakta yok.")
    if year == data.budget_year and max(months) > data.closing_month:
        raise ValueError("Kapanıştan sonraki ay raporlanamaz.")
    output = Path(output).expanduser().resolve()
    cost = (data.root / data.config["maliyet_klasoru"]).resolve()
    if output == cost or cost in output.parents:
        raise ValueError("Çıktıyı maliyet kaynak klasörünün dışında seçin.")
    output.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    final = output / ("Raporlar_" + stamp)
    with TemporaryDirectory(prefix=".butce_hazirlaniyor_", dir=output) as folder:
        stage = Path(folder) / "raporlar"
        stage.mkdir()
        checkpoint("Raporlar yazılıyor; bu adım sonunda durdurulabilir")
        actual = butce.run(
            SimpleNamespace(
                ayar=data.settings,
                donem=f"{year}-{max(months):02}",
                kapanis=butce.label(data.closing),
                aylar=",".join(str(m) for m in months),
                cikti=stage / f'Direktorluk_Genel_{year}_Aylar_{"-".join(f"{m:02}" for m in months)}.xlsx',
            )
        )
        checkpoint("Rapor dosyaları doğrulandı")
        if notes and year == data.budget_year:
            write_notes(actual, notes)
        checkpoint("Tamamlanan raporlar yayımlanıyor")
        shutil.move(str(stage), str(final))
    return {"folder": str(final), "files": [str(final / p.name) for p in actual]}


def write_notes(paths, notes):
    from openpyxl import load_workbook

    for path in paths:
        book = load_workbook(path)
        try:
            kind = book["_Rapor Kimliği"]["B1"].value
            _, members, columns, items = butce.explanation_layout(book, kind)
            if not any(notes.get((n, code), {}).get("note") for n in members for code, group, label in items):
                continue
            if kind == "Müdürlük":
                for name in members:
                    for (code, group, label), row in items.items():
                        record = notes.get((name, code), {})
                        value = report_text(record)
                        if value:
                            cell = book["Giderler"].cell(row, columns[name] + 4)
                            literal(cell, value)
                            if record.get("status") == "pending":
                                from copy import copy

                                font = copy(cell.font)
                                font.color = "C63842"
                                cell.font = font
            else:
                if "Alt Gider Açıklamaları" in book:
                    del book["Alt Gider Açıklamaları"]
                sheet = book.create_sheet("Alt Gider Açıklamaları")
                sheet.append(["Ana gider", "Alt gider kodu", "Alt gider", "Bölüm açıklamaları"])
                for code, group, label in items:
                    text = "\n".join(
                        f"{n}: {report_text(notes[n,code])}"
                        for n in members
                        if notes.get((n, code), {}).get("note")
                    )
                    if len(text) > 32767:
                        raise ValueError("Birleşik açıklama hücre sınırını aşıyor: " + label)
                    if text:
                        sheet.append([group, code, label, text])
                        literal(sheet.cell(sheet.max_row, 4), text)
                        if any(notes.get((n, code), {}).get("status") == "pending" for n in members):
                            from openpyxl.styles import Font

                            sheet.cell(sheet.max_row, 4).font = Font(color="C63842")
                sheet.column_dimensions["A"].width = 25
                sheet.column_dimensions["C"].width = 35
                sheet.column_dimensions["D"].width = 100
                from openpyxl.styles import Alignment

                for row in sheet:
                    for cell in row:
                        cell.alignment = Alignment(wrap_text=True, vertical="top")
            if "Açıklama Durumu" in book:
                del book["Açıklama Durumu"]
            audit = book.create_sheet("Açıklama Durumu")
            audit.append(
                [
                    "Bölüm",
                    "Gider kodu",
                    "Gider",
                    "Durum",
                    "Önceki kapanış",
                    "Önceki ay seçimi",
                    "Kaynak",
                    "Açıklama",
                ]
            )
            for name in members:
                for code, group, label in items:
                    record = notes.get((name, code), {})
                    if record.get("note"):
                        audit.append(
                            [
                                name,
                                code,
                                label,
                                (
                                    "Kontrol bekliyor"
                                    if record.get("status") == "pending"
                                    else "Onaylı / bu dönem"
                                ),
                                record.get("previousClosing", ""),
                                record.get("previousMonths", ""),
                                record.get("source", ""),
                                record["note"],
                            ]
                        )
                        for c in (1, 3, 4, 5, 6, 7, 8):
                            literal(audit.cell(audit.max_row, c), audit.cell(audit.max_row, c).value or "")
                        if record.get("status") == "pending":
                            from openpyxl.styles import Font

                            for cell in audit[audit.max_row]:
                                cell.font = Font(color="C63842")
            audit.column_dimensions["H"].width = 90
            audit.column_dimensions["G"].width = 45
            from openpyxl.styles import Alignment

            for row in audit:
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
            butce.save_report(book, Path(path))
        finally:
            book.close()


def merge(general, managers, output):
    output = Path(output).expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    path = output / ("Direktorluk_Aciklamali_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".xlsx")
    butce.merge_explanations(Path(general), [Path(p) for p in managers], path)
    from openpyxl import load_workbook
    from openpyxl.styles import Font

    b = load_workbook(path)
    try:
        if "Açıklama Durumu" in b:
            del b["Açıklama Durumu"]
        if "Alt Gider Açıklamaları" in b:
            for row in b["Alt Gider Açıklamaları"].iter_rows(min_row=2):
                if row[3].value and "[Kontrol bekliyor" in str(row[3].value):
                    row[3].font = Font(color="C63842")
        butce.save_report(b, path)
    finally:
        b.close()
    return {"folder": str(output), "files": [str(path)]}
