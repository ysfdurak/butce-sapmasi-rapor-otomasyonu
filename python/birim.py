"""Bütçe işlem adetleri: kapanış dosyasındaki bütçe bloğundan, işlem adedi satırı."""

from pathlib import Path
from openpyxl import load_workbook


def read_budget_counts(root, config, snapshots, closing):
    from butce import key, positioned_columns, number

    books = {}
    result = {}
    try:
        for name in config["bolumler"]:
            snap = snapshots.get((closing, name))
            if not snap:
                continue
            path = Path(root) / config["maliyet_klasoru"] / snap[3]
            if path not in books:
                books[path] = load_workbook(path, data_only=True)
            # Kaynak sayfa adı ayardaki bölüm adıyla büyük/küçük harf veya Türkçe harf farkıyla eşleşebilir.
            sheet = next(ws for ws in books[path] if key(ws.title) == key(name))
            _, cols = positioned_columns(sheet, config["maliyet"], closing)
            result[name] = {
                p: number(sheet.cell(config["maliyet"]["islem_adedi_satiri"], c).value)
                for p, c in cols.items()
            }
    finally:
        for book in books.values():
            book.close()
    return result
