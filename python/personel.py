"""Personel kaynak okuma ve seçili-ay hesapları; gider toplamlarından bağımsız."""

from pathlib import Path
from openpyxl import load_workbook

STAFF_HEADERS = (
    [f"Personel {p} Gerçekleşen" for p in range(1, 5)]
    + [f"Personel {p} Bütçe" for p in range(1, 5)]
    + [f"Personel {p} Gerçekleşen Mevcut" for p in range(1, 5)]
    + [f"Personel {p} Bütçe Mevcut" for p in range(1, 5)]
    + ["Personel Bütçe Uyumsuz", "Personel 3 Tamamlama", "Personel 4 Tamamlama"]
)


def budget_values(values):
    values = list(values)
    mismatch = values[0] is not None and values[1] is not None and values[0] != values[1]
    reference = values[1] if values[1] is not None else values[0]
    filled = [0, 0]
    for p in (2, 3):
        if values[p] is None and not mismatch and reference is not None:
            values[p] = reference
            filled[p - 2] = 1
    return values, int(mismatch), filled


def add_staff(payload, config, snapshots, closing, root):
    from butce import positioned_columns, number

    books = {}
    lookup = {}
    try:
        for name in config["bolumler"]:
            snapshot = snapshots.get((closing, name))
            if not snapshot:
                continue
            path = Path(root) / config["maliyet_klasoru"] / snapshot[3]
            if path not in books:
                books[path] = load_workbook(path, data_only=True)
            sheet = books[path][name]
            actual, budget = positioned_columns(sheet, config["maliyet"], closing)
            first = config["maliyet"]["islem_adedi_satiri"] - 4
            for year in payload["years"]:
                for month in range(1, 13):
                    p = (year, month)
                    a = [
                        (
                            number(sheet.cell(first + i, actual[p]).value)
                            if p in actual and p <= closing
                            else None
                        )
                        for i in range(4)
                    ]
                    raw_b = [
                        sheet.cell(first + i, budget[p]).value if p in budget else None for i in range(4)
                    ]
                    b = [number(value) for value in raw_b]
                    if any(raw not in (None, "") and value is None for raw, value in zip(raw_b, b)):
                        raise ValueError(
                            f"{name} {year}-{month:02d}: Personel bütçesinde sayısal olmayan hücre var."
                        )
                    b, mismatch, filled = budget_values(b)
                    lookup[name, year, month] = [
                        *[float(v) if v is not None else None for v in a + b],
                        *[int(v is not None) for v in a + b],
                        mismatch,
                        *filled,
                    ]
    finally:
        for book in books.values():
            book.close()
    for row in payload["rows"]:
        row.extend(lookup.get((row[3], row[4], row[6]), [None] * 8 + [0] * 11))
    payload["headers"] = [
        "Müdürlük",
        "Müdürlük Altı",
        "Lokasyon",
        "Bölüm",
        "Yıl",
        "Ay",
        "Ay No",
        "Gerçekleşen TL",
        "Bütçe TL",
        "Gider Mevcut",
        "Bütçe Mevcut",
        "Görünür",
        "Kaynak",
        "Kontrol Doğru",
        "Kontrol Açıklaması",
        *STAFF_HEADERS,
    ]
    payload["staffModel"] = staff_model(payload)
    payload["staffInitial"] = [
        (
            sum(
                row[15 + p] for row in payload["rows"] if row[4] == closing[0] and row[6] in payload["months"]
            )
            if all(
                row[23 + p] for row in payload["rows"] if row[4] == closing[0] and row[6] in payload["months"]
            )
            else None
        )
        for p in range(4)
    ]
    return payload


def staff_model(payload):
    from openpyxl.utils import get_column_letter as letter

    n = len(payload["rows"]) + 1

    def src(c):
        return f"'Veri'!${c}$2:${c}${n}"

    vis = src("L")
    cells = {
        "A1": "Görünüm",
        "B1": "aylik",
        "A2": "Seçili ay sayısı",
        "B2": "=COUNT(G10:G21)",
        "A3": "Bütçe kullanılabilir",
        "B3": "=Hesap!B6",
        "A4": "Bütçe 1/2 uyumsuzluğu",
        "B4": f'=SUMIFS({src("AF")},{vis},1)',
        "A5": "Eksik personel değeri",
        "B5": f'=4*SUM({vis})-SUMIFS({src("X")},{vis},1)-SUMIFS({src("Y")},{vis},1)-SUMIFS({src("Z")},{vis},1)-SUMIFS({src("AA")},{vis},1)',
        "A6": "Tamamlanan bütçe hücresi",
        "B6": f'=SUMIFS({src("AG")},{vis},1)+SUMIFS({src("AH")},{vis},1)',
        "A7": "Kaynak yıl sayısı",
        "B7": len(payload["years"]),
        "A8": "Durum",
        "B8": '=IF(Hesap!B3=0,"Seçime uygun personel verisi yok.",IF(B4>0,"Personel 1 ve 2 bütçesi farklı; boş Personel 3/4 bütçeleri tamamlanmadı.",IF(B5>0,"Seçili kapsamda eksik personel gerçekleşeni var.",IF(NOT(B3),"Yalnız gerçekleşenler gösteriliyor; bütçe yılı tek başına seçili değil.",""))))',
        "A9": "Yıl",
        "B9": "Seçili kayıt",
        "C9": "Yıl sırası",
        "D9": "Ay no",
        "E9": "Ay",
        "F9": "Seçili kayıt",
        "G9": "Ay sırası",
    }
    for i, y in enumerate(payload["years"], 10):
        cells.update(
            {
                f"A{i}": y,
                f"B{i}": f'=SUMIFS({vis},{src("E")},A{i})',
                f"C{i}": f'=IF(B{i}>0,COUNT($C$9:C{i-1})+1,"")',
            }
        )
    for i, m in enumerate(range(1, 13), 10):
        from dashboard_excel import MONTHS

        cells.update(
            {
                f"D{i}": m,
                f"E{i}": MONTHS[m - 1],
                f"F{i}": f'=SUMIFS({vis},{src("G")},D{i})',
                f"G{i}": f'=IF(F{i}>0,COUNT($G$9:G{i-1})+1,"")',
            }
        )
    for p in range(4):
        start = 30 + 25 * p
        actual, budget = src(letter(16 + p)), src(letter(20 + p))
        aflag, bflag = src(letter(24 + p)), src(letter(28 + p))
        cells[f"A{start-1}"] = "Dönem"
        for yi, y in enumerate(payload["years"], 2):
            cells[f"{letter(yi)}{start-1}"] = str(y)
        bcol = letter(len(payload["years"]) + 2)
        cells[f"{bcol}{start-1}"] = "Bütçe"
        for i in range(12):
            r = start + i
            month = f"$Z{r}"
            cells[f"Z{r}"] = f"=IFERROR(INDEX($D$10:$D$21,MATCH({i+1},$G$10:$G$21,0)),0)"
            cells[f"A{r}"] = (
                f'=IF($B$1="toplam",IF({i}=0,"Seçili aylar",""),IFERROR(INDEX($E$10:$E$21,MATCH({i+1},$G$10:$G$21,0)),""))'
            )
            for yi, y in enumerate(payload["years"], 2):
                common = f'{vis},1,{src("E")},{y}'
                all_count = f'SUMIFS({vis},{src("E")},{y})'
                a_total = f"IF(AND({all_count}>0,SUMIFS({aflag},{common})={all_count}),SUMIFS({actual},{common}),NA())"
                part = f'{common},{src("G")},{month}'
                count = f'SUMIFS({vis},{src("E")},{y},{src("G")},{month})'
                a_month = f"IF(AND({count}>0,SUMIFS({aflag},{part})={count}),SUMIFS({actual},{part}),NA())"
                cells[f"{letter(yi)}{r}"] = f'=IF(A{r}="",NA(),IF($B$1="toplam",{a_total},{a_month}))'
            common = f"{vis},1"
            count = f"SUM({vis})"
            total = f"IF(AND({count}>0,SUMIFS({bflag},{common})={count}),SUMIFS({budget},{common}),NA())"
            part = f'{common},{src("G")},{month}'
            count = f'SUMIFS({vis},{src("G")},{month})'
            monthly = f"IF(AND({count}>0,SUMIFS({bflag},{part})={count}),SUMIFS({budget},{part}),NA())"
            cells[f"{bcol}{r}"] = f'=IF(OR(A{r}="",NOT($B$3)),NA(),IF($B$1="toplam",{total},{monthly}))'
    return cells
