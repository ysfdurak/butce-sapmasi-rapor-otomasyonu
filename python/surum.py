"""Bütçe sürümleri, bütçe tabloları, bütçe değişimi, trend ve girdi kaydı (V5.4).

Power Query şablonu (tools/pq_sablon.py) ve VBA (MakroSurum modülü) ile aynı kurallar; bkz. docs/VERI_VE_HESAP.md:

* Güncel bütçe: kapanış dosyasındaki bütçe. Önceki ay bütçesi: aynı yılın bir önceki kapanış dosyası.
  İlk bütçe: bölümün kapanış yılında okunan ilk kapanış dosyasındaki bütçe (yıl içi revizyonlardan önceki hali).
* Karşılaştırma bütçesi "İlk bütçe" seçilirse raporların bütçesi ilk bütçe olur (bütçe işlem adedi ve bütçe cirosu
  güncel kalır).
* Bütçe değişimi: güncel − önceki ay bütçesi; |fark| > kontrol toleransı olan ay "değişti" sayılır.
* Trend: ay ay gerçekleşen − karşılaştırma bütçesi. Kalemde kapanış ayında biten ve her ay aylık eşiği (varsayılan
  önemlilik tutarı / 12) aşan aleyhte seri en az N ay (varsayılan 3) ise "üst üste aleyhte giden kalem"dir.
* Her toplam katıdır: bir değer boşsa toplam boş kalır.
"""

from datetime import datetime
from decimal import Decimal

ZERO = Decimal(0)
MONTH_NAMES = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım",
               "Aralık"]
FIRST = "İlk bütçe (yıl başı)"
CURRENT = "Güncel bütçe"


def _total(values):
    values = list(values)
    return sum(values, ZERO) if values and None not in values else None


def _delta(a, b):
    return None if a is None or b is None else a - b


def comparison(rules):
    """Ayarlar › Rapor › karsilastirma_butcesi: 'ilk…' ise ilk bütçe, değilse güncel bütçe."""
    value = str(rules.get("karsilastirma_butcesi") or "").strip().translate(str.maketrans("İI", "iı")).casefold()
    return FIRST if value.startswith("ilk") else CURRENT


def trend_rules(rules):
    threshold = rules.get("trend_aylik_esik")
    threshold = Decimal(str(threshold)) if threshold not in (None, "") else rules["onemlilik_tutari"] / 12
    least = rules.get("trend_en_az_ay")
    return threshold, int(least) if least not in (None, "") else 3


def periods_of(snapshots, name, year):
    return sorted(p for (p, n) in snapshots if n == name and p[0] == year)


def versions(config, snapshots, closing):
    """Bölüm → {"current", "first", "first_period", "previous"} anlık görüntüleri (yoksa None)."""
    year, month = closing
    previous = (year, month - 1) if month > 1 else None
    out = {}
    for name in config["bolumler"]:
        found = periods_of(snapshots, name, year)
        out[name] = {
            "current": snapshots.get((closing, name)),
            "first": snapshots.get((found[0], name)) if found else None,
            "first_period": found[0] if found else None,
            "previous": snapshots.get((previous, name)) if previous else None,
        }
    return out


def with_first_budget(config, snapshots, closing):
    """Kapanış anlık görüntülerinde bütçeyi ilk bütçeyle değiştirir (karşılaştırma bütçesi = ilk bütçe)."""
    out = dict(snapshots)
    for name, v in versions(config, snapshots, closing).items():
        current, first = v["current"], v["first"]
        if current is None:
            continue
        first_items = first[0] if first else {}
        items = {
            code: (item, actual, dict(first_items[code][2]) if code in first_items else {})
            for code, (item, actual, _) in current[0].items()
        }
        out[(closing, name)] = (items, *current[1:])
    return out


def _month(snap, slot, period):
    return None if snap is None else _total(record[slot].get(period) for record in snap[0].values())


def _group(values, names):
    parts = [values.get(n) for n in names]
    return None if not parts or any(p is None for p in parts) else _total(parts)


def _months_by_group(per_dept, names):
    return [_group({n: per_dept[n][i] for n in per_dept}, names) for i in range(12)]


def _year_total(values):
    return _total(values) if len(values) == 12 and None not in values else None


# --------------------------------------------------------------------------- Satır hesapları


def version_rows(config, snapshots, closing, layout, months, year):
    """Bütçe Sürümleri: [ad, tür, gerçekleşen, güncel, ilk, güncel−ilk, gerç−güncel, gerç−ilk, gerç−ilk / ilk]."""
    v = versions(config, snapshots, closing)
    periods = [(year, m) for m in months]
    actual = {n: _total(_month(x["current"], 1, p) for p in periods) for n, x in v.items()}
    current = {n: _total(_month(x["current"], 2, p) for p in periods) for n, x in v.items()}
    first = {n: _total(_month(x["first"], 2, p) for p in periods) for n, x in v.items()}
    rows = []
    for title, level, names in layout:
        a, c, f = _group(actual, names), _group(current, names), _group(first, names)
        gap = _delta(a, f)
        rows.append([title, level, a, c, f, _delta(c, f), _delta(a, c), gap,
                     None if gap is None or f in (None, ZERO) else gap / f])
    return rows


def budget_table_rows(config, snapshots, closing, layout):
    """Bütçe tabloları: seviye → [[ad, üst seviye, 12 ay güncel bütçe, yıl, ilk yıl, güncel − ilk]]."""
    year = closing[0]
    v = versions(config, snapshots, closing)
    cur = {n: [_month(x["current"], 2, (year, m)) for m in range(1, 13)] for n, x in v.items()}
    first = {n: _year_total([_month(x["first"], 2, (year, m)) for m in range(1, 13)]) for n, x in v.items()}
    departments = config["bolumler"]
    tables = {"Bölüm": [], "Müdürlük": [], "Direktörlük": []}
    for title, level, names in layout:
        months = _months_by_group(cur, names)
        total = _year_total(months)
        initial = _group(first, names)
        row = [title] + months + [total, initial, _delta(total, initial)]
        if level == "Bölüm":
            tables["Bölüm"].append(row[:1] + [departments[title]["mudurluk"]] + row[1:])
        elif level == "Müdürlük toplamı":
            tables["Müdürlük"].append(row[:1] + [departments[names[0]]["direktorluk"]] + row[1:])
        else:
            tables["Direktörlük"].append(row)
    return tables


def change_rows(config, snapshots, closing, layout, detail_scope, tolerance):
    """Bütçe değişimi: (seviye satırları, kalem satırları). Ocak kapanışında veya önceki dosya yoksa boş."""
    year, month = closing
    if month == 1:
        return [], []
    tol = Decimal(str(tolerance))
    v = versions(config, snapshots, closing)
    cur = {n: [_month(x["current"], 2, (year, m)) for m in range(1, 13)] for n, x in v.items()}
    prev = {n: [_month(x["previous"], 2, (year, m)) for m in range(1, 13)] for n, x in v.items()}
    level_rows = []
    for title, level, names in layout:
        c, p = _months_by_group(cur, names), _months_by_group(prev, names)
        diff = [_delta(a, b) for a, b in zip(c, p)]
        if not any(d is not None and abs(d) > tol for d in diff):
            continue
        numeric = [d for d in diff if d is not None]
        level_rows.append([title, level] + diff + [sum(numeric, ZERO) if numeric else None, _year_total(p), _year_total(c)])
    item_rows = []
    departments = config["bolumler"]
    for name in detail_scope:
        current, previous = v[name]["current"], v[name]["previous"]
        if current is None or previous is None:
            continue
        for code, (item, _, budget) in sorted(current[0].items()):
            if code not in previous[0]:
                continue
            old = previous[0][code][2]
            g = [budget.get((year, m)) for m in range(1, 13)]
            o = [old.get((year, m)) for m in range(1, 13)]
            diff = [_delta(a, b) for a, b in zip(g, o)]
            changed = sum(1 for d in diff if d is not None and abs(d) > tol)
            if not changed:
                continue
            gy, py = _year_total(g), _year_total(o)
            item_rows.append([departments[name]["mudurluk"], name, code, item, changed, py, gy, _delta(gy, py)] + diff)
    return level_rows, item_rows


def trend_rows(config, snapshots, closing, layout, names_in_scope, year, rules):
    """Trend: (seviye satırları, kapsam toplamı, kalem serileri). Bütçe yılı dışında boş."""
    if year != closing[0]:
        return [], [None] * 12, []
    last = closing[1]
    threshold, least = trend_rules(rules)
    per = {}
    for name in config["bolumler"]:
        snap = snapshots.get((closing, name))
        per[name] = [
            _delta(_month(snap, 1, (year, m)), _month(snap, 2, (year, m))) if m <= last and snap else None
            for m in range(1, 13)
        ]
    level_rows = [[title, level] + _months_by_group(per, names) for title, level, names in layout]
    scope_total = _months_by_group(per, list(names_in_scope)) if names_in_scope else [None] * 12
    scope_total = [v if m < last else None for m, v in enumerate(scope_total)]
    departments = config["bolumler"]
    items = []
    for name in names_in_scope:
        snap = snapshots.get((closing, name))
        if snap is None:
            continue
        for code, (item, actual, budget) in sorted(snap[0].items()):
            g = [actual.get((year, m)) for m in range(1, 13)]
            b = [budget.get((year, m)) for m in range(1, 13)]
            bad = [g[i] is not None and b[i] is not None and g[i] - b[i] >= threshold for i in range(12)]
            streak = 0
            for i in range(last - 1, -1, -1):
                if not bad[i]:
                    break
                streak += 1
            if streak < least:
                continue
            series = sum((g[i] - b[i] for i in range(last - streak, last)), ZERO)
            items.append([departments[name]["mudurluk"], name, code, item, streak, sum(bad[:last]), series,
                          _delta(g[last - 1], b[last - 1])])
    items.sort(key=lambda r: (-r[4], -r[6]))
    return level_rows, scope_total, items


def input_rows(settings_path, config, closing):
    """Girdi kaydı: [tür, dosya, dönem, boyut KB, değiştirilme, durum]."""
    import butce

    root = settings_path.parent

    def info(path, kind, status=None, period=None):
        if not path.is_file():
            return [kind, path.name, None, None, None, "Bulunamadı"]
        stat = path.stat()
        return [kind, path.name, butce.label(period) if period else None, round(stat.st_size / 1024, 1),
                datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"), status or "Okundu"]

    rows = [info(settings_path, "Ayarlar"), info(root / config["kontrol_dosyasi"], "Kontrol"),
            info(root / config["endeks_dosyasi"], "Endeks")]
    cost_dir = root / config["maliyet_klasoru"]
    for path in sorted(cost_dir.iterdir(), key=lambda p: p.name):
        if not path.is_file() or path.suffix.lower() != ".xlsx" or path.name.startswith("~$"):
            continue
        try:
            p = butce.filename_period(path)
        except ValueError:
            rows.append(info(path, "Maliyet", "Dönem bulunamadı"))
            continue
        rows.append(info(path, "Maliyet", "Okundu" if butce.in_closing_scope(p, closing) else "Kapsam dışı", p))
    return rows


def facts(item_rows, has_previous, closing, trend_items, rules, budget_year):
    """Yönetici Özeti'ne eklenen iki satır: bütçe değişimi ve üst üste aleyhte giden kalem."""
    if closing[1] == 1 or not has_previous:
        change = "Karşılaştırma yok (Ocak kapanışı veya önceki ay dosyası yok)"
    else:
        effect = sum((r[7] for r in item_rows if r[7] is not None), ZERO)
        change = f"{len(item_rows)} kalemde değişti, yıllık toplam etki {_money0(effect)} TL"
    threshold, least = trend_rules(rules)
    trend = (f"{len(trend_items)} kalem (≥ {least} ay, aylık ≥ {_money0(threshold)} TL)" if budget_year
             else "Bütçe yılı dışında")
    return change, trend


def analysis(settings_path, config, snapshots, compared, closing, layout, months, year, names, rules):
    """Bir rapor (genel veya müdürlük) için sayfa verileri. compared: karşılaştırma bütçeli anlık görüntüler."""
    budget_year = year == closing[0]
    names = [n for n in config["bolumler"] if n in names]
    changes = change_rows(config, snapshots, closing, layout, names, config["kontrol"]["tolerans"])
    trend = trend_rows(config, compared, closing, layout, names, year, rules)
    previous = (closing[0], closing[1] - 1) if closing[1] > 1 else None
    has_previous = previous is not None and any((previous, n) in snapshots for n in names)
    return {
        "closing": closing,
        "rules": rules,
        "mode": comparison(rules),
        "versions": version_rows(config, snapshots, closing, layout, months, year),
        "tables": budget_table_rows(config, snapshots, closing, layout),
        "changes": changes,
        "trend": trend,
        "inputs": input_rows(settings_path, config, closing),
        "facts": facts(changes[1], has_previous, closing, trend[2], rules, budget_year),
    }


def _money0(value):
    return f"{value:,.0f}".replace(",", ".")


# --------------------------------------------------------------------------- Excel sayfaları


def _style(sheet, header_row=1, widths=None):
    from openpyxl.styles import Alignment, Font, PatternFill

    for cell in sheet[header_row]:
        if cell.value is not None:
            cell.fill = PatternFill("solid", fgColor="E7EEF2")
            cell.font = Font(bold=True, color="263746")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
    for column, width in (widths or {}).items():
        sheet.column_dimensions[column].width = width


def _numbers(sheet, first_row, columns, number_format):
    from openpyxl.utils import get_column_letter

    for row in sheet.iter_rows(min_row=first_row):
        for cell in row:
            if get_column_letter(cell.column) in columns and isinstance(cell.value, (int, float, Decimal)):
                cell.number_format = number_format


def _letters(first, count):
    from openpyxl.utils import get_column_letter

    return {get_column_letter(c) for c in range(first, first + count)}


MONEY = "#,##0.00;-#,##0.00;0.00"
WHOLE = "#,##0;-#,##0;0"


def add_version_sheet(book, rows, budget_year):
    sheet = book.create_sheet("Bütçe Sürümleri")
    sheet.append(["Bölüm / müdürlük / direktörlük", "Satır türü", "Gerçekleşen (TL)", "Güncel bütçe (TL)", "İlk bütçe (TL)",
                  "Güncel − ilk bütçe (TL)", "Gerçekleşen − güncel bütçe (TL)", "Gerçekleşen − ilk bütçe (TL)",
                  "İlk bütçeye göre sapma (%)"])
    if not budget_year:
        sheet.append(["Seçili yıl bütçe yılı değil."])
    for row in rows if budget_year else []:
        sheet.append(row)
    _numbers(sheet, 2, _letters(3, 6), MONEY)
    _numbers(sheet, 2, {"I"}, "0.00%")
    sheet["K1"] = ("İlk bütçe: bölümün bu yıl okunan ilk kapanış dosyasındaki bütçe. Klasörde Ocak dosyası yoksa ilk "
                   "bulunan ay alınır (Girdi Kaydı).")
    _style(sheet, widths={"A": 34, "B": 20, "C": 20, "D": 20, "E": 20, "F": 20, "G": 22, "H": 22, "I": 14, "K": 60})
    sheet.freeze_panes = "A2"


def add_budget_tables(book, tables, closing):
    titles = {"Bölüm": "Müdürlük", "Müdürlük": "Direktörlük", "Direktörlük": None}
    for level, parent in titles.items():
        sheet = book.create_sheet(f"Bütçe - {level}")
        sheet["A1"] = f"{level} bütçeleri · {closing[0]} (kapanış {closing[0]}-{closing[1]:02d} dosyasındaki güncel bütçe)"
        sheet["A2"] = "Ay boşsa bütçe verisi eksiktir; yıl toplamı yalnız 12 ay doluysa hesaplanır."
        sheet.append([])
        sheet.append([level] + ([parent] if parent else []) + [m[:3] for m in MONTH_NAMES]
                     + ["Yıl toplamı", "İlk bütçe yıl toplamı", "Güncel − ilk (TL)"])
        for row in tables[level]:
            sheet.append(row)
        first = 3 if parent else 2
        _numbers(sheet, 5, _letters(first, 15), MONEY)
        _style(sheet, 4, {"A": 30, "B": 24 if parent else 15})
        sheet.freeze_panes = sheet.cell(5, first).coordinate


def add_change_sheet(book, changes, closing):
    level_rows, item_rows = changes
    sheet = book.create_sheet("Bütçe Değişimi")
    year, month = closing
    previous = f"{year}-{month - 1:02d}" if month > 1 else ""
    sheet["A1"] = f"Bütçe değişimi: {previous} → {year}-{month:02d} kapanış dosyası (güncel − önceki)"
    sheet["A2"] = ("Ocak kapanışında önceki ay farklı yılın bütçesidir; karşılaştırma yapılmaz." if month == 1 else
                   "Yalnız bütçesi değişen satırlar. Önceki ay dosyası olmayan bölüm karşılaştırılmaz.")
    months = [m[:3] for m in MONTH_NAMES]
    sheet.append([])
    header = ["Bölüm / müdürlük / direktörlük", "Satır türü"] + months + ["Toplam değişim (TL)", "Önceki yıl bütçesi",
                                                                         "Güncel yıl bütçesi"]
    items_header = ["Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Değişen ay", "Önceki yıl bütçesi (TL)",
                    "Güncel yıl bütçesi (TL)", "Değişim (TL)"] + months
    for i, title in enumerate(header, 1):
        sheet.cell(4, i, title)
    for i, title in enumerate(items_header, 21):
        sheet.cell(4, i, title)
    sheet.cell(3, 21, "Bütçesi değişen kalemler")
    empty = month > 1 and not level_rows
    for r, row in enumerate(level_rows or ([["Bütçe değişmedi"]] if empty else []), 5):
        for c, value in enumerate(row, 1):
            sheet.cell(r, c, value)
    for r, row in enumerate(item_rows or ([["Değişen kalem yok"]] if month > 1 and not item_rows else []), 5):
        for c, value in enumerate(row, 21):
            sheet.cell(r, c, value)
    _numbers(sheet, 5, _letters(3, 15) | _letters(26, 3), MONEY)
    _numbers(sheet, 5, _letters(29, 12), WHOLE)
    _style(sheet, 4, {"A": 32, "B": 18, "U": 16, "V": 16, "W": 9, "X": 26, "Y": 10, "Z": 18, "AA": 18, "AB": 16})
    sheet.freeze_panes = "C5"


def add_trend_sheet(book, trend, threshold, least, budget_year):
    from openpyxl.chart import LineChart, Reference

    level_rows, scope_total, items = trend
    sheet = book.create_sheet("Trend")
    sheet["A1"] = "Aylık sapma trendi (gerçekleşen − karşılaştırma bütçesi)"
    sheet["B3"], sheet["C3"] = "Aylık eşik (TL)", float(threshold)
    sheet["B4"], sheet["C4"] = "En az üst üste ay", least
    sheet["C3"].number_format = "#,##0"
    if not budget_year:
        sheet["A5"] = "Seçili yıl bütçe yılı değil; trend hesaplanmadı."
    months = [m[:3] for m in MONTH_NAMES]
    for i, title in enumerate(["Bölüm / müdürlük / direktörlük", "Satır türü"] + months, 1):
        sheet.cell(7, i, title)
    sheet.cell(8, 1, "Kapsam toplamı")
    sheet.cell(8, 2, "Toplam")
    for m, value in enumerate(scope_total):
        sheet.cell(8, 3 + m, value)
    for r, row in enumerate(level_rows, 9):
        for c, value in enumerate(row, 1):
            sheet.cell(r, c, value)
    _numbers(sheet, 8, _letters(3, 12), WHOLE)
    sheet.cell(18, 17, "Üst üste aleyhte giden kalemler")
    for i, title in enumerate(["Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Üst üste aleyhte ay",
                               "Toplam aleyhte ay", "Seri toplam sapma (TL)", "Kapanış ayı sapması (TL)"], 17):
        sheet.cell(19, i, title)
    rows = items or ([["Seri yok"]] if budget_year else [])
    for r, row in enumerate(rows, 20):
        for c, value in enumerate(row, 17):
            sheet.cell(r, c, value)
    _numbers(sheet, 20, {"W", "X"}, MONEY)
    _style(sheet, 7, {"A": 32, "B": 18, "Q": 18, "R": 16, "S": 9, "T": 26, "U": 12, "V": 12, "W": 18, "X": 18})
    _style(sheet, 19)
    chart = LineChart()
    chart.title = "Kapsam toplamı: aylık sapma (TL)"
    chart.height, chart.width = 7, 18
    chart.add_data(Reference(sheet, min_col=3, max_col=14, min_row=8), from_rows=True, titles_from_data=False)
    chart.set_categories(Reference(sheet, min_col=3, max_col=14, min_row=7))
    chart.legend = None
    sheet.add_chart(chart, "Q2")
    sheet.freeze_panes = "C8"


def add_input_sheet(book, rows):
    sheet = book.create_sheet("Girdi Kaydı")
    sheet["A1"] = "Raporun okuduğu dosyalar"
    sheet["A2"] = ("Kaynak dosya sonradan değişirse boyut veya değiştirilme zamanı farklı olur; aylık raporu ayrı saklamak, o "
                   "ayın girdileriyle birlikte arşivlenmesini sağlar.")
    sheet.append([])
    sheet.append(["Tür", "Dosya", "Dönem", "Boyut (KB)", "Değiştirilme zamanı", "Durum"])
    for row in rows:
        sheet.append(row)
    _numbers(sheet, 5, {"D"}, "#,##0.0")
    _style(sheet, 4, {"A": 12, "B": 40, "C": 10, "D": 12, "E": 20, "F": 18})
    sheet.freeze_panes = "A5"


def add_sheets(book, data, budget_year):
    """butce.write_report → analiz.add_sheets içinden çağrılır."""
    add_version_sheet(book, data["versions"], budget_year)
    add_budget_tables(book, data["tables"], data["closing"])
    add_change_sheet(book, data["changes"], data["closing"])
    threshold, least = trend_rules(data["rules"])
    add_trend_sheet(book, data["trend"], threshold, least, budget_year)
    add_input_sheet(book, data["inputs"])
