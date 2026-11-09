"""Önemli sapmalar, yıl sonu tahmini, hacim/birim maliyet ayrıştırması ve yönetici özeti.

Hesap kuralları VBA'daki MakroAnaliz modülüyle birebir aynıdır (bkz. docs/VERI_VE_HESAP.md):

* Sapma = gerçekleşen − bütçe. Pozitif sapma bütçe aşımıdır (Aleyhte), negatif sapma bütçe altıdır (Lehte).
* Önemli sapma: |sapma| ≥ önemlilik tutarı VE |sapma / bütçe| ≥ önemlilik oranı. Bütçe sıfırsa oran
  tanımsızdır ve yalnız tutar eşiği uygulanır.
* Yıl sonu tahmini (kapanış yılı): Ocak–kapanış ayı gerçekleşen + kalan ayların bütçesi. İkinci yöntem
  olarak aylık ortalama gidiş (Ocak–kapanış gerçekleşen / kapanış ayı × 12) gösterilir.
* Hacim etkisi = (gerçekleşen adet − bütçe adet) × bütçe birim maliyeti.
  Birim maliyet etkisi = gerçekleşen gider − bütçe birim maliyeti × gerçekleşen adet.
  İki etkinin toplamı gerçekleşen − bütçe farkına eşittir.
"""

from decimal import Decimal

ZERO = Decimal(0)

DEFAULTS = {
    "onemlilik_tutari": Decimal("1000000"),
    "onemlilik_orani": Decimal("0.10"),
    "ozet_kalem_sayisi": 5,
    "arsiv_saklama_adedi": 20,
}

LEVEL_TOTALS = {"Müdürlük toplamı", "Direktörlük toplamı", "Genel toplam"}


def settings(config):
    """Rapor ayarlarını varsayılanlarla birleştirir (Ayarlar.xlsx › Dosya ve Konum Ayarları › Rapor)."""
    values = dict(DEFAULTS)
    values.update(config.get("rapor") or {})
    values["onemlilik_tutari"] = Decimal(str(values["onemlilik_tutari"]))
    values["onemlilik_orani"] = Decimal(str(values["onemlilik_orani"]))
    values["ozet_kalem_sayisi"] = int(values["ozet_kalem_sayisi"])
    values["arsiv_saklama_adedi"] = int(values["arsiv_saklama_adedi"])
    return values


def _total(values):
    values = list(values)
    return sum(values, ZERO) if values and None not in values else None


def _delta(a, b):
    return None if a is None or b is None else a - b


def _ratio(a, b):
    return None if a is None or b is None or b == ZERO else a / b


def direction(variance):
    """Gerçekleşen − bütçe farkının finansal yönü."""
    if variance is None:
        return None
    if variance > ZERO:
        return "Aleyhte"
    if variance < ZERO:
        return "Lehte"
    return "Bütçede"


def threshold_text(rules):
    """'≥ 1.000.000 TL ve ≥ %10' biçiminde eşik metni (VBA: MakroAnaliz.ThresholdText)."""
    amount = f"{rules['onemlilik_tutari']:,.0f}".replace(",", ".")
    rate = f"{(rules['onemlilik_orani'] * 100).normalize():f}".replace(".", ",")
    return f"≥ {amount} TL ve ≥ %{rate}"


def is_material(variance, budget, rules):
    if variance is None or budget is None:
        return False
    if abs(variance) < rules["onemlilik_tutari"]:
        return False
    rate = _ratio(variance, budget)
    return rate is None or abs(rate) >= rules["onemlilik_orani"]


def material_rows(detail, rules):
    """detail: butce.make_report satırları. Dönüş: önem sırasına göre önemli sapmalar."""
    rows = []
    for name, manager, code, item, group, actual, _, planned, *_ in detail:
        variance = _delta(actual, planned)
        if not is_material(variance, planned, rules):
            continue
        rows.append(
            [
                manager,
                name,
                code,
                item,
                group,
                actual,
                planned,
                variance,
                _ratio(variance, planned),
                direction(variance),
            ]
        )
    rows.sort(key=lambda r: (-abs(r[7]), r[1], r[2]))
    return rows


def _department_month_total(snapshot, slot, period):
    return _total(record[slot].get(period) for record in snapshot[0].values())


def forecast_rows(config, snapshots, closing, layout):
    """Kapanış yılı için bölüm ve toplam satırlarında yıl sonu tahmini."""
    year, last = closing
    done = [(year, m) for m in range(1, last + 1)]
    rest = [(year, m) for m in range(last + 1, 13)]
    values = {}
    for name in config["bolumler"]:
        snap = snapshots.get((closing, name))
        if snap is None:
            values[name] = None
            continue
        ytd = _total(_department_month_total(snap, 1, p) for p in done)
        ytd_budget = _total(_department_month_total(snap, 2, p) for p in done)
        remaining = _total(_department_month_total(snap, 2, p) for p in rest) if rest else ZERO
        values[name] = (
            ytd,
            ytd_budget,
            _total([ytd_budget, remaining]) if ytd_budget is not None else None,
            remaining,
        )
    rows = []
    for title, level, names in layout:
        parts = [values.get(n) for n in names]
        if not parts or any(p is None for p in parts):
            ytd = ytd_budget = annual = remaining = None
        else:
            ytd, ytd_budget, annual, remaining = (_total(p[i] for p in parts) for i in range(4))
        plan = _total([ytd, remaining])
        pace = ytd * 12 / last if ytd is not None else None
        rows.append(
            [
                title,
                level,
                ytd,
                ytd_budget,
                _delta(ytd, ytd_budget),
                annual,
                plan,
                _delta(plan, annual),
                direction(_delta(plan, annual)),
                pace,
                _delta(pace, annual),
            ]
        )
    return rows


def volume_price_rows(config, snapshots, closing, months, budget_counts, layout):
    """Seçili aylarda bütçe farkını hacim ve birim maliyet etkisine ayırır."""
    year = closing[0]
    periods = [(year, m) for m in months]
    base = {}
    for name in config["bolumler"]:
        snap = snapshots.get((closing, name))
        if snap is None:
            base[name] = None
            continue
        actual = _total(_department_month_total(snap, 1, p) for p in periods)
        budget = _total(_department_month_total(snap, 2, p) for p in periods)
        count = _total(v if (v := snap[2].get(p)) is not None and v > ZERO else None for p in periods)
        planned = (budget_counts or {}).get(name, {})
        planned_count = _total(v if (v := planned.get(p)) is not None and v > ZERO else None for p in periods)
        unit_budget = _ratio(budget, planned_count)
        volume = price = None
        if None not in (actual, unit_budget, count):
            volume = (count - planned_count) * unit_budget
            price = actual - unit_budget * count
        base[name] = (actual, budget, count, planned_count, volume, price)
    rows = []
    for title, level, names in layout:
        parts = [base.get(n) for n in names]
        complete = parts and all(p is not None for p in parts)
        if level == "Bölüm":
            actual, budget, count, planned_count, volume, price = parts[0] if complete else (None,) * 6
            rows.append(
                [
                    title,
                    level,
                    actual,
                    budget,
                    _delta(actual, budget),
                    count,
                    planned_count,
                    _ratio(actual, count),
                    _ratio(budget, planned_count),
                    volume,
                    price,
                ]
            )
            continue
        actual = _total(p[0] for p in parts) if complete else None
        budget = _total(p[1] for p in parts) if complete else None
        volume = _total(p[4] for p in parts) if complete else None
        price = _total(p[5] for p in parts) if complete else None
        rows.append(
            [title, level, actual, budget, _delta(actual, budget), None, None, None, None, volume, price]
        )
    return rows


def summary(detail, material, forecast, checks, warnings, rules, missing=()):
    """Yönetici özeti için ana göstergeler. Kapanış kaynağı eksik bölüm varsa toplamlar boş kalır."""
    actual = _total(r[5] for r in detail) if detail and not missing else None
    planned = _total(r[7] for r in detail) if detail and not missing else None
    variance = _delta(actual, planned)
    top = forecast[-1] if forecast else None
    verified = sum(1 for c in checks if c[7] == "Doğrulandı")
    limit = rules["ozet_kalem_sayisi"]
    adverse = [r for r in material if r[9] == "Aleyhte"][:limit]
    favourable = [r for r in material if r[9] == "Lehte"][:limit]
    return {
        "actual": actual,
        "budget": planned,
        "variance": variance,
        "rate": _ratio(variance, planned),
        "direction": direction(variance),
        "forecast": top[6] if top else None,
        "annual": top[5] if top else None,
        "forecastGap": top[7] if top else None,
        "checks": (verified, len(checks)),
        "warnings": len(warnings),
        "material": (len(material), sum(1 for r in material if r[9] == "Aleyhte")),
        "adverse": adverse,
        "favourable": favourable,
    }


# --------------------------------------------------------------------------- Excel çıktıları

MONEY = "#,##0.00;-#,##0.00;0.00"
PERCENT = "0.00%"
COUNT = "#,##0"
RED = "BE303A"
GREEN = "1B754C"


def _paint_direction(cell):
    from openpyxl.styles import Font

    if cell.value == "Aleyhte":
        cell.font = Font(color=RED, bold=True)
    elif cell.value == "Lehte":
        cell.font = Font(color=GREEN, bold=True)


def _formats(sheet, formats, first_row=2):
    for col, fmt in formats.items():
        for r in range(first_row, sheet.max_row + 1):
            cell = sheet.cell(r, col)
            if isinstance(cell.value, (int, float, Decimal)):
                cell.number_format = fmt


def note_cells(book):
    """Müdürlük raporunda (bölüm, gider kodu) → Giderler sayfasındaki açıklama hücresi."""
    from openpyxl.utils import get_column_letter

    if "Giderler" not in book.sheetnames:
        return {}
    sheet = book["Giderler"]
    result = {}
    starts = [c for c in range(4, sheet.max_column + 1, 5) if sheet.cell(4, c).value]
    for row in range(6, sheet.max_row + 1):
        code = sheet.cell(row, 2).value
        if not isinstance(code, int):
            continue
        for col in starts:
            result[(sheet.cell(4, col).value, code)] = f"{get_column_letter(col + 4)}{row}"
    return result


def add_material_sheet(book, rows, rules, editable, budget_year):
    from openpyxl.styles import Font, PatternFill
    from butce import add_table

    headers = [
        "Müdürlük",
        "Bölüm",
        "Gider kodu",
        "Gider kalemi",
        "Gider grubu",
        "Gerçekleşen (TL)",
        "Bütçe (TL)",
        "Gerçekleşen − bütçe (TL)",
        "Sapma (%)",
        "Yön",
        "Açıklama durumu",
    ]
    body = [r + [None] for r in rows]
    if not budget_year:
        body = [["Seçili yıl bütçe yılı değil; önemlilik kontrolü uygulanmadı."] + [None] * 10]
    elif not rows:
        body = [["Önemlilik eşiğini aşan sapma yok."] + [None] * 10]
    add_table(book, "Önemli Sapmalar", headers, body, (9,))
    sheet = book["Önemli Sapmalar"]
    _formats(sheet, {3: "0"})
    if not rows or not budget_year:
        return sheet
    cells = note_cells(book) if editable else {}
    required = PatternFill("solid", fgColor="F8CBAD")
    for r, row in enumerate(rows, 2):
        _paint_direction(sheet.cell(r, 10))
        status = sheet.cell(r, 11)
        address = cells.get((row[1], row[2]))
        if address:
            status.value = f'=IF(LEN(TRIM(Giderler!{address}))>0,"Açıklandı","Açıklama gerekli")'
            book["Giderler"][address].fill = required
        else:
            status.value = "Müdürlük dosyalarında açıklanır"
    sheet.cell(
        len(rows) + 3,
        1,
        f"Önemlilik eşiği: {threshold_text(rules)} (tutar ve oran birlikte). "
        "Açıklama gerekli hücreler müdürlük dosyalarının Giderler sayfasında turuncudur.",
    )
    sheet.cell(len(rows) + 3, 1).font = Font(italic=True, color="5B6B75")
    return sheet


def add_forecast_sheet(book, rows, closing):
    from butce import MONTHS, add_table, style_subtotals

    span = f"Ocak–{MONTHS[closing[1] - 1]} {closing[0]}"
    headers = [
        "Bölüm / müdürlük / direktörlük",
        "Satır türü",
        f"Gerçekleşen\n{span} (TL)",
        f"Bütçe\n{span} (TL)",
        f"Gerçekleşen − bütçe\n{span} (TL)",
        f"Yıllık bütçe\n{closing[0]} (TL)",
        "Yıl sonu tahmini\ngerçekleşen + kalan bütçe (TL)",
        "Tahmin − yıllık bütçe (TL)",
        "Yön",
        "Yıl sonu tahmini\naylık ortalama gidiş (TL)",
        "Ortalama gidiş − yıllık bütçe (TL)",
    ]
    add_table(book, "Yıl Sonu Tahmini", headers, rows)
    sheet = book["Yıl Sonu Tahmini"]
    for r in range(2, sheet.max_row + 1):
        _paint_direction(sheet.cell(r, 9))
    style_subtotals(sheet, [(r, row[1]) for r, row in enumerate(rows, 2) if row[1] in LEVEL_TOTALS])
    return sheet


def add_volume_sheet(book, rows, caption):
    from butce import add_table, style_subtotals

    headers = [
        "Bölüm / müdürlük / direktörlük",
        "Satır türü",
        f"Gerçekleşen gider\n{caption} (TL)",
        f"Bütçe gider\n{caption} (TL)",
        "Gerçekleşen − bütçe (TL)",
        "Gerçekleşen işlem adedi",
        "Bütçe işlem adedi",
        "Gerçekleşen birim maliyet (TL/işlem)",
        "Bütçe birim maliyet (TL/işlem)",
        "Hacim etkisi (TL)",
        "Birim maliyet etkisi (TL)",
    ]
    add_table(book, "Hacim ve Birim Maliyet", headers, rows)
    sheet = book["Hacim ve Birim Maliyet"]
    _formats(sheet, {6: COUNT, 7: COUNT})
    style_subtotals(sheet, [(r, row[1]) for r, row in enumerate(rows, 2) if row[1] in LEVEL_TOTALS])
    note = sheet.cell(
        len(rows) + 3,
        1,
        (
            "Hacim etkisi = (gerçekleşen adet − bütçe adet) × bütçe birim maliyeti. Birim maliyet etkisi = "
            "gerçekleşen gider − bütçe birim maliyeti × gerçekleşen adet. Toplamları gerçekleşen − bütçe farkını verir."
        ),
    )
    from openpyxl.styles import Font

    note.font = Font(italic=True, color="5B6B75")
    return sheet


def add_executive_sheet(book, facts, rules, caption, closing, budget_year):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.properties import PageSetupProperties

    sheet = book.create_sheet("Yönetici Özeti", 0)
    sheet["A1"] = "Yönetici özeti"
    sheet["A1"].font = Font(size=16, bold=True, color="233C57")
    sheet["A2"] = f"{caption} · Kaynak kapanış {closing[0]:04d}-{closing[1]:02d}"
    sheet["A2"].font = Font(color="5B6B75")
    header_fill = PatternFill("solid", fgColor="E7EEF2")
    sheet.append([])
    sheet.append(["Gösterge", "Değer"])
    checks = facts["checks"]
    count, adverse = facts["material"]
    kpis = [
        ("Gerçekleşen (TL)", facts["actual"], MONEY),
        ("Bütçe (TL)", facts["budget"] if budget_year else None, MONEY),
        ("Gerçekleşen − bütçe (TL)", facts["variance"] if budget_year else None, MONEY),
        ("Sapma oranı (%)", facts["rate"] if budget_year else None, PERCENT),
        ("Yön", facts["direction"] if budget_year else "Bütçe yılı dışında", None),
        (
            "Yıl sonu tahmini: gerçekleşen + kalan bütçe (TL)",
            facts["forecast"] if budget_year else None,
            MONEY,
        ),
        ("Yıllık bütçe (TL)", facts["annual"] if budget_year else None, MONEY),
        ("Tahmin − yıllık bütçe (TL)", facts["forecastGap"] if budget_year else None, MONEY),
        ("Veri kontrolü", f"{checks[0]} / {checks[1]} kontrol doğrulandı", None),
        ("Uyarı sayısı", facts["warnings"], "0"),
        ("Önemli sapma", f"{count} kalem ({adverse} aleyhte)" if budget_year else "Uygulanmadı", None),
        ("Önemlilik eşiği", threshold_text(rules), None),
    ]
    for label, value, fmt in kpis:
        sheet.append([label, value])
        cell = sheet.cell(sheet.max_row, 2)
        if fmt and value is not None:
            cell.number_format = fmt
        if label == "Yön":
            _paint_direction(cell)
    for cell in sheet[4]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color="263746")
    columns = [
        "Müdürlük",
        "Bölüm",
        "Gider kodu",
        "Gider kalemi",
        "Gerçekleşen (TL)",
        "Bütçe (TL)",
        "Gerçekleşen − bütçe (TL)",
        "Sapma (%)",
    ]
    for title, rows in (
        ("En büyük aleyhte önemli sapmalar", facts["adverse"]),
        ("En büyük lehte önemli sapmalar", facts["favourable"]),
    ):
        sheet.append([])
        sheet.append([title])
        sheet.cell(sheet.max_row, 1).font = Font(bold=True, color="233C57")
        sheet.append(columns)
        for cell in sheet[sheet.max_row]:
            cell.fill = header_fill
            cell.font = Font(bold=True, color="263746")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        if not rows:
            sheet.append(["Yok" if budget_year else "Bütçe yılı dışında"])
        for row in rows:
            sheet.append([row[0], row[1], row[2], row[3], row[5], row[6], row[7], row[8]])
            for col in (5, 6, 7):
                sheet.cell(sheet.max_row, col).number_format = MONEY
            sheet.cell(sheet.max_row, 8).number_format = PERCENT
    widths = {"A": 46, "B": 30, "C": 11, "D": 32, "E": 20, "F": 20, "G": 22, "H": 12}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    sheet.sheet_view.showGridLines = False
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    sheet.print_area = f"A1:H{sheet.max_row}"
    book.active = 0
    for other in book.worksheets:
        other.sheet_view.tabSelected = other is sheet
    return sheet


def add_sheets(book, analysis, editable):
    """butce.write_report içinden çağrılır; mevcut sayfa düzenlerini değiştirmez."""
    rules = analysis["rules"]
    budget_year = analysis["budgetYear"]
    add_material_sheet(book, analysis["material"], rules, editable, budget_year)
    if budget_year:
        add_forecast_sheet(book, analysis["forecast"], analysis["closing"])
        add_volume_sheet(book, analysis["volume"], analysis["caption"])
    facts = summary(
        analysis["detail"],
        analysis["material"],
        analysis["forecast"],
        analysis["checks"],
        analysis["warnings"],
        rules,
        analysis.get("missing", ()),
    )
    add_executive_sheet(book, facts, rules, analysis["caption"], analysis["closing"], budget_year)
