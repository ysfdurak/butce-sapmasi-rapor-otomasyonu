import argparse
import json
import re
import sys
from tempfile import NamedTemporaryFile
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.utils import get_column_letter, range_boundaries

MONTHS = "Ocak Şubat Mart Nisan Mayıs Haziran Temmuz Ağustos Eylül Ekim Kasım Aralık".split()
ZERO = Decimal(0)
ONE = Decimal(1)


def key(value):
    return str(value or "").strip().translate(str.maketrans("İI", "iı")).casefold()


def number(value):
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        return None
    amount = Decimal(str(value))
    return amount if amount.is_finite() else None


TR_NUMBER = re.compile(r"[-+]?\d{1,3}(?:\.\d{3})+(?:,\d+)?|[-+]?\d+(?:,\d+)?")
DOT_DECIMAL = re.compile(r"[-+]?\d+\.\d+")


def text_number(value):
    """Metin olarak girilmiş tutarı çevirir: '1.234,56', '1234,56', '(1.234)', '1234.56'. Değilse None.

    Nokta + üçlü basamak grubu Türkçe binlik ayırıcı sayılır ('1.234' = 1234). VBA: MakroOrtak.TextNumber.
    """
    if not isinstance(value, str):
        return None
    text = value.strip().replace(" ", "").replace(" ", "").replace("₺", "")
    if text.upper().endswith("TL"):
        text = text[:-2]
    negative = len(text) > 2 and text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    if TR_NUMBER.fullmatch(text):
        text = text.replace(".", "").replace(",", ".")
    elif not DOT_DECIMAL.fullmatch(text):
        return None
    amount = Decimal(text)
    return -amount if negative else amount


def sample_cells(cells):
    """İlk beş hücre adresi; fazlası '…' ile belirtilir (VBA: MakroVeri.SampleCells)."""
    return ", ".join(cells[:5]) + (" …" if len(cells) > 5 else "")


class SourceIssues:
    """Kaynak okurken metin sayıları ve boş hücreleri hücre adresiyle toplar."""

    def __init__(self):
        self.converted, self.unreadable, self.empty = [], [], []

    def value(self, sheet, row, col):
        raw = sheet.cell(row, col).value
        amount = number(raw)
        if amount is not None:
            return amount
        address = f"{get_column_letter(col)}{row}"
        if isinstance(raw, str) and raw.strip():
            amount = text_number(raw)
            (self.converted if amount is not None else self.unreadable).append(address)
            return amount
        if raw is None:
            self.empty.append(address)
        return None

    def warnings(self, sheet_name):
        sample = sample_cells
        result = []
        if self.converted:
            result.append(
                f"Metin olarak girilmiş {len(self.converted)} tutar sayıya çevrildi ({sheet_name}: "
                f"{sample(self.converted)}). Kaynakta sayı biçimine çevirin."
            )
        if self.unreadable:
            result.append(
                f"Sayıya çevrilemeyen {len(self.unreadable)} hücre boş sayıldı ({sheet_name}: "
                f"{sample(self.unreadable)})."
            )
        return result


def month(value):
    if isinstance(value, int) and 1 <= value <= 12:
        return value
    for i, name in enumerate(MONTHS, 1):
        if key(value) in {key(name), key(name[:3])}:
            return i
    raise ValueError(f"Geçersiz ay başlığı: {value!r}")


def period(value):
    if isinstance(value, date):
        return value.year, value.month
    year, m = str(value).strip().split("-")
    if not 1 <= int(m) <= 12:
        raise ValueError(f"Geçersiz dönem: {value}")
    return int(year), int(m)


def label(p):
    return f"{p[0]:04d}-{p[1]:02d}"


def selected_months(value, default):
    if value is None:
        return [default]
    entries = value.split(",") if isinstance(value, str) else list(value)
    if not entries:
        raise ValueError("En az bir ay seçin")
    if any(
        isinstance(entry, bool) or not re.fullmatch(r"[0-9]{1,2}", str(entry).strip()) for entry in entries
    ):
        raise ValueError("Ayları 1–12 arasında, virgülle ayrılmış tam sayılarla girin")
    months = sorted({int(str(entry).strip()) for entry in entries})
    if months[0] < 1 or months[-1] > 12:
        raise ValueError("Aylar 1–12 arasında olmalı")
    return months


def month_caption(months, year):
    return ", ".join(MONTHS[m - 1] for m in months) + f" {year}"


def month_signature(months):
    return "aylar:" + ",".join(f"{m:02d}" for m in months)


def total(values):
    values = list(values)
    return sum(values, ZERO) if values and None not in values else None


def delta(a, b):
    return None if a is None or b is None else a - b


def ratio(a, b):
    return None if a is None or b in (None, ZERO) else a / b


def excel_bounds(value, setting, allow_cell=False):
    address = str(value).strip().upper().replace("$", "")
    pattern = r"[A-Z]{1,3}[1-9][0-9]{0,6}"
    if not re.fullmatch(f"{pattern}(?::{pattern})" + ("?" if allow_cell else ""), address):
        raise ValueError(f"{setting}: Excel hücre aralığı girin ({value!r})")
    left, top, right, bottom = range_boundaries(address)
    if left > right or top > bottom or right > 16384 or bottom > 1048576:
        raise ValueError(f"{setting}: Excel aralığı sınırları geçersiz ({value!r})")
    return left, top, right, bottom


def dated_columns(sheet, start, limit=None, year=None, month_labels=False):
    left, row, _, _ = excel_bounds(start, "Tarih başlangıcı", allow_cell=True)
    periods = {}
    for col in range(left, (limit or sheet.max_column) + 1):
        value = sheet.cell(row, col).value
        if value in (None, ""):
            if year is None:
                break
            continue
        try:
            p = index_period(value)
        except ValueError:
            if year is None:
                raise
            if not month_labels:
                continue
            p = (year, month(value))
        if year is not None and p[0] != year:
            continue
        if p in periods:
            raise ValueError(f"Mükerrer tarih: {label(p)}")
        periods[p] = col
    return periods


def positioned_columns(sheet, cfg, closing):
    start_col, _, _, _ = excel_bounds(cfg["tarih_baslangici"], "Gerçekleşen", allow_cell=True)
    budget_col, marker_row, _, _ = excel_bounds(cfg["butce_etiketi"], "Bütçe", allow_cell=True)
    starts = [
        c
        for c in range(budget_col, sheet.max_column + 1)
        if key(sheet.cell(marker_row, c).value) in {"bütçe", "butce"}
    ]
    if not starts:
        raise ValueError(f"Bütçe başlığı bulunamadı: {cfg['butce_etiketi']}")
    actual = dated_columns(sheet, cfg["tarih_baslangici"], starts[0] - 1)
    expected_budget = {(closing[0], m) for m in range(1, 13)}
    budgets = [
        dated_columns(sheet, f"{get_column_letter(c)}{marker_row + 1}", c + 11, closing[0], month_labels=True)
        for c in starts
    ]
    valid = [budget for budget in budgets if set(budget) == expected_budget]
    if closing not in actual or len(valid) != 1:
        raise ValueError("Gerçekleşen kapanışı veya kapanış yılının tek 12 aylık bütçe bloğu bulunamadı")
    budget = valid[0]
    return actual, budget


def item_count(cfg):
    """Gider kalemi sayısı: ayar yoksa dip toplam satırına kadar, o da yoksa 27. VBA: MakroOrtak.ItemCount."""
    if cfg.get("gider_kalem_sayisi"):
        return cfg["gider_kalem_sayisi"]
    if cfg.get("dip_toplam_satiri") and cfg["dip_toplam_satiri"] > cfg["ilk_gider_satiri"]:
        return cfg["dip_toplam_satiri"] - cfg["ilk_gider_satiri"]
    return 27


def is_period(value):
    try:
        index_period(value)
    except (ValueError, TypeError):
        return False
    return value not in (None, "")


def layout_hint(sheet, cfg):
    """Tarih ve Bütçe başlığı ayardaki hücrede değilse yakın bir kaymayı arar ve önerir."""
    left, top, _, _ = excel_bounds(cfg["tarih_baslangici"], "Gerçekleşen", allow_cell=True)
    b_left, b_top, _, _ = excel_bounds(cfg["butce_etiketi"], "Bütçe", allow_cell=True)
    if is_period(sheet.cell(top, left).value) and key(sheet.cell(b_top, b_left).value) in {"bütçe", "butce"}:
        return ""  # Başlıklar ayardaki yerinde; hata düzen kaymasından değil.
    for dr, dc in sorted(
        ((r, c) for r in range(-6, 7) for c in range(-6, 7) if (r, c) != (0, 0)),
        key=lambda x: (abs(x[0]) + abs(x[1]), x),
    ):
        if min(top + dr, b_top + dr, left + dc, b_left + dc) < 1:
            continue
        if is_period(sheet.cell(top + dr, left + dc).value) and key(
            sheet.cell(b_top + dr, b_left + dc).value
        ) in {"bütçe", "butce"}:
            start = f"{get_column_letter(left + dc)}{top + dr}"
            marker = f"{get_column_letter(b_left + dc)}{b_top + dr}"
            return (
                f"Kaynak düzeni kaymış görünüyor ({sheet.title}): tarih_baslangici {cfg['tarih_baslangici']} yerine "
                f"{start}, butce_etiketi {cfg['butce_etiketi']} yerine {marker} olabilir; satır ayarları da "
                f"{dr:+d} kaymış olabilir. Tablo düzeni ekranından kontrol edip güncelleyin."
            )
    return ""


def positioned_cost(sheet, cfg, closing, issues=None):
    try:
        actual, budget = positioned_columns(sheet, cfg, closing)
    except ValueError as error:
        hint = layout_hint(sheet, cfg)
        raise ValueError(f"{error}. {hint}" if hint else str(error)) from error
    start_col, _, _, _ = excel_bounds(cfg["tarih_baslangici"], "Gerçekleşen", allow_cell=True)
    issues = issues or SourceIssues()

    def values(row, columns):
        return {p: issues.value(sheet, row, c) for p, c in columns.items()}

    first = cfg["ilk_gider_satiri"]
    size = item_count(cfg)
    name_col = cfg.get("gider_sutunu", start_col - 1)
    items = {}
    for code, row in enumerate(range(first, first + size), 1):
        name = sheet.cell(row, name_col).value
        if not isinstance(name, str) or not name.strip():
            hint = layout_hint(sheet, cfg)
            raise ValueError(
                f"Gider adı bulunamadı: {sheet.title}!{get_column_letter(name_col)}{row}"
                + (f". {hint}" if hint else "")
            )
        items[code] = (name.strip(), values(row, actual), values(row, budget))
    bottom = {p: total(item[1][p] for item in items.values()) for p in actual}
    return items, bottom, values(cfg["islem_adedi_satiri"], actual)


def merge_cost_items(items, mappings):
    if not mappings:
        return items
    merged = {}
    for code, record in items.items():
        target, name = mappings.get(code, (code, record[0]))
        merged.setdefault(target, []).append((name, record))
    result = {}
    for target, records in merged.items():
        names = {name for name, _ in records}
        if len(names) != 1:
            raise ValueError(f"{target} rapor indeksinde gider adları farklı. Eşleştirmeyi tamamlayın.")
        result[target] = (
            next(iter(names)),
            *[
                {
                    p: total(record[slot].get(p) for _, record in records)
                    for p in set().union(*(record[slot] for _, record in records))
                }
                for slot in (1, 2)
            ],
        )
    return result


def read_cost(sheet, cfg, closing, issues=None):
    if cfg.get("tarih_baslangici"):
        return positioned_cost(sheet, cfg, closing, issues)

    def cell(row, col):
        return sheet.cell(row, col).value

    actual, budget = {}, {}

    def add_period(target, col, y, m):
        p = (int(y), month(m))
        if p in target:
            raise ValueError(f"Tekrarlanan dönem: {label(p)}")
        target[p] = col

    actual_range = cfg.get("gerceklesen_araligi")
    budget_range = cfg.get("butce_araligi")
    if bool(actual_range) != bool(budget_range):
        raise ValueError("Gerçekleşen ve bütçe aralıklarını birlikte tanımlayın")
    if actual_range:
        a_left, a_top, a_right, a_bottom = excel_bounds(actual_range, "Gerçekleşen")
        b_left, b_top, b_right, b_bottom = excel_bounds(budget_range, "Bütçe")
        if a_right >= b_left:
            raise ValueError("Gerçekleşen ve bütçe aralıkları çakışıyor veya sırasız")
        for target, left, top, right, bottom in (
            (actual, a_left, a_top, a_right, a_bottom),
            (budget, b_left, b_top, b_right, b_bottom),
        ):
            needed = [cfg["yil_satiri"], cfg["ay_satiri"], cfg["ilk_gider_satiri"], cfg["dip_toplam_satiri"]]
            if target is actual and cfg.get("islem_adedi_satiri"):
                needed.append(cfg["islem_adedi_satiri"])
            if min(needed) < top or max(needed) > bottom:
                raise ValueError("Maliyet aralığı gerekli başlık ve veri satırlarını kapsamıyor")
            for col in range(left, right + 1):
                add_period(target, col, cell(cfg["yil_satiri"], col), cell(cfg["ay_satiri"], col))
    else:
        target = actual
        for col in range(cfg["ilk_donem_sutunu"], sheet.max_column + 1):
            y, m = cell(cfg["yil_satiri"], col), cell(cfg["ay_satiri"], col)
            if y is None and m is None:
                if not actual:
                    raise ValueError("Gerçekleşen başlığı bulunamadı")
                if target is budget:
                    break
                target = budget
                continue
            add_period(target, col, y, m)

    for periods in (actual, budget):
        if not periods or list(periods) != sorted(periods):
            raise ValueError("Dönem başlıkları eksik veya sırasız")
        for y in {p[0] for p in periods}:
            if any((y, m) not in periods for m in range(1, 13)):
                raise ValueError(f"{y} döneminde 12 ay bulunamadı")
    if closing not in actual or set(y for y, _ in budget) != {closing[0]}:
        raise ValueError("Dosya dönemi ile gerçekleşen/bütçe başlıkları uyuşmuyor")

    def values(row, columns):
        return {p: number(cell(row, col)) for p, col in columns.items()}

    items = {}
    for r in range(cfg["ilk_gider_satiri"], cfg["dip_toplam_satiri"]):
        index = cell(r, cfg["indeks_sutunu"])
        if index is None:
            continue
        code = int(index)
        name = cell(r, cfg["gider_sutunu"])
        if code in items or not name:
            raise ValueError(f"Geçersiz veya mükerrer gider indeksi: {index}, satır {r}")
        items[code] = (name, values(r, actual), values(r, budget))
    if not items:
        raise ValueError("Gider satırı bulunamadı")
    transactions = values(cfg["islem_adedi_satiri"], actual) if cfg.get("islem_adedi_satiri") else {}
    return items, values(cfg["dip_toplam_satiri"], actual), transactions


def filename_period(path):
    tokens = re.findall(r"[^\W_]+", path.stem)
    periods = set()
    for left, right in zip(tokens, tokens[1:]):
        year, name = (left, right) if re.fullmatch(r"[0-9]{4}", left) else (right, left)
        if not re.fullmatch(r"[0-9]{4}", year) or name.isdecimal():
            continue
        try:
            periods.add((int(year), month(name)))
        except ValueError:
            continue
    if len(periods) > 1:
        raise ValueError("Dosya adında birden fazla kapanış dönemi bulundu")
    if not periods:
        raise ValueError("Dosya adında yıl ve ay yan yana bulunamadı")
    return periods.pop()


def cost_files(root):
    if not root.is_dir():
        raise ValueError(f"Maliyet klasörü bulunamadı: {root}")
    for path in sorted(root.iterdir()):
        if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$"):
            yield path


def read_snapshots(root, config, closing, warnings, progress=None):
    snapshots = {}
    seen = set()
    ignored = {key(s) for s in config["ortak_sayfalar"]}
    configured = {key(s): s for s in config["bolumler"]}
    for path in cost_files(root):
        if progress:
            progress("Kaynak dosyası: " + path.name)
        try:
            p = filename_period(path)
        except ValueError as error:
            warnings.append([None, None, path.name, str(error)])
            continue
        if not in_closing_scope(p, closing):
            continue
        book = load_workbook(path, data_only=True)
        formulas = None
        try:
            for sheet in book:
                name = configured.get(key(sheet.title))
                if key(sheet.title) in ignored:
                    continue
                if name is None:
                    warnings.append([label(p), sheet.title, path.name, "Bölüm eşleştirmesi bulunamadı"])
                    continue
                k = (p, name)
                if k in seen:
                    snapshots.pop(k, None)
                    warnings.append(
                        [label(p), name, path.name, "Aynı bölüm ve kapanış için birden fazla kaynak"]
                    )
                    continue
                seen.add(k)
                issues = SourceIssues()
                try:
                    items, bottom, counts = read_cost(sheet, config["maliyet"], p, issues)
                    snapshots[k] = (
                        merge_cost_items(items, config.get("gider_eslestirme", {})),
                        bottom,
                        counts,
                        path.name,
                    )
                except (ValueError, TypeError) as error:
                    warnings.append([label(p), name, path.name, str(error)])
                for message in issues.warnings(sheet.title):
                    warnings.append([label(p), name, path.name, message])
                if issues.empty:
                    if formulas is None:
                        formulas = load_workbook(path)
                    uncached = [a for a in issues.empty if formulas[sheet.title][a].data_type == "f"]
                    if uncached:
                        warnings.append(
                            [
                                label(p),
                                name,
                                path.name,
                                f"{len(uncached)} formül hücresinin sonucu dosyada kayıtlı değil ({sheet.title}: "
                                f"{sample_cells(uncached)}). Dosyayı Excel'de açıp kaydedin; bu hücreler boş okundu.",
                            ]
                        )
        finally:
            book.close()
            if formulas is not None:
                formulas.close()
    return snapshots


def in_closing_scope(p, closing):
    """Kapanış yılının ≤ kapanış dosyaları; Ocak kapanışında yıl dönümü revizyonu için önceki Aralık da."""
    return p <= closing and (p[0] == closing[0] or (closing[1] == 1 and p == (closing[0] - 1, 12)))


def control_period(value, year):
    if isinstance(value, date) or re.fullmatch(r"\d{4}-\d{1,2}", str(value)):
        return period(value)
    return year, month(value)


def positioned_control(sheet, cfg, year, directors):
    budget_left, _, _, _ = excel_bounds(cfg["butce_ciro_baslangici"], "Bütçe ciro", allow_cell=True)
    actual = dated_columns(sheet, cfg["ciro_baslangici"])
    if not actual:
        raise ValueError("Ciro tarihleri bulunamadı")
    _, header, _, _ = excel_bounds(cfg["ciro_baslangici"], "Ciro", allow_cell=True)
    expected = {(year, m) for m in range(1, 13)}
    budgets = []
    for c in range(max(max(actual.values()) + 1, budget_left), sheet.max_column - 10):
        try:
            block = dated_columns(
                sheet, f"{get_column_letter(c)}{header}", c + 11, year=year, month_labels=True
            )
        except ValueError:
            continue
        if set(block) == expected:
            budgets.append(block)
    if len(budgets) != 1:
        raise ValueError("Ciro için gerçekleşenden sonraki tek 12 aylık bütçe bloğu bulunamadı")
    planned = budgets[0]
    company = {p: number(sheet.cell(header + 1, c).value) for p, c in actual.items()}
    company_budget = {p: number(sheet.cell(header + 1, c).value) for p, c in planned.items()}
    title_col = cfg["etiket_sutunu"]
    rows = {}
    for r in range(header + 2, sheet.max_row + 1):
        director = key(sheet.cell(r, title_col).value)
        if director not in directors or key(sheet.cell(r + 2, title_col).value) != "opex toplamı":
            continue
        rows.setdefault((director, "opex toplamı"), []).append(
            {p: number(sheet.cell(r + 2, c).value) for p, c in actual.items()}
        )
        for k in range(r + 3, sheet.max_row + 1):
            name = key(sheet.cell(k, title_col).value)
            if not name or name in directors:
                break
            rows.setdefault((director, name), []).append(
                {p: number(sheet.cell(k, c).value) for p, c in actual.items()}
            )
    return company, rows, company_budget


def read_control(path, cfg, year, directors=()):
    book = load_workbook(path, data_only=True)
    try:
        sheet = book[cfg["sayfa"]]
        if cfg.get("ciro_baslangici"):
            return positioned_control(sheet, cfg, year, set(directors))
        company, rows = {}, {}
        left = cfg["etiket_sutunu"] - 1
        first = cfg["ilk_ay_sutunu"] - 1
        data = list(sheet.values)
        if cfg.get("ciro_araligi"):
            left_col, top, right, bottom = excel_bounds(cfg["ciro_araligi"], "Ciro")
            if bottom != top + 1:
                raise ValueError("Ciro aralığında üst satır aylar, alt satır tutarlar olmalı")
            header, values = sheet.iter_rows(
                min_row=top, max_row=bottom, min_col=left_col, max_col=right, values_only=True
            )
            first_revenue = 0
        else:
            header = data[cfg["ciro_satiri"] - 2]
            values = data[cfg["ciro_satiri"] - 1]
            first_revenue = first
        for col in range(first_revenue, len(header)):
            if header[col] is not None:
                p = control_period(header[col], year)
                if p in company:
                    raise ValueError(f"Mükerrer ciro dönemi: {label(p)}")
                company[p] = number(values[col])

        director = None
        columns = {}
        for row in data[cfg["ilk_blok_satiri"] - 1 :]:
            title = row[left]
            cells = {i: row[i] for i in range(first, len(row)) if row[i] is not None}
            if not any(v is not None and str(v).strip() for v in row):
                director, columns = None, {}
            elif title and not cells:
                director, columns = key(title), {}
            elif key(title) in {"bölüm", "bolum"} and director:
                columns = {i: control_period(v, year) for i, v in cells.items()}
                if len(set(columns.values())) != len(columns):
                    raise ValueError(f"Mükerrer kontrol dönemi: {director}")
            elif director and columns and title:
                rows.setdefault((director, key(title)), []).append(
                    {p: number(row[i]) for i, p in columns.items()}
                )
        return company, rows, {}
    finally:
        book.close()


def index_period(value):
    if isinstance(value, date):
        return value.year, value.month
    text = str(value).strip()
    try:
        if re.fullmatch(r"\d{4}-\d{1,2}", text):
            return period(text)
        if "." in text:
            day, month_value, year = map(int, text.split("."))
            parsed = date(year, month_value, day)
        else:
            parsed = date.fromisoformat(text)
        return parsed.year, parsed.month
    except ValueError as exc:
        raise ValueError(f"Geçersiz endeks tarihi: {value}") from exc


def period_index(indices, name, periods):
    factors = []
    for year, m in periods:
        current = indices.get((key(name), (year, m)))
        previous = indices.get((key(name), (year - 1, m)))
        if current is None or previous is None or min(current, previous) <= ZERO:
            return None
        factors.append(current / previous)
    return ratio(total(factors), len(factors))


def read_indices(path, cfg):
    book = load_workbook(path, data_only=True, read_only=True)
    result = {}
    try:
        sheet = book[cfg["sayfa"]]
        if cfg.get("baslik_satiri") or cfg.get("endeks_araligi"):
            area = cfg.get("endeks_araligi")
            left, top, right, bottom = (3, cfg.get("baslik_satiri"), None, None)
            if area:
                left, top, right, bottom = excel_bounds(area, "Endeks", allow_cell=True)
                if ":" not in area:
                    right, bottom = None, None
            rows = sheet.iter_rows(min_row=top, min_col=left, max_col=right, max_row=bottom, values_only=True)
            header = next(rows, ())
            if not header or key(header[0]) not in {"bölüm adı", "bölüm gider endeksi"}:
                raise ValueError(
                    "Endeks aralığının sol üstünde Bölüm Adı veya Bölüm Gider Endeksi başlığı bulunamadı"
                )
            periods = []
            for value in header[1:]:
                if value in (None, ""):
                    break
                p = index_period(value)
                if periods:
                    y, m = periods[-1]
                    expected = (y + 1, 1) if m == 12 else (y, m + 1)
                    if p != expected:
                        raise ValueError(
                            f"Endeks ayları ardışık değil: beklenen {label(expected)}, bulunan {label(p)}"
                        )
                periods.append(p)
            if not periods:
                raise ValueError("İkinci endeks tablosunda ay başlığı bulunamadı")
            seen = set()
            for row in rows:
                name = key(row[0])
                if not name:
                    continue
                if name in seen:
                    raise ValueError(f"Mükerrer endeks bölümü: {row[0]}")
                seen.add(name)
                result.update({(name, p): number(value) for p, value in zip(periods, row[1:])})
            return result
        for name, p, value in sheet.iter_rows(min_row=cfg["ilk_veri_satiri"], max_col=3, values_only=True):
            if name is None and p is None and value is None:
                continue
            k = key(name), period(p)
            if k in result:
                raise ValueError(f"Mükerrer endeks kaydı: {name}, {p}")
            result[k] = number(value)
        return result
    finally:
        book.close()


def revision_pairs(year, selected):
    """(önceki kapanış, yeni kapanış, karşılaştırılacak [(tür, slot, etkilenen dönemler)]) üçlüleri.

    Ocak kapanışında önceki yılın Aralık dosyasıyla geçen yılın 12 ayının gerçekleşenleri karşılaştırılır;
    bütçe yılları farklı olduğundan bütçe karşılaştırılmaz. VBA: MakroHesap.CalculateChecks.
    """
    if selected == 1:
        yield (year - 1, 12), (year, 1), [("Gerçekleşen", 1, [(year - 1, m) for m in range(1, 13)])]
    for m in range(2, selected + 1):
        yield (year, m - 1), (year, m), [
            ("Gerçekleşen", 1, [(year, a) for a in range(1, m)]),
            ("Bütçe", 2, [(year, a) for a in range(1, 13)]),
        ]


def revisions(snapshots, config, closing):
    year, selected = closing
    for name in config["bolumler"]:
        for before_close, p, comparisons in revision_pairs(year, selected):
            older = snapshots.get((before_close, name))
            newer = snapshots.get((p, name))
            if not older or not newer:
                continue
            old_items, _, _, old_file = older
            new_items, _, _, new_file = newer
            for code in sorted(old_items.keys() | new_items.keys()):
                old = old_items.get(code, ("", {}, {}))
                new = new_items.get(code, ("", {}, {}))
                for kind, slot, affected_periods in comparisons:
                    for affected_period in affected_periods:
                        before = old[slot].get(affected_period)
                        after = new[slot].get(affected_period)
                        if before != after:
                            yield [
                                label(p),
                                label(affected_period),
                                config["bolumler"][name]["mudurluk"],
                                name,
                                code,
                                new[0] or old[0],
                                kind,
                                before,
                                after,
                                delta(after, before),
                                old_file,
                                new_file,
                            ]


def make_report(config, closing, report_period, months, snapshots, company, controls, indices, warnings):
    detail, checks, coverage = [], [], []
    year = report_period[0]
    for name, mapping in config["bolumler"].items():
        snapshot = snapshots.get((closing, name))
        if snapshot is None:
            warnings.append([label(closing), name, None, "Kapanış dosyası yok veya kaynak okunamadı"])
            coverage.append(name)
            continue
        items, bottom, _, source = snapshot
        for before_close, after_close, _ in revision_pairs(*closing):
            if not (snapshots.get((before_close, name)) and snapshots.get((after_close, name))):
                warnings.append(
                    [
                        label(after_close),
                        name,
                        None,
                        "Ardışık dosya eksik. Bu kapanış için revizyon karşılaştırması yapılamadı.",
                    ]
                )
        index = period_index(indices, mapping["endeks"], [(year, m) for m in months])
        if index is None and mapping.get("endeks"):
            warnings.append([label(closing), name, None, "Seçilen dönem için endeks hesabı tamamlanamadı"])
        for code, (item, actuals, budget) in sorted(items.items()):
            group = next(
                (g for low, high, g in config["gider_gruplari"] if low <= code <= high), "Eşleşmeyen"
            )
            actual = total(actuals.get((year, m)) for m in months)
            prior = total(actuals.get((year - 1, m)) for m in months)
            planned = total(budget.get((year, m)) for m in months)
            if actual is None or prior is None or (year == closing[0] and planned is None):
                warnings.append(
                    [
                        label(report_period),
                        name,
                        source,
                        f"{code} {item}: karşılaştırma için bir veya daha fazla aylık tutar eksik",
                    ]
                )
            base = prior * index if prior is not None and index is not None else None
            savings = delta(base, actual)
            detail.append(
                [
                    name,
                    mapping["mudurluk"],
                    code,
                    item,
                    group,
                    actual,
                    prior,
                    planned,
                    delta(actual, planned),
                    ratio(delta(actual, planned), planned),
                    delta(actual, prior),
                    ratio(delta(actual, prior), prior),
                    index,
                    base,
                    savings,
                    ratio(savings, base),
                    source,
                ]
            )
    control_groups = {}
    for name, mapping in config["bolumler"].items():
        target = mapping.get("kontrol_bolumu") or name
        control_groups.setdefault((key(mapping["direktorluk"]), key(target)), [target, []])[1].append(name)
    for (director, normalized), (target, names) in control_groups.items():
        matches = controls.get((director, normalized), [])
        for m in months:
            p = (year, m)
            sources = [snapshots.get((closing, name)) for name in names]
            a = total(source[1].get(p) if source else None for source in sources)
            b = matches[0].get(p) if len(matches) == 1 else None
            diff = delta(a, b)
            status = (
                "Bölüm eşleşmedi"
                if not matches
                else (
                    "Birden fazla bölüm satırı eşleşti"
                    if len(matches) > 1
                    else (
                        "Kaynak dip toplamı eksik"
                        if a is None
                        else (
                            "Kontrol tutarı veya ay bulunamadı"
                            if b is None
                            else (
                                "Veri doğrulanamadı: tutar farklı"
                                if abs(diff) > Decimal(str(config["kontrol"]["tolerans"]))
                                else "Doğrulandı"
                            )
                        )
                    )
                )
            )
            managers = sorted({config["bolumler"][name]["mudurluk"] for name in names})
            files = ", ".join(dict.fromkeys(source[3] for source in sources if source))
            checks.append(
                [
                    target,
                    ", ".join(managers),
                    config["bolumler"][names[0]]["direktorluk"],
                    label(p),
                    a,
                    b,
                    diff,
                    status,
                    files,
                    ", ".join(names),
                    tuple(names),
                ]
            )
    return detail, checks, coverage


def unit_cost_rows(config, closing, report_period, months, snapshots, indices, warnings, budget_counts=None):
    if not config["maliyet"].get("islem_adedi_satiri"):
        return []
    year = report_period[0]
    current = [(year, m) for m in months]
    prior = [(year - 1, m) for m in months]
    rows = []
    known_indices = {k for k, _ in indices}
    for name, mapping in config["bolumler"].items():
        if not mapping.get("endeks") or key(mapping["endeks"]) not in known_indices:
            continue
        snapshot = snapshots.get((closing, name))
        if snapshot is None:
            continue
        _, costs, counts, source = snapshot
        for p in current + prior:
            count = counts.get(p)
            if count is None or count <= ZERO:
                warnings.append([label(p), name, source, "İşlem adedi eksik veya sıfır/negatif"])
            if costs.get(p) is None:
                warnings.append([label(p), name, source, "Birim maliyet için gider dip toplamı eksik"])
        a, b = [total(costs.get(p) for p in span) for span in (current, prior)]
        x, y = [
            total(counts[p] if counts.get(p) is not None and counts[p] > ZERO else None for p in span)
            for span in (current, prior)
        ]
        index = period_index(indices, mapping["endeks"], current)
        adjusted = b * index if b is not None and index is not None else None
        comparable = (
            a is not None
            and a >= ZERO
            and x is not None
            and x > ZERO
            and b is not None
            and b > ZERO
            and y is not None
            and y > ZERO
            and index is not None
        )
        if (year == closing[0] and not comparable) or (
            year != closing[0] and (a is None or x is None or x <= ZERO)
        ):
            warnings.append(
                [
                    label(report_period),
                    name,
                    source,
                    "Birim maliyet kapsamı dışında: karşılaştırılabilir veri eksik",
                ]
            )
            continue
        planned_cost = total(
            (
                sum((v[2].get(p) for v in snapshot[0].values()), ZERO)
                if all(v[2].get(p) is not None for v in snapshot[0].values())
                else None
            )
            for p in current
        )
        planned_count = total(
            v if (v := (budget_counts or {}).get(name, {}).get(p)) is not None and v > ZERO else None
            for p in current
        )
        planned_bm = (
            ratio(planned_cost, planned_count)
            if year == closing[0] and planned_count is not None and planned_count > ZERO
            else None
        )
        rows.append([name, mapping["mudurluk"], a, x, b, y, adjusted, index, planned_bm])
    by_name = {row[0]: row for row in rows}
    departments = config["bolumler"]
    tree = department_tree(departments)

    def scoped_cost(names):
        return total(by_name[n][2] for n in names if n in by_name)

    def improvement(row):
        actual, base = ratio(row[2], row[3]), ratio(row[6], row[5])
        return ratio(delta(base, actual), base)

    weights = {}
    for managers in tree.values():
        director_cost = scoped_cost([n for names in managers.values() for n in names])
        for names in managers.values():
            manager_cost = scoped_cost(names)
            for name in names:
                row = by_name.get(name)
                if row is None:
                    continue
                rate = improvement(row)
                manager_weight = ratio(row[2], manager_cost)
                director_weight = ratio(row[2], director_cost)
                weights[name] = [
                    manager_weight,
                    manager_weight * rate if manager_weight is not None and rate is not None else None,
                    director_weight,
                    director_weight * rate if director_weight is not None and rate is not None else None,
                ]

    result = []
    for name, level, names in hierarchy(departments):
        manager = departments[name]["mudurluk"] if level == "Bölüm" else level
        names = [n for n in names if n in by_name]
        if not names:
            result.append([name, manager, *([None] * 14)])
            continue
        values = [by_name[department] for department in names]
        a, x, b, y, adjusted = [total(row[i] for row in values) for i in range(2, 7)]
        actual, previous, base = ratio(a, x), ratio(b, y), ratio(adjusted, y)
        if level == "Bölüm":
            rate = improvement(values[0])
            allocation = weights[name]
        else:
            allocation = [total(weights[n][i] for n in names) for i in range(4)]
            if level == "Müdürlük toplamı":
                rate = allocation[1]
            elif level == "Direktörlük toplamı":
                rate = allocation[3]
                allocation[:2] = [None, None]
            else:
                rate = ratio(
                    total(
                        (
                            by_name[n][2] * improvement(by_name[n])
                            if by_name[n][2] is not None and improvement(by_name[n]) is not None
                            else None
                        )
                        for n in names
                    ),
                    a,
                )
                allocation = [None] * 4
        result.append(
            [
                name,
                manager,
                a,
                x,
                actual,
                b,
                y,
                previous,
                base,
                rate,
                values[0][7] if level == "Bölüm" else None,
                *allocation,
                values[0][8] if level == "Bölüm" else None,
            ]
        )
    return result


def aggregate(rows):
    a, p, b = [total(row[i] for row in rows) for i in (5, 6, 7)]
    base, saving = [total(row[i] for row in rows) for i in (13, 14)]
    weighted = total(
        row[5] * row[15] if row[5] is not None and row[15] is not None else ZERO if row[5] == ZERO else None
        for row in rows
    )
    return [
        a,
        p,
        b,
        delta(a, b),
        ratio(delta(a, b), b),
        delta(a, p),
        ratio(delta(a, p), p),
        base,
        saving,
        ratio(weighted, a),
    ]


def add_table(book, name, headers, rows, percentages=()):
    sheet = book.create_sheet(name)
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    if any("\n" in str(title) for title in headers):
        sheet.row_dimensions[1].height = 32
    for cell in sheet[1]:
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.fill = PatternFill("solid", fgColor="E7EEF2")
        cell.font = Font(color="263746", bold=True)
    for i, cells in enumerate(sheet.columns, 1):
        width = max(len(line) for c in cells for line in str(c.value or "").split("\n"))
        sheet.column_dimensions[get_column_letter(i)].width = max(14, min(62, width + 3))
        for c in cells[1:]:
            if c.data_type == "f":
                c.data_type = "s"
            if (
                c.data_type == "n"
                and c.value is not None
                and headers[i - 1] not in {"İndeks", "Doğrulanamayan bölüm-ay", "Eksik bölüm"}
            ):
                c.number_format = (
                    "0.00%"
                    if i in percentages
                    else (
                        "0.000"
                        if str(headers[i - 1]).startswith("Dönem endeksi")
                        else "#,##0.00;-#,##0.00;0.00"
                    )
                )


def department_tree(departments):
    directors = {}
    for department, mapping in departments.items():
        managers = directors.setdefault(mapping["direktorluk"], {})
        managers.setdefault(mapping["mudurluk"], []).append(department)
    return directors


def hierarchy(departments):
    directors = department_tree(departments)
    for director, managers in directors.items():
        director_departments = []
        for manager, names in managers.items():
            for department in names:
                yield department, "Bölüm", [department]
            yield manager, "Müdürlük toplamı", names
            director_departments.extend(names)
        yield director, "Direktörlük toplamı", director_departments
    if len(directors) > 1:
        yield "Tüm direktörlükler", "Genel toplam", list(departments)


def style_subtotals(sheet, subtotal_rows):
    for r, level in subtotal_rows:
        light = level == "Müdürlük toplamı"
        for cell in sheet[r]:
            cell.fill = PatternFill("solid", fgColor="EFF3F5" if light else "D8E4E8")
            cell.font = Font(bold=True, color="263746")


def add_general_summary(book, detail, layout, missing, metrics, operating, revenue, ratio_header):
    rows, subtotals = [], []
    for name, level, names in layout:
        incomplete = missing.intersection(names)
        values = (
            [None] * len(metrics)
            if incomplete
            else aggregate([row for row in detail if row[0] in names])[: len(metrics)]
        )
        values[3] = delta(values[2], values[0])
        values[0], values[1] = values[1], values[0]
        opx = None if incomplete else total(operating.get(name) for name in names)
        rows.append([name, level, *values, ratio(opx, revenue)])
        if level != "Bölüm":
            subtotals.append((len(rows) + 1, level))
    add_table(
        book,
        "Genel Özet",
        ["Bölüm / müdürlük / direktörlük", "Satır türü", *metrics, ratio_header],
        rows,
        (7, 8),
    )
    style_subtotals(book["Genel Özet"], subtotals)


def add_expense_matrix(book, detail, departments, period_name, editable):
    tree = department_tree(departments)
    names = [name for managers in tree.values() for group in managers.values() for name in group]
    groups = {}
    for row in detail:
        groups.setdefault(row[4], {}).setdefault((row[2], row[3]), []).append(row)

    block_size = 5 if editable else 4

    def expense_row(group, code, title, records):
        by_department = {}
        for row in records:
            by_department.setdefault(row[0], []).append(row)
        values = [group, code, title]
        for name in names:
            rows = by_department.get(name, ())
            actual, budget = (total(row[i] for row in rows) for i in (5, 7))
            values.extend([actual, budget, delta(budget, actual), ratio(delta(actual, budget), budget)])
            if editable:
                values.append(None)
        return values

    columns = ("Gerçekleşen (TL)", "Bütçe (TL)", "Bütçe − gerçekleşen (TL)", "Bütçe aşım oranı (%)")
    headers = ["Gider grubu", "İndeks", "Gider kalemi"]
    for _ in names:
        headers.extend(f"{title}\n{period_name}" for title in columns)
        if editable:
            headers.append("Açıklama")
    rows, outlines, totals = [], [], []
    for group, items in groups.items():
        totals.append(len(rows) + 6)
        rows.append(
            expense_row(group, None, "Grup toplamı", [row for records in items.values() for row in records])
        )
        start = len(rows) + 6
        for (code, name), records in sorted(items.items()):
            rows.append(expense_row(group, code, name, records))
        outlines.append((start, len(rows) + 5))
    rows.append(expense_row("TOPLAM", None, "Tüm giderler", detail))
    sheet_name = "Giderler"
    add_table(book, sheet_name, headers, rows, tuple(range(7, len(headers) + 1, block_size)))
    sheet = book[sheet_name]
    sheet.insert_rows(1, 4)
    sheet.row_dimensions[1].height = 23
    sheet.row_dimensions[4].height = 23
    sheet.row_dimensions[5].height = 34
    sheet.freeze_panes = "D6"
    sheet.auto_filter.ref = f"A5:{get_column_letter(sheet.max_column)}{sheet.max_row}"
    sheet.merge_cells("A1:C1")
    sheet["A1"] = f"{sheet_name} · {period_name}"
    sheet.merge_cells("A2:C2")
    sheet["A2"] = "Alt giderleri açmak için soldaki + işaretini kullanın"
    sheet.merge_cells("D2:F2")
    sheet["D2"] = "Bütçe aşıldı"
    sheet.merge_cells("G2:I2")
    sheet["G2"] = "Bütçe altında"
    sheet.merge_cells("A4:C4")
    sheet["A4"] = period_name

    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    sheet["D2"].fill = red
    sheet["G2"].fill = green
    edge = Side(style="thick", color="233C57")
    last = sheet.max_row
    for i, name in enumerate(names):
        start = 4 + block_size * i
        end = start + block_size - 1
        sheet.merge_cells(start_row=4, start_column=start, end_row=4, end_column=end)
        sheet.cell(4, start, name)
        if editable:
            sheet.column_dimensions[get_column_letter(end)].width = 42
            for r in range(6, last):
                if isinstance(sheet.cell(r, 2).value, int):
                    note = sheet.cell(r, end)
                    note.fill = PatternFill("solid", fgColor="FFF4DA")
                    note.alignment = Alignment(wrap_text=True, vertical="top")
                    note.number_format = "@"
        for r in range(4, last + 1):
            for c in range(start, end + 1):
                sheet.cell(r, c).border = Border(
                    top=edge if r == 4 else None,
                    bottom=edge if r == last else None,
                    left=edge if c == start else None,
                    right=edge if c == end else None,
                )
            if r >= 6:
                difference = sheet.cell(r, start + 2)
                if difference.value is not None and difference.value != ZERO:
                    fill = red if difference.value < ZERO else green
                    difference.fill = fill
                    if sheet.cell(r, start + 3).value is not None:
                        sheet.cell(r, start + 3).fill = fill
    for cell in sheet[4]:
        cell.fill = PatternFill("solid", fgColor="E7EEF2")
        cell.font = Font(color="263746", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.sheet_properties.outlinePr.summaryBelow = False
    for start, end in outlines:
        sheet.row_dimensions.group(start, end, outline_level=1, hidden=True)
        sheet.row_dimensions[start - 1].collapsed = True
    for r in totals + [last]:
        for cell in sheet[r]:
            cell.font = Font(bold=True)


def add_audit(book, checks, revision_rows, warnings, revenue_rows):
    blocks = [
        (
            "Veri kontrolü",
            [
                "Kontrol bölümü",
                "Müdürlük",
                "Direktörlük",
                "Ay",
                "Alınan dip toplam TL",
                "Kontrol TL",
                "Fark TL",
                "Durum",
                "Kaynak",
                "Kaynak bölümler",
            ],
            [r[:-1] for r in checks],
        ),
        (
            "Revizyon geçmişi",
            [
                "Kapanış",
                "Etkilenen ay",
                "Müdürlük",
                "Bölüm",
                "İndeks",
                "Alt gider kalemi",
                "Tür",
                "Eski TL",
                "Yeni TL",
                "Fark TL",
                "Önceki dosya",
                "Yeni dosya",
            ],
            revision_rows,
        ),
        ("Uyarılar", ["Kapanış", "Bölüm", "Kaynak", "Uyarı"], warnings),
        ("Ciro kontrolü", ["Dönem", "Gerçekleşen ciro", "Bütçe ciro"], revenue_rows),
    ]
    rows, section_rows = [], []
    for title, headers, records in blocks:
        section_rows.append(len(rows) + 3)
        rows.extend([[title], headers])
        rows.extend(records)
        rows.append([])
    add_table(book, "Kontrol ve Revizyon", ["Veri doğrulama ve değişiklik kayıtları"] + [None] * 11, rows)
    sheet = book["Kontrol ve Revizyon"]
    sheet.auto_filter.ref = None
    for r in section_rows:
        for cell in sheet[r]:
            if cell.value is not None:
                cell.fill = PatternFill("solid", fgColor="DCE8F3")
                cell.font = Font(bold=True, color="233C57")
    for r in range(section_rows[1] + 1, section_rows[2] - 2):
        sheet.cell(r, 5).number_format = "0"


def read_settings(path):
    if path.suffix.lower() == ".json":
        return json.loads(path.read_text(encoding="utf-8-sig"))

    def rows(name, width, required=None):
        required = width if required is None else required
        for i, cells in enumerate(book[name].iter_rows(min_row=5, max_col=width), 5):
            values = [cell.value.strip() if isinstance(cell.value, str) else cell.value for cell in cells]
            if all(value in (None, "") for value in values):
                continue
            if any(cell.data_type == "f" for cell in cells) or any(
                value in (None, "") for value in values[:required]
            ):
                raise ValueError(f"{name}, satır {i}: Zorunlu hücrelere sabit değer girin")
            yield i, values

    def index(value, location):
        if not isinstance(value, int) or isinstance(value, bool) or value < 0:
            raise ValueError(f"{location}: Negatif olmayan tam sayı girin")
        return value

    book = load_workbook(path, read_only=True)
    try:
        sections = {
            "Dosyalar": {"maliyet_klasoru", "kontrol_dosyasi", "endeks_dosyasi"},
            "Maliyet": {"ilk_gider_satiri", "gider_sutunu"},
            "Kontrol": {"sayfa", "etiket_sutunu", "tolerans"},
            "Endeks": {"sayfa"},
            "Rapor": set(),
        }
        # Konumlu düzen (tarih_baslangici, ciro_baslangici, baslik_satiri/endeks_araligi) yoksa eski satır/sütun
        # düzeninin ayarları zorunludur. Konumlu düzen varsa eski ayarlar yalnız yedek olarak okunur.
        legacy = {
            "Maliyet": (
                "tarih_baslangici",
                {"yil_satiri", "ay_satiri", "dip_toplam_satiri", "indeks_sutunu", "ilk_donem_sutunu"},
            ),
            "Kontrol": ("ciro_baslangici", {"ciro_satiri", "ilk_blok_satiri", "ilk_ay_sutunu"}),
            "Endeks": ("baslik_satiri", {"ilk_veri_satiri"}),
        }
        report_settings = {"onemlilik_tutari", "onemlilik_orani", "ozet_kalem_sayisi", "arsiv_saklama_adedi"}
        cell_settings = {"tarih_baslangici", "butce_etiketi", "ciro_baslangici", "butce_ciro_baslangici"}
        range_settings = {"gerceklesen_araligi", "butce_araligi", "ciro_araligi", "endeks_araligi"}
        optional = {
            "Maliyet": {
                "islem_adedi_satiri",
                "gerceklesen_araligi",
                "butce_araligi",
                "tarih_baslangici",
                "butce_etiketi",
                "gider_kalem_sayisi",
                "gosterge_adi_sutunu",
            },
            "Kontrol": {"ciro_araligi", "ciro_baslangici", "butce_ciro_baslangici"},
            "Endeks": {"baslik_satiri", "endeks_araligi"},
            "Rapor": report_settings,
        }
        for section, (_, codes) in legacy.items():
            optional[section] |= codes
        settings = {section: {} for section in sections}
        for i, (section, code, value) in rows("Dosya ve Konum Ayarları", 3):
            location = f"Dosya ve Konum Ayarları, satır {i}"
            if (
                code not in (sections.get(section, set()) | optional.get(section, set()))
                or code in settings[section]
            ):
                raise ValueError(f"{location}: Tanımsız veya tekrarlanan ayar")
            if code in cell_settings:
                if not isinstance(value, str) or ":" in value:
                    raise ValueError(f"{location}: Tek bir Excel hücresi girin")
                excel_bounds(value, code, allow_cell=True)
            elif code in range_settings:
                excel_bounds(value, code, allow_cell=code == "endeks_araligi")
            elif code == "tolerans":
                amount = number(value)
                if amount is None or amount < 0:
                    raise ValueError(f"{location}: Geçersiz tolerans")
            elif code == "onemlilik_tutari":
                amount = number(value)
                if amount is None or amount < 0:
                    raise ValueError(f"{location}: Önemlilik tutarı sıfır veya pozitif TL olmalı")
            elif code == "onemlilik_orani":
                amount = number(value)
                if amount is None or not 0 <= amount <= 1:
                    raise ValueError(
                        f"{location}: Önemlilik oranı 0 ile 1 arasında olmalı (örnek: 0,10 = %10)"
                    )
            elif code in {"ozet_kalem_sayisi", "arsiv_saklama_adedi"}:
                low = 1 if code == "ozet_kalem_sayisi" else 0
                if index(value, location) < low or value > 200:
                    raise ValueError(f"{location}: {low} ile 200 arasında tam sayı girin")
            elif section != "Dosyalar" and code != "sayfa":
                if index(value, location) == 0:
                    raise ValueError(f"{location}: Konum 1 veya daha büyük olmalı")
            elif not isinstance(value, str):
                raise ValueError(f"{location}: Dosya, klasör veya sayfa adı metin olmalı")
            settings[section][code] = value
        for section, (positioned, codes) in legacy.items():
            alternative = {"Endeks": "endeks_araligi"}.get(section)
            if positioned not in settings[section] and alternative not in settings[section]:
                sections[section] = sections[section] | codes
        for section, expected in sections.items():
            missing = expected - settings[section].keys()
            if missing:
                raise ValueError(f"{section}: Eksik ayarlar: {', '.join(sorted(missing))}")

        excluded = [str(values[0]) for _, values in rows("Ortak Sayfalar", 1)]
        departments = {}
        used_names = {key(name) for name in excluded}
        for i, (name, manager, director, index_name, location, control_name, submanager) in rows(
            "Bölüm Eşleştirme", 7, required=3
        ):
            if not all(isinstance(v, str) for v in (name, manager, director)) or key(name) in used_names:
                raise ValueError(f"Bölüm Eşleştirme, satır {i}: Geçersiz, mükerrer veya hariç tutulan bölüm")
            if index_name not in (None, "") and not isinstance(index_name, str):
                raise ValueError(f"Bölüm Eşleştirme, satır {i}: Endeks adı metin olmalı")
            if location not in (None, "") and not isinstance(location, str):
                raise ValueError(f"Bölüm Eşleştirme, satır {i}: Lokasyon metin olmalı")
            if control_name not in (None, "") and not isinstance(control_name, str):
                raise ValueError(f"Bölüm Eşleştirme, satır {i}: Kontrol bölümü metin olmalı")
            if submanager not in (None, "") and not isinstance(submanager, str):
                raise ValueError(f"Bölüm Eşleştirme, satır {i}: Müdürlük altı metin olmalı")
            used_names.add(key(name))
            departments[name] = {
                "mudurluk": manager,
                "direktorluk": director,
                "endeks": index_name or "",
                "lokasyon": location or "Tanımsız",
                "kontrol_bolumu": control_name or name,
                "mudurluk_alti": submanager or "Tanımsız",
            }
        if not departments:
            raise ValueError("Bölüm Eşleştirme: En az bir bölüm girin")

        groups = []
        for i, (low, high, name) in rows("Gider Grupları", 3):
            location = f"Gider Grupları, satır {i}"
            low, high = index(low, location), index(high, location)
            if low > high or not isinstance(name, str) or any(low <= b and a <= high for a, b, _ in groups):
                raise ValueError(f"{location}: Geçersiz veya çakışan gider aralığı")
            groups.append((low, high, name))
        expense_map = {}
        if "Gider Eşleştirme" in book:
            for i, (source, target, name) in rows("Gider Eşleştirme", 3):
                if source in expense_map or not isinstance(name, str):
                    raise ValueError(f"Gider Eşleştirme, satır {i}: Geçersiz veya mükerrer kaynak")
                expense_map[index(source, f"Gider Eşleştirme, satır {i}")] = (
                    index(target, f"Gider Eşleştirme, satır {i}"),
                    name,
                )
        return {
            **settings["Dosyalar"],
            "rapor": settings["Rapor"],
            "maliyet": settings["Maliyet"],
            "kontrol": settings["Kontrol"],
            "endeks": settings["Endeks"],
            "ortak_sayfalar": excluded,
            "bolumler": departments,
            "gider_gruplari": groups,
            "gider_eslestirme": expense_map,
        }
    finally:
        book.close()


def run(args):
    config = read_settings(args.ayar)
    root = args.ayar.parent
    report_period = period(args.donem)
    closing = period(args.kapanis) if args.kapanis else report_period
    if report_period > closing:
        raise ValueError("Rapor dönemi kapanıştan sonra olamaz")
    year, selected = report_period
    months = selected_months(getattr(args, "aylar", None), selected)
    if (year, months[-1]) > closing:
        raise ValueError("Seçilen aylar kaynak kapanış döneminden sonra olamaz")
    selection_id = month_signature(months)
    filename_months = "-".join(f"{m:02d}" for m in months)
    warnings = []
    if year != closing[0]:
        warnings.append(
            [
                label(closing),
                None,
                None,
                f"{year} bütçesi bu kapanış dosyasında yok. Bütçe karşılaştırması boş bırakıldı.",
            ]
        )
    inputs = [args.ayar, root / config["kontrol_dosyasi"], root / config["endeks_dosyasi"]]
    cost_dir = (root / config["maliyet_klasoru"]).resolve()
    if not cost_dir.is_dir():
        raise ValueError("Maliyet klasörü bulunamadı")
    if args.cikti.resolve().is_relative_to(cost_dir):
        raise ValueError("Rapor çıktısını maliyet klasörünün dışında seçin")
    inputs.extend(cost_dir.glob("*.xlsx"))
    if args.cikti.resolve() in {p.resolve() for p in inputs}:
        raise ValueError("Çıktı yolu bir kaynak dosyayla aynı olamaz")
    snapshots = read_snapshots(root / config["maliyet_klasoru"], config, closing, warnings)
    company, controls, company_budget = read_control(
        root / config["kontrol_dosyasi"],
        config["kontrol"],
        closing[0],
        {key(mapping["direktorluk"]) for mapping in config["bolumler"].values()},
    )
    indices = read_indices(root / config["endeks_dosyasi"], config["endeks"])
    detail, checks, missing = make_report(
        config, closing, report_period, months, snapshots, company, controls, indices, warnings
    )
    revenue = total(company.get((year, m)) for m in months)
    revenue_rows = [[label((year, m)), company.get((year, m)), company_budget.get((year, m))] for m in months]
    if (
        config["kontrol"].get("ciro_baslangici")
        and year == closing[0]
        and any(r[2] is None for r in revenue_rows)
    ):
        warnings.append([label(report_period), None, None, "Seçilen ayların bütçe cirosu eksik"])
    operating = {
        name: total(snapshots[(closing, name)][1].get((year, m)) for m in months)
        for name in config["bolumler"]
        if (closing, name) in snapshots
    }
    year_span = month_caption(months, year)
    previous_span = month_caption(months, year - 1)
    if revenue is None:
        warnings.append([label(closing), None, None, "Seçilen dönem için şirket cirosu eksik"])
    departments = config["bolumler"]
    layout = list(hierarchy(departments))
    from birim import read_budget_counts

    budget_counts = read_budget_counts(root, config, snapshots, closing)
    units = unit_cost_rows(
        config, closing, report_period, months, snapshots, indices, warnings, budget_counts
    )
    revision_rows = list(revisions(snapshots, config, closing))
    reports = [(args.cikti, departments, layout, units)]
    seq = 0
    for managers in department_tree(departments).values():
        for manager, names in managers.items():
            seq += 1
            safe = re.sub(r"[^\w -]", "_", manager).strip()[:65] or "Mudurluk"
            path = args.cikti.with_name(
                f"Mudurluk_{seq:02d}_{safe}_{label(report_period)}_Aylar_{filename_months}.xlsx"
            )
            selected_names = set(names)
            selected_layout = [
                (name, level, group)
                for name, level, group in layout
                if level in {"Bölüm", "Müdürlük toplamı"} and set(group) <= selected_names
            ]
            selected_units = [
                row
                for (_, level, group), row in zip(layout, units)
                if level in {"Bölüm", "Müdürlük toplamı"} and set(group) <= selected_names
            ]
            reports.append(
                (path, {name: departments[name] for name in names}, selected_layout, selected_units)
            )
    protected = {p.resolve() for p in inputs}
    for path, *_ in reports:
        if path.resolve() in protected or path.resolve().is_relative_to(cost_dir):
            raise ValueError(f"Çıktı bir kaynak dosyanın üzerine yazılamaz: {path}")
    if len({path.resolve() for path, *_ in reports}) != len(reports):
        raise ValueError("Genel ve müdürlük dosya adları çakışıyor")
    import analiz

    rules = analiz.settings(config)
    budget_year = year == closing[0]
    for file_index, (path, scope, report_layout, unit_rows) in enumerate(reports):
        names = set(scope)
        scoped_detail = [r for r in detail if r[0] in names]
        scoped_checks = [r for r in checks if set(r[-1]) <= names]
        scope_warnings = [r for r in warnings if r[1] is None or r[1] in names or r[1] not in departments]
        cross_scope = {r[0] for r in checks if set(r[-1]) & names and not set(r[-1]) <= names}
        for target in sorted(cross_scope):
            scope_warnings.append(
                [
                    label(report_period),
                    None,
                    None,
                    f"{target}: Birleşik bölüm kontrolü direktörlük raporunda yer alır",
                ]
            )
        write_report(
            path,
            scope,
            report_layout,
            scoped_detail,
            set(missing),
            scoped_checks,
            operating,
            revenue,
            revenue_rows,
            year_span,
            previous_span,
            unit_rows,
            [r for r in revision_rows if r[3] in names],
            scope_warnings,
            (label(report_period), label(closing), selection_id),
            editable=file_index != 0,
            snapshots=snapshots,
            closing=closing,
            analysis={
                "rules": rules,
                "budgetYear": budget_year,
                "closing": closing,
                "caption": year_span,
                "detail": scoped_detail,
                "checks": scoped_checks,
                "warnings": scope_warnings,
                "missing": sorted(set(missing) & names),
                "material": analiz.material_rows(scoped_detail, rules) if budget_year else [],
                "forecast": (
                    analiz.forecast_rows(config, snapshots, closing, report_layout) if budget_year else []
                ),
                "volume": (
                    analiz.volume_price_rows(config, snapshots, closing, months, budget_counts, report_layout)
                    if budget_year
                    else []
                ),
            },
        )
    return [path.resolve() for path, *_ in reports]


def add_selection_view(book, departments, snapshots, closing, report_id, *, sheet_name="Gider Seçim Paneli"):
    selectors = book.create_sheet("_Panel Dönem")
    scope = book.create_sheet("_Panel Kapsam")
    raw = book.create_sheet("_Panel Kaynak")
    view = book.create_sheet(sheet_name)
    select_ref = "'_Panel Dönem'"
    raw_ref = "'_Panel Kaynak'"
    scope_ref = "'_Panel Kapsam'"
    panel_ref = f"'{sheet_name}'"
    source = []
    for name in departments:
        snapshot = snapshots.get((closing, name))
        if snapshot is None:
            continue
        items, _, _, _ = snapshot
        for code, (item, actuals, budget) in items.items():
            for (year, month_number), value in actuals.items():
                if (year, month_number) <= closing:
                    source.append(
                        (
                            name,
                            (year, month_number),
                            code,
                            item,
                            value,
                            actuals.get((year - 1, month_number)),
                            budget.get((year, month_number)),
                        )
                    )
    periods = sorted({row[1] for row in source} | {period(report_id[0])})
    period_rows = {p: i for i, p in enumerate(periods, 2)}
    names = [name for name, level, _ in hierarchy(departments) if level == "Bölüm"]
    scope_rows = {name: i for i, name in enumerate(names, 2)}
    selectors.append(["Yıl", "Ay no", "Dahil", None, "Ay seçenekleri", "Seçili ay no", "Yıl seçenekleri"])
    scope.append(
        [
            "Müdürlük",
            "Bölüm",
            "Lokasyon",
            "Dahil",
            None,
            "Müdürlük seçenekleri",
            "Bölüm seçenekleri",
            "Lokasyon seçenekleri",
        ]
    )
    for r, m in enumerate(MONTHS, 2):
        selectors.cell(r, 5, m)
        selectors.cell(r, 6, f"=IF(COUNTIF({panel_ref}!$B$6:$G$7,E{r})>0,{r - 1},0)")
    years = sorted({y for y, _ in periods})
    for r, y in enumerate(years, 2):
        selectors.cell(r, 7, y)
    for r, (year, m) in enumerate(periods, 2):
        selectors.cell(r, 1, year)
        selectors.cell(r, 2, m)
        selectors.cell(r, 3, f"=--AND(A{r}={panel_ref}!$B$3,F{m + 1}>0)")
    for r, name in enumerate(names, 2):
        mapping = departments[name]
        for c, value in enumerate((mapping["mudurluk"], name, mapping.get("lokasyon") or "Tanımsız"), 1):
            scope.cell(r, c, value).data_type = "s"
        scope.cell(
            r,
            4,
            f'=--AND(OR({panel_ref}!$F$3="Tümü",A{r}={panel_ref}!$F$3),'
            f'OR({panel_ref}!$B$4="Tümü",B{r}={panel_ref}!$B$4),'
            f'OR({panel_ref}!$D$4="Tümü",C{r}={panel_ref}!$D$4))',
        )

    def add_choice(name, sheet, cells):
        definition = DefinedName(name, attr_text=f"'{sheet.title}'!{cells}")
        if hasattr(book.defined_names, "add"):
            book.defined_names.add(definition)
        else:
            book.defined_names.append(definition)
        return name

    option_ranges = {}
    for field, col, options in (
        ("Müdürlük", "F", {m["mudurluk"] for m in departments.values()}),
        ("Bölüm", "G", set(names)),
        ("Lokasyon", "H", {m.get("lokasyon") or "Tanımsız" for m in departments.values()}),
    ):
        values = ["Tümü", *sorted(options - {"Tümü"}, key=key)]
        for r, value in enumerate(values, 2):
            scope[f"{col}{r}"] = value
            scope[f"{col}{r}"].data_type = "s"
        option_ranges[field] = add_choice(f"PanelKapsam{col}", scope, f"${col}$2:${col}${len(values) + 1}")

    raw.append(
        [
            "İndeks",
            "Gider",
            "Gerçekleşen",
            "Önceki yıl",
            "Bütçe",
            "Dahil",
            "Gerçekleşen mevcut",
            "Önceki yıl mevcut",
            "Bütçe mevcut",
        ]
    )
    r = raw.max_row  # max_row her çağrıda tüm hücreleri tarar; döngüde sayaç kullanılır.
    for name, p, code, item, actual, previous, budget in source:
        r += 1
        raw.append(
            [
                code,
                item,
                actual,
                previous,
                budget,
                f"={select_ref}!C{period_rows[p]}*{scope_ref}!D{scope_rows[name]}",
                f"=IF(ISNUMBER(C{r}),F{r},0)",
                f"=IF(ISNUMBER(D{r}),F{r},0)",
                f"=IF(ISNUMBER(E{r}),F{r},0)",
            ]
        )
        raw.cell(r, 2).data_type = "s"
    last = max(raw.max_row, 2)

    view.merge_cells("A1:G1")
    view["A1"] = "Bütçe Dashboard" if sheet_name == "Dashboard" else sheet_name
    view["A1"].font = Font(size=16, bold=True, color="263746")
    view.merge_cells("A2:G2")
    view["A2"] = "Seçtiğin aylar gerçekleşen, bütçe ve önceki yıl için aynı şekilde toplanır."
    y, m = period(report_id[0])
    for address, value in {
        "A3": "Yıl",
        "B3": y,
        "E3": "Müdürlük",
        "F3": "Tümü",
        "A4": "Bölüm",
        "B4": "Tümü",
        "C4": "Lokasyon",
        "D4": "Tümü",
        "A6": "Aylar",
    }.items():
        view[address] = value
    initial_months = selected_months(report_id[2].removeprefix("aylar:"), m)
    for i, month_number in enumerate(initial_months):
        view.cell(6 + i // 6, 2 + i % 6, MONTHS[month_number - 1])
    view.merge_cells("A5:G5")
    view["A5"] = "Bir veya daha fazla ay seç. Bir ayı kaldırmak için hücreyi Delete ile temizle."
    view.merge_cells("A8:G8")
    view["A8"] = (
        f'=IF(COUNTIF({select_ref}!F2:F13,">0")=0,"En az bir ay seçin.",'
        f'IF(SUM({scope_ref}!D2:D{len(names) + 1})=0,"Kapsam seçimleri eşleşmiyor.",'
        f'IF(SUM({select_ref}!C2:C{len(periods) + 1})<>COUNTIF({select_ref}!F2:F13,">0"),'
        '"Seçilen aylardan en az biri kaynakta yok. Seçimi kontrol edin.","")))'
    )
    validations = [
        ("B3", add_choice("PanelYillar", selectors, f"$G$2:$G${len(years) + 1}"), False),
        ("F3", option_ranges["Müdürlük"], False),
        ("B4", option_ranges["Bölüm"], False),
        ("D4", option_ranges["Lokasyon"], False),
        ("B6:G7", add_choice("PanelAylar", selectors, "$E$2:$E$13"), True),
    ]
    for address, formula, blank in validations:
        rule = DataValidation(
            type="list",
            formula1=formula,
            allow_blank=blank,
            showErrorMessage=True,
            errorTitle="Geçersiz seçim",
            error="Açılır listeden bir değer seçin.",
        )
        view.add_data_validation(rule)
        rule.add(address)
    for row in (view[3], view[4], view[6], view[7]):
        for cell in row:
            if cell.coordinate in {"B3", "F3", "B4", "D4"} or cell.row in {6, 7} and 2 <= cell.column <= 7:
                cell.fill = PatternFill("solid", fgColor="E3F2F4")
    book.active = view
    view.freeze_panes = "C11"
    view.sheet_view.showGridLines = False
    headings = [
        "İndeks",
        "Alt gider",
        "Gerçekleşen (TL)",
        "Önceki yıl (TL)",
        "Bütçe (TL)",
        "Bütçe − gerçekleşen (TL)",
        "Bütçe aşım oranı (%)",
    ]
    for c, title in enumerate(headings, 1):
        cell = view.cell(10, c, title)
        cell.fill = PatternFill("solid", fgColor="E7EEF2")
        cell.font = Font(bold=True, color="263746")
    for c, width in enumerate((12, 32, 20, 20, 20, 25, 24), 1):
        view.column_dimensions[get_column_letter(c)].width = width
    for r, (code, item) in enumerate([*sorted({(row[2], row[3]) for row in source}), (None, "Toplam")], 11):
        if code is not None:
            view.cell(r, 1, code)
        view.cell(r, 2, item).data_type = "s"

        def summarize(column):
            area = f"{raw_ref}!${column}$2:${column}${last}"
            if code is None:
                return f"SUM({area})"
            return f"SUMIFS({area},{raw_ref}!$A$2:$A${last},A{r},{raw_ref}!$B$2:$B${last},B{r})"

        for c, data_col, presence_col in ((3, "C", "G"), (4, "D", "H"), (5, "E", "I")):
            source_range = f"{raw_ref}!${data_col}$2:${data_col}${last}"
            active_range = f"{raw_ref}!$F$2:$F${last}"
            criteria = (
                ""
                if code is None
                else f",--({raw_ref}!$A$2:$A${last}=A{r}),--({raw_ref}!$B$2:$B${last}=B{r})"
            )
            active, present = summarize("F"), summarize(presence_col)
            view.cell(
                r,
                c,
                f'=IF(OR($A$8<>"",{active}=0,{present}<>{active}),"",SUMPRODUCT({source_range},{active_range}{criteria}))',
            )
        view.cell(r, 6, f'=IF(COUNT(C{r},E{r})<2,"",E{r}-C{r})')
        view.cell(r, 7, f'=IF(OR(COUNT(C{r},E{r})<2,E{r}=0),"",(C{r}-E{r})/E{r})')
        for c in range(3, 8):
            view.cell(r, c).number_format = "0.00%" if c == 7 else "#,##0.00;-#,##0.00;0.00"
        if code is None:
            for cell in view[r]:
                cell.font = Font(bold=True, color="263746")
    for col, positive, negative in (("F", "C6EFCE", "FFC7CE"), ("G", "FFC7CE", "C6EFCE")):
        for operator, fill in (("greaterThan", positive), ("lessThan", negative)):
            view.conditional_formatting.add(
                f"{col}11:{col}{view.max_row}",
                CellIsRule(operator=operator, formula=["0"], fill=PatternFill("solid", fgColor=fill)),
            )
    selectors.sheet_state = scope.sheet_state = raw.sheet_state = "veryHidden"
    if "Giderler" in book:
        main = book["Giderler"]
        main["A3"] = "Yıl, ay ve kapsam seçim panelini aç"
        main["A3"].hyperlink = Hyperlink(ref="A3", location=f"'{sheet_name}'!A1")
        main["A3"].font = Font(color="245C82", underline="single")
    return view


def write_report(
    path,
    departments,
    layout,
    detail,
    missing,
    checks,
    operating,
    revenue,
    revenue_rows,
    year_span,
    previous_span,
    unit_rows,
    revision_rows,
    warnings,
    report_id,
    editable,
    snapshots,
    closing,
    analysis=None,
):
    book = Workbook()
    book.remove(book.active)

    def title(name, span=year_span):
        return f"{name}\n{span}"

    comparison_span = f"{year_span} / {previous_span}"
    metrics = [
        title("Önceki yıl gerçekleşen (TL)", previous_span),
        title("Gerçekleşen (TL)"),
        title("Bütçe (TL)"),
        title("Bütçe − gerçekleşen (TL)"),
        title("Bütçe aşım oranı (%)"),
    ]
    add_general_summary(
        book, detail, layout, set(missing), metrics, operating, revenue, title("OPEX / şirket cirosu")
    )
    add_expense_matrix(book, detail, departments, year_span, editable)
    add_selection_view(book, departments, snapshots, closing, report_id)
    unit_rows = [r for r in unit_rows if r[4] is not None]
    if unit_rows:
        add_table(
            book,
            "Birim Maliyet",
            [
                "Bölüm / müdürlük / direktörlük",
                "Müdürlük / satır türü",
                title("Toplam gider (TL)"),
                title("İşlem adedi"),
                title("Birim maliyet (TL/işlem)"),
                title("Toplam gider (TL)", previous_span),
                title("İşlem adedi", previous_span),
                title("Birim maliyet (TL/işlem)", previous_span),
                title("Endekslenmiş birim maliyet (TL/işlem)", previous_span),
                title("Birim maliyet iyileşmesi (%)", comparison_span),
                title("Dönem endeksi", comparison_span),
                title("Müdürlük gider ağırlığı"),
                title("Direktörlük gider ağırlığı"),
                title("Bütçe BM (TL/işlem)"),
            ],
            [row[:12] + [row[13], row[15]] for row in unit_rows],
            (10, 12, 13),
        )
        if not str(report_id[0]).startswith(str(closing[0])):
            sheet = book["Birim Maliyet"]
            sheet.delete_cols(6, sheet.max_column - 5)
            sheet.auto_filter.ref = sheet.dimensions
        style_subtotals(
            book["Birim Maliyet"],
            [
                (r, row[1])
                for r, row in enumerate(unit_rows, 2)
                if row[1] in {"Müdürlük toplamı", "Direktörlük toplamı", "Genel toplam"}
            ],
        )
    if analysis:
        import analiz

        analiz.add_sheets(book, analysis, editable)
    add_audit(book, checks, revision_rows, warnings, revenue_rows)
    metadata = book.create_sheet("_Rapor Kimliği")
    metadata.append(["ButceSapmasi-1", "Müdürlük" if editable else "Direktörlük", *report_id])
    for name in (
        n for managers in department_tree(departments).values() for group in managers.values() for n in group
    ):
        metadata.append([name, departments[name]["mudurluk"], departments[name]["direktorluk"]])
    metadata.sheet_state = "veryHidden"
    save_report(book, path)


def save_report(book, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as handle:
        temp_path = Path(handle.name)
    try:
        book.save(temp_path)
        temp_path.replace(path)
    finally:
        temp_path.unlink(missing_ok=True)


def explanation_layout(book, kind):
    sheet = book["_Rapor Kimliği"]
    signature = tuple(cell.value for cell in sheet[1])
    if len(signature) < 5 or signature[:2] != ("ButceSapmasi-1", kind):
        raise ValueError(f"{kind} türünde güncel rapor dosyası seçin")
    departments = {}
    for name, manager, director in sheet.iter_rows(min_row=2, max_col=3, values_only=True):
        if name in departments or not all(isinstance(v, str) and v for v in (name, manager, director)):
            raise ValueError("Rapor bölüm kimlikleri eksik veya mükerrer")
        departments[name] = (manager, director)
    expense = book["Giderler"]
    step = 5 if kind == "Müdürlük" else 4
    columns = {expense.cell(4, col).value: col for col in range(4, expense.max_column + 1, step)}
    if set(columns) != set(departments) or len(columns) != len(departments):
        raise ValueError("Giderler sayfasındaki bölüm başlıkları rapor kimliğiyle eşleşmiyor")
    rows = {}
    for row in range(6, expense.max_row + 1):
        index = expense.cell(row, 2).value
        if not isinstance(index, int):
            continue
        item = (index, expense.cell(row, 1).value, expense.cell(row, 3).value)
        if item in rows:
            raise ValueError(f"Mükerrer alt gider: {item[2]}")
        rows[item] = row
    return signature[2:5], departments, columns, rows


def merge_explanations(general_path, manager_paths, output_path):
    manager_paths = list(manager_paths)
    selected = [general_path, *manager_paths]
    if output_path.resolve() in {p.resolve() for p in selected}:
        raise ValueError("Birleşik raporu yeni bir dosya adıyla kaydedin")
    general = load_workbook(general_path)
    try:
        period_id, departments, columns, items = explanation_layout(general, "Direktörlük")
        recorded = set()
        notes = {}
        target_sheet = general["Giderler"]
        for path in manager_paths:
            try:
                book = load_workbook(path)
                try:
                    source_id, members, source_cols, source_items = explanation_layout(book, "Müdürlük")
                    if source_id != period_id:
                        raise ValueError("Rapor yılı, kapanışı veya seçili aylar farklı")
                    if any(departments.get(name) != pair for name, pair in members.items()):
                        raise ValueError("Bölüm, müdürlük veya direktörlük eşleşmiyor")
                    expected = set(members.values())
                    if len(expected) != 1 or set(members) != {
                        n for n, pair in departments.items() if pair in expected
                    }:
                        raise ValueError("Müdürlük kapsamı eksik veya geçersiz")
                    if recorded.intersection(members):
                        raise ValueError("Aynı bölüm birden fazla dosyada seçildi")
                    recorded.update(members)
                    source = book["Giderler"]
                    for name, col in source_cols.items():
                        if source.cell(5, col + 4).value != "Açıklama":
                            raise ValueError("Açıklama sütunu bulunamadı")
                        for item, row in source_items.items():
                            note = source.cell(row, col + 4)
                            value = str(note.value).strip() if note.value is not None else ""
                            if not value:
                                continue
                            if note.data_type == "f":
                                raise ValueError(f"Açıklama metin olmalı: {name}, {item[2]}")
                            target_row = items.get(item)
                            if target_row is None or any(
                                source.cell(row, col + offset).value
                                != target_sheet.cell(target_row, columns[name] + offset).value
                                for offset in (0, 1)
                            ):
                                raise ValueError(f"Alt gider veya rapor tutarı uyuşmuyor: {name}, {item[2]}")
                            notes.setdefault(item, {})[name] = value
                finally:
                    book.close()
            except (KeyError, ValueError) as exc:
                raise ValueError(f"{path.name}: {exc}") from exc
        if missing := departments.keys() - recorded:
            raise ValueError("Müdürlük dosyası eksik. Kapsanmayan bölümler: " + ", ".join(sorted(missing)))

        result = []
        for index, group, title in items:
            entries = notes.get((index, group, title), {})
            if entries:
                combined = "\n".join(f"{name}: {entries[name]}" for name in departments if name in entries)
                if len(combined) > 32767:
                    raise ValueError(f"Birleşik açıklama Excel hücre sınırını aşıyor: {title}")
                result.append([group, index, title, combined])
        sheet_name = "Alt Gider Açıklamaları"
        if sheet_name in general:
            del general[sheet_name]
        add_table(
            general,
            sheet_name,
            ["Gider grubu", "İndeks", "Alt gider kalemi", "Birleşik bölüm açıklaması"],
            result,
        )
        sheet = general[sheet_name]
        sheet.column_dimensions["D"].width = 78
        for r in range(2, sheet.max_row + 1):
            cell = sheet.cell(r, 4)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            sheet.row_dimensions[r].height = min(
                400, 16 * sum(max(1, (len(line) + 69) // 70) for line in cell.value.split("\n"))
            )
        save_report(general, output_path)
        return len(result)
    finally:
        general.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Bütçe sapması Excel raporu")
    parser.add_argument("--ayar", type=Path, required=True, help="Ayarlar.xlsx dosyası")
    parser.add_argument("--donem", required=True, help="YYYY-AA")
    parser.add_argument("--kapanis", help="Kaynak kapanışı YYYY-AA. Varsayılan: rapor dönemi.")
    parser.add_argument("--aylar", help="Hesaplanacak aylar. Örnek: 2,4,6. Varsayılan: --donem ayı.")
    parser.add_argument("--cikti", type=Path, required=True)
    try:
        args = parser.parse_args()
        paths = run(args)
        print(f"{len(paths)} rapor kaydedildi: {paths[0].parent}")
    except (OSError, ValueError, KeyError, TypeError, RuntimeError) as exc:
        sys.exit(f"Rapor üretilemedi: {exc}")
