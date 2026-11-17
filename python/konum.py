"""Kaynak dosyalarda konumların otomatik bulunması (V5.5).

Ayarlar › Dosya ve Konum Ayarları'ndaki konum satırları artık isteğe bağlıdır:
- Satır varsa o değer kullanılır (öncelik Ayarlar'dadır).
- Satır yoksa konum dosyadan bulunur ve Uyarılar'da "Konum otomatik bulundu" diye yazılır.
- Ayardaki konum bir maliyet sayfasında tutmazsa (satır/sütun kaymışsa) o sayfa için konum yeniden bulunur
  ve uyarılır (butce.read_cost).

Kurallar (makro: MakroKonum, Power Query: Butce_Veri.pq › Konum):
Maliyet  "Bütçe" yazan hücrenin hemen altı tarih/ay ise o satır başlık satırıdır. Başlık satırında "Bütçe"
         etiketinin solundaki ilk tarih gerçekleşen başlangıcıdır (tarih_baslangici). Bunun solunda, altında en
         az üç yazı olan en yakın sütun gider adı sütunudur (gider_sutunu). Bu sütunda "İşlem Adedi" yazan satır
         işlem adedi satırı, bir altı ilk gider satırıdır; gider kalemleri ilk boş ada kadar sayılır.
Kontrol  Aynı satırda en az 12 tarih bulunan ve altında sayı olan en üst satır ciro başlığıdır
         (ciro_baslangici = ilk tarih). Tarihler kesildikten sonra gelen ilk tarih bütçe cirosu başlangıcıdır.
         "Opex Toplamı" yazan sütun etiket sütunudur.
Endeks   "Bölüm Gider Endeksi" veya "Bölüm Adı" yazan ve sağında tarih olan hücre endeks tablosunun sol üstüdür.
"""

from datetime import date
import re

from openpyxl.utils import get_column_letter

MAX_ROWS, MAX_COLS = 400, 1500
BUDGET = {"bütçe", "butce"}
INDEX_TITLES = {"bölüm adı", "bölüm gider endeksi"}
MONTHS = "ocak şubat mart nisan mayıs haziran temmuz ağustos eylül ekim kasım aralık".split()

COST_KEYS = ("tarih_baslangici", "butce_etiketi", "gider_sutunu", "islem_adedi_satiri", "ilk_gider_satiri",
             "gider_kalem_sayisi")
COST_LEGACY = ("yil_satiri", "ay_satiri", "dip_toplam_satiri", "indeks_sutunu", "ilk_donem_sutunu",
               "gerceklesen_araligi", "butce_araligi")
CONTROL_KEYS = ("ciro_baslangici", "butce_ciro_baslangici", "etiket_sutunu")
CONTROL_LEGACY = ("ciro_satiri", "ilk_blok_satiri", "ilk_ay_sutunu", "ciro_araligi")
INDEX_LEGACY = ("ilk_veri_satiri",)


def key(value):
    return str(value or "").strip().translate(str.maketrans("İI", "iı")).casefold()


def is_date(value):
    """Tarih hücresi: gerçek tarih, '2026-01' veya '01.01.2026'."""
    if isinstance(value, date):
        return True
    text = str(value or "").strip()
    return bool(re.fullmatch(r"\d{4}-\d{1,2}(-\d{1,2})?", text) or re.fullmatch(r"\d{1,2}\.\d{1,2}\.\d{4}", text))


def is_month(value):
    if is_date(value):
        return True
    if isinstance(value, int) and not isinstance(value, bool):
        return 1 <= value <= 12
    k = key(value)
    return any(k in (m, m[:3]) for m in MONTHS)


def is_text(value):
    return isinstance(value, str) and bool(value.strip())


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def grid(sheet):
    """Dolu hücreler: {(satır, sütun): değer}. İlk 400 satır ve 1500 sütun taranır."""
    cells = {}
    for r, row in enumerate(sheet.iter_rows(values_only=True), 1):
        if r > MAX_ROWS:
            break
        for c, value in enumerate(row[:MAX_COLS], 1):
            if value is not None and value != "":
                cells[(r, c)] = value
    return cells


def cell(col, row):
    return f"{get_column_letter(col)}{row}"


def cost_layout(sheet, cells=None):
    """Maliyet sayfası konumları; bulunamayan anahtar sonuçta yer almaz. Belirsizse {}."""
    cells = grid(sheet) if cells is None else cells
    markers = sorted((r, c) for (r, c), v in cells.items() if key(v) in BUDGET and is_month(cells.get((r + 1, c))))
    if not markers or len({r for r, _ in markers}) != 1:
        return {}
    marker_row, marker_col = markers[0]
    header = marker_row + 1
    dates = sorted(c for (r, c), v in cells.items() if r == header and c < marker_col and is_date(v))
    if not dates:
        return {}
    start = dates[0]
    found = {"tarih_baslangici": cell(start, header), "butce_etiketi": cell(marker_col, marker_row)}
    name_col = next((c for c in range(start - 1, 0, -1)
                     if sum(is_text(cells.get((r, c))) for r in range(header + 1, header + 60)) >= 3), None)
    if name_col is None:
        return found
    found["gider_sutunu"] = name_col
    count_row = next((r for r in range(header + 1, header + 60)
                      if "işlem aded" in key(cells.get((r, name_col))) or "islem aded" in key(cells.get((r, name_col)))),
                     None)
    if count_row is None:
        return found
    found["islem_adedi_satiri"] = count_row
    found["ilk_gider_satiri"] = count_row + 1
    size = 0
    while is_text(cells.get((count_row + 1 + size, name_col))):
        size += 1
    if size:
        found["gider_kalem_sayisi"] = size
    return found


def control_layout(sheet, cells=None):
    cells = grid(sheet) if cells is None else cells
    rows = {}
    for (r, c), v in cells.items():
        if is_date(v):
            rows.setdefault(r, []).append(c)
    header = next((r for r in sorted(rows)
                   if len(rows[r]) >= 12 and is_number(cells.get((r + 1, min(rows[r]))))), None)
    if header is None:
        return {}
    cols = sorted(rows[header])
    found = {"ciro_baslangici": cell(cols[0], header)}
    gap = next((b for a, b in zip(cols, cols[1:]) if b != a + 1), None)
    if gap is not None:
        found["butce_ciro_baslangici"] = cell(gap, header)
    labels = {c for (r, c), v in cells.items() if r > header + 1 and key(v) == "opex toplamı"}
    if len(labels) == 1:
        found["etiket_sutunu"] = labels.pop()
    return found


def index_layout(sheet, cells=None):
    cells = grid(sheet) if cells is None else cells
    corners = sorted((r, c) for (r, c), v in cells.items() if key(v) in INDEX_TITLES and is_date(cells.get((r, c + 1))))
    return {"endeks_araligi": cell(corners[0][1], corners[0][0])} if len(corners) == 1 else {}


def describe(found):
    return ", ".join(f"{k} = {v}" for k, v in found.items())


def legacy_mode(cfg, positioned, legacy):
    """Konumlu ayar yok, eski satır/sütun düzeninden ayar var: eski okuma yolu kullanılır, konum aranmaz."""
    return not cfg.get(positioned) and any(cfg.get(k) not in (None, "") for k in legacy)


def missing(cfg, keys, positioned, legacy):
    """Konumlu düzende (veya hiç konum/eski düzen ayarı yoksa) eksik anahtarlar."""
    if legacy_mode(cfg, positioned, legacy):
        return []
    return [k for k in keys if cfg.get(k) in (None, "")]


def complete(config, root):
    """Eksik konum ayarlarını dosyalardan bulur. config['otomatik_konum'] = [(bölüm, dosya, açıklama)]."""
    from openpyxl import load_workbook

    import butce

    notes = config.setdefault("otomatik_konum", [])
    cost = config["maliyet"]
    wanted = missing(cost, COST_KEYS, "tarih_baslangici", COST_LEGACY)
    # gider_kalem_sayisi yoksa dip toplam satırı da kullanılabilir (butce.item_count); yine de bulunur.
    if wanted:
        ignored = {key(s) for s in config["ortak_sayfalar"]}
        configured = {key(s) for s in config["bolumler"]}
        files = sorted(butce.cost_files(root / config["maliyet_klasoru"]),
                       key=lambda p: (_period_or_none(butce, p) or (0, 0), p.name), reverse=True)
        for path in files:
            book = load_workbook(path, data_only=True, read_only=True)
            try:
                sheet = next((s for s in book if key(s.title) in configured and key(s.title) not in ignored), None)
                found = cost_layout(sheet) if sheet is not None else {}
            finally:
                book.close()
            if "tarih_baslangici" in found:
                added = {k: found[k] for k in wanted if k in found}
                cost.update(added)
                if added:
                    notes.append(("Maliyet", f"{path.name} › {sheet.title}", describe(added)))
                break
    control = config["kontrol"]
    wanted = missing(control, CONTROL_KEYS, "ciro_baslangici", CONTROL_LEGACY)
    path = root / config["kontrol_dosyasi"]
    if wanted and path.is_file():
        book = load_workbook(path, data_only=True, read_only=True)
        try:
            sheet = book[control["sayfa"]] if control.get("sayfa") in book.sheetnames else book.worksheets[0]
            found = control_layout(sheet)
        finally:
            book.close()
        added = {k: found[k] for k in wanted if k in found}
        control.update(added)
        if added:
            notes.append(("Kontrol", path.name, describe(added)))
    index = config["endeks"]
    path = root / config["endeks_dosyasi"]
    if (not index.get("baslik_satiri") and not index.get("endeks_araligi") and not index.get("ilk_veri_satiri")
            and path.is_file()):
        book = load_workbook(path, data_only=True, read_only=True)
        try:
            sheet = book[index["sayfa"]] if index.get("sayfa") in book.sheetnames else book.worksheets[0]
            found = index_layout(sheet)
        finally:
            book.close()
        index.update(found)
        if found:
            notes.append(("Endeks", path.name, describe(found)))
    lacking = []
    if not legacy_mode(cost, "tarih_baslangici", COST_LEGACY):
        lacking += [f"Maliyet › {k}" for k in ("tarih_baslangici", "butce_etiketi", "islem_adedi_satiri",
                                                "ilk_gider_satiri") if cost.get(k) in (None, "")]
    if not legacy_mode(control, "ciro_baslangici", CONTROL_LEGACY):
        lacking += [f"Kontrol › {k}" for k in CONTROL_KEYS if control.get(k) in (None, "")]
    if not any(index.get(k) for k in ("baslik_satiri", "endeks_araligi", "ilk_veri_satiri")):
        lacking.append("Endeks › endeks_araligi")
    if lacking:
        raise ValueError("Konum dosyadan bulunamadı; Ayarlar › Dosya ve Konum Ayarları'na girin: " + ", ".join(lacking))
    return config


def _period_or_none(butce, path):
    try:
        return butce.filename_period(path)
    except ValueError:
        return None
