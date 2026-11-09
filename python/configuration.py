"""Excel settings locations: preview, validate, backup and atomically update.

Only three existing location cells change. Mapping sheets remain authoritative.
"""

from pathlib import Path
from tempfile import NamedTemporaryFile
from datetime import datetime
import hashlib, shutil, os
from openpyxl import load_workbook
import butce

KEYS = {"costFolder": "maliyet_klasoru", "controlFile": "kontrol_dosyasi", "indexFile": "endeks_dosyasi"}


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def preview(path):
    path = Path(path).expanduser().resolve()
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Ayar dosyası .xlsx olmalı.")
    cfg = butce.read_settings(path)
    locations = {key: str((path.parent / cfg[code]).resolve()) for key, code in KEYS.items()}
    return {
        "settings": str(path),
        "fingerprint": digest(path),
        "locations": locations,
        "mappings": [{"department": n, **m} for n, m in cfg["bolumler"].items()],
        "expenseGroups": cfg["gider_gruplari"],
        "expenseMappings": [
            [source, *mapping] for source, mapping in sorted(cfg.get("gider_eslestirme", {}).items())
        ],
        "excludedSheets": cfg["ortak_sayfalar"],
        "mappingCounts": {
            "departments": len(cfg["bolumler"]),
            "managers": len({m["mudurluk"] for m in cfg["bolumler"].values()}),
            "groups": len(cfg["gider_gruplari"]),
            "expenseMappings": len(cfg.get("gider_eslestirme", {})),
            "excludedSheets": len(cfg["ortak_sayfalar"]),
        },
    }


def update_locations(settings, expected_settings, fingerprint, locations, closing, factory):
    target = Path(settings).expanduser().resolve()
    if target != Path(expected_settings).expanduser().resolve():
        raise ValueError("Seçilen ayar dosyası değişti. Önce Ayarları oku düğmesine basın.")
    current = preview(target)
    if current["fingerprint"] != fingerprint:
        raise ValueError("Ayar Excel’i son okumadan sonra değişti. Ayarları yeniden okuyun.")
    if not isinstance(locations, dict) or set(locations) != set(KEYS):
        raise ValueError("Üç kaynak konumunu birlikte seçin.")
    resolved = {}
    for key, value in locations.items():
        if not isinstance(value, str) or not value.strip():
            raise ValueError("Kaynak yolu boş olamaz: " + key)
        path = Path(value).expanduser()
        if not path.is_absolute():
            path = target.parent / path
        path = path.resolve()
        if key == "costFolder":
            if not path.is_dir():
                raise ValueError("Maliyet klasörü bulunamadı: " + str(path))
        elif not path.is_file() or path.suffix.lower() != ".xlsx":
            raise ValueError("Kaynak Excel dosyası bulunamadı: " + str(path))
        resolved[key] = str(path)
    if not isinstance(closing, list) or len(closing) != 2 or any(type(v) is not int for v in closing):
        raise ValueError("Kapanış yılı ve ayı tam sayı olmalı.")
    book = load_workbook(target)
    temp = None
    try:
        sheet = book["Dosya ve Konum Ayarları"]
        seen = set()
        for row in sheet.iter_rows(min_row=5, min_col=1, max_col=3):
            section, code, value = [c.value for c in row]
            if section == "Dosyalar" and code in KEYS.values():
                if code in seen:
                    raise ValueError("Mükerrer konum ayarı: " + code)
                seen.add(code)
                key = next(k for k, v in KEYS.items() if v == code)
                row[2].value = resolved[key]
        if seen != set(KEYS.values()):
            raise ValueError("Ayar dosyasında kaynak konumları eksik.")
        with NamedTemporaryFile(dir=target.parent, prefix=".butce_ayar_", suffix=".xlsx", delete=False) as f:
            temp = Path(f.name)
        book.save(temp)
        book.close()
        # Candidate keeps the same settings root, so relative non-location metadata stays valid.
        candidate = factory(temp, closing)
        if digest(target) != fingerprint:
            raise ValueError("Ayar dosyası işlem sırasında değişti; üzerine yazılmadı.")
        folder = target.parent / "Ayar_Yedekleri"
        folder.mkdir(exist_ok=True)
        backup = folder / (target.stem + "_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".xlsx")
        shutil.copy2(target, backup)
        # Excel locking the file on Windows causes a clean failure; existing bytes stay intact.
        os.replace(temp, target)
        temp = None
        candidate.settings = target
        candidate.identity = hashlib.sha256(str(target).encode()).hexdigest()
        return candidate, str(backup)
    finally:
        book.close()
        if temp is not None:
            temp.unlink(missing_ok=True)


def reconcile_selection(state, old_options, new_options):
    """All stays all after additions; deleted explicit values cannot remain active."""
    import copy

    if not isinstance(state, dict):
        return None
    result = copy.deepcopy(state)
    filters = result.setdefault("filters", {})
    for key, new in new_options.items():
        values = filters.get(key)
        if not isinstance(values, list):
            continue
        old = old_options.get(key, [])
        if key not in ("year", "month") and old and set(values) == set(old):
            filters[key] = None
        else:
            keep = [v for v in values if v in new]
            filters[key] = keep if keep or not values else None
    for key in ("bmLocal", "opexLocal"):
        values = result.get(key)
        if isinstance(values, list):
            if old_options.get("department") and set(values) == set(old_options["department"]):
                result[key] = None
            else:
                keep = [v for v in values if v in new_options.get("department", [])]
                result[key] = keep if keep or not values else None
    return result


def layout_preview(settings, department=None):
    """Read a small source window around configured indicator/expense/date anchors."""
    path = Path(settings).resolve()
    cfg = butce.read_settings(path)
    rows = []
    book = load_workbook(path, data_only=False)
    try:
        for r in book["Dosya ve Konum Ayarları"].iter_rows(min_row=5, max_col=3):
            if r[0].value in ("Maliyet", "Kontrol", "Endeks"):
                rows.append([r[0].value, r[1].value, r[2].value])
    finally:
        book.close()
    sample = []
    source = ""
    sheet = ""
    for f in sorted((path.parent / cfg["maliyet_klasoru"]).glob("*.xlsx"), reverse=True):
        if f.name.startswith("~$"):
            continue
        book = load_workbook(f, data_only=True)
        try:
            name = department or next((n for n in cfg["bolumler"] if n in book.sheetnames), None)
            if name not in book.sheetnames:
                continue
            ws = book[name]
            first = cfg["maliyet"]["islem_adedi_satiri"] - 4
            from openpyxl.utils.cell import coordinate_to_tuple

            header, col = coordinate_to_tuple(cfg["maliyet"]["tarih_baslangici"])
            for r in [header, *range(first, first + 7)]:
                sample.append(
                    [
                        [ws.cell(r, c).coordinate, str(ws.cell(r, c).value or "")]
                        for c in [cfg["maliyet"]["gider_sutunu"], *range(col, col + 3)]
                    ]
                )
            source = str(f)
            sheet = name
            break
        finally:
            book.close()
    return {
        "settings": str(path),
        "fingerprint": digest(path),
        "parameters": rows,
        "sample": sample,
        "file": source,
        "sheet": sheet,
    }


def update_layout(settings, fingerprint, parameters, closing, factory):
    """Whitelist existing parameter cells; a full candidate load precedes replacement."""
    target = Path(settings).resolve()
    if digest(target) != fingerprint:
        raise ValueError("Ayar değişti; düzeni yeniden okuyun.")
    if not isinstance(parameters, list) or len(parameters) > 100:
        raise ValueError("Geçersiz düzen parametreleri.")
    book = load_workbook(target)
    temp = None
    try:
        cells = {
            (r[0].value, r[1].value): r[2]
            for r in book["Dosya ve Konum Ayarları"].iter_rows(min_row=5, max_col=3)
            if r[0].value in ("Maliyet", "Kontrol", "Endeks")
        }
        seen = set()
        for section, code, value in parameters:
            key = section, code
            if key not in cells or key in seen:
                raise ValueError("Tanımsız veya yinelenen düzen parametresi.")
            if isinstance(value, str) and value.startswith("="):
                raise ValueError("Parametre formül olamaz.")
            seen.add(key)
            cells[key].value = value
        with NamedTemporaryFile(dir=target.parent, suffix=".xlsx", delete=False) as f:
            temp = Path(f.name)
        book.save(temp)
        candidate = factory(temp, closing)
        if digest(target) != fingerprint:
            raise ValueError("Ayar işlem sırasında değişti.")
        folder = target.parent / "Ayar_Yedekleri"
        folder.mkdir(exist_ok=True)
        backup = folder / (target.stem + "_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".xlsx")
        shutil.copy2(target, backup)
        os.replace(temp, target)
        temp = None
        candidate.settings = target
        candidate.identity = hashlib.sha256(str(target).encode()).hexdigest()
        return candidate, str(backup)
    finally:
        book.close()
        if temp:
            temp.unlink(missing_ok=True)
