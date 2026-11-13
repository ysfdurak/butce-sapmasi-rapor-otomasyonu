"""Butce_Aciklama sorgusunun (powerquery/Butce_Aciklama.pq) Python benzetimi.

Müdürlüklerden dönen rapor dosyalarının "Açıklamalar" sayfasını okur, önceki onay tablosuyla birleştirir ve işlem
geçmişini ekler. Sorgunun ürettiği tablonun kesin tanımıdır; CI'da gerçek M motorunun çıktısıyla karşılaştırılır.

Kurallar:
  - Anahtar = Bölüm|Gider kodu|Kapanış|Aylar (açıklama yalnız aynı kapanış ve ay seçimi için geçerlidir).
  - Durum ve Onay notu kullanıcı tarafından tabloda doldurulur; yenilemede korunur.
  - Durum, Kayıtlı durum'dan farklıysa işlem geçmişine "zaman · kişi · durum · not" eklenir.
  - Dosyadaki açıklama metni değişirse durum Bekliyor'a döner. Kaynak dosyası kaldırılan satır silinmez (Kaynakta=Yok).
  - "Bu döneme aktar": önceki kapanışın açıklaması bu kapanış ve ay seçimi için onaylı yeni satır olur.
"""

import sys
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "python"))

COLUMNS = ["Anahtar", "Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Açıklama", "Yazan", "Yazma tarihi",
           "Kapanış", "Aylar", "Gerçekleşen (dosyada)", "Bütçe (dosyada)", "Kaynak dosya", "Dosya tarihi",
           "Okuma notu", "Kaynakta", "Durum", "Onay notu", "Son işlem", "İşlemi yapan", "İşlem zamanı",
           "İşlem geçmişi", "Kayıtlı durum"]
STATES = ["Bekliyor", "Onaylandı", "Reddedildi", "Bu döneme aktar"]
REQUIRED = ["Bölüm", "Gider kodu", "Açıklama", "Aylar", "Kapanış"]


def key(value):
    return "" if value is None else str(value).strip().translate(str.maketrans("İI", "iı")).lower()


def empty(v):
    return v is None or (isinstance(v, str) and not v.strip())


def text(v):
    return None if empty(v) else str(v).strip()


def number(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def date_text(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return text(v)


def num_text(v):
    """M Text.From(sayı): 1 → "1", 1.5 → "1.5"."""
    return str(int(v)) if isinstance(v, (int, float)) and float(v).is_integer() else str(v)


def read_file(path, departments):
    from openpyxl import load_workbook

    book = load_workbook(path, data_only=True, read_only=True)
    try:
        if "Açıklamalar" not in book.sheetnames:
            raise ValueError("Açıklamalar sayfası yok; dosya rapor şablonundan hazırlanmamış")
        rows = list(book["Açıklamalar"].iter_rows(values_only=True))
    finally:
        book.close()
    header = [str(h) if h is not None else f"Column{i}" for i, h in enumerate(rows[0], 1)]
    missing = [c for c in REQUIRED if c not in header]
    if missing:
        raise ValueError("Açıklamalar sayfasında sütun eksik: " + ", ".join(missing))
    stamp = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
    result = []
    for values in rows[1:]:
        r = dict(zip(header, values))
        if empty(r.get("Bölüm")) or empty(r.get("Açıklama")):
            continue
        try:
            code = float(r.get("Gider kodu"))
        except (TypeError, ValueError):
            code = None
        dept = departments.get(key(r["Bölüm"]))
        dept_name = text(r["Bölüm"]) if dept is None else dept[0]
        manager = None if dept is None else dept[1]
        file_manager = text(r.get("Dosya müdürlüğü"))
        item = text(r.get("Gider kalemi"))
        closing, months = text(r.get("Kapanış")), text(r.get("Aylar"))
        if dept is None:
            note = "Bölüm Ayarlar'da yok"
        elif code is None:
            note = "Gider kodu sayı değil"
        elif item == "Bu bölümde bu kod yok":
            note = "Bu bölümde bu kod yok"
        elif file_manager is not None and key(file_manager) != key(manager):
            note = f"Bölüm dosyanın müdürlüğünde değil ({file_manager})"
        elif closing is None or months is None:
            note = "Kapanış/ay bilgisi yok: dosya hesaplanmadan kaydedilmiş"
        else:
            note = None
        code_text = num_text(code) if code is not None else str(r.get("Gider kodu"))
        result.append({
            "Anahtar": f"{dept_name}|{code_text}|{closing or '?'}|{months or '?'}",
            "Müdürlük": manager, "Bölüm": dept_name, "Gider kodu": code,
            "Gider kalemi": None if item == "Bu bölümde bu kod yok" else item,
            "Açıklama": text(r["Açıklama"]), "Yazan": text(r.get("Yazan")), "Yazma tarihi": date_text(r.get("Tarih")),
            "Kapanış": closing, "Aylar": months,
            "Gerçekleşen (dosyada)": number(r.get("Gerçekleşen (seçili aylar)")),
            "Bütçe (dosyada)": number(r.get("Bütçe (seçili aylar)")),
            "Kaynak dosya": path.name, "Dosya tarihi": stamp, "Okuma notu": note,
        })
    return result


def departments_of(settings_path):
    from openpyxl import load_workbook

    book = load_workbook(settings_path, data_only=True, read_only=True)
    try:
        rows = list(book["Bölüm Eşleştirme"].iter_rows(values_only=True))
    finally:
        book.close()
    start = next(i for i, r in enumerate(rows) if key(r[0]) == "maliyet sayfası / bölüm")
    header = list(rows[start])
    out = {}
    for r in rows[start + 1:]:
        if empty(r[0]):
            continue
        d = dict(zip(header, r))
        out.setdefault(key(r[0]), (text(r[0]), text(d.get("Müdürlük"))))
    return out


def build(settings_path, folder, previous=(), user="", period=None, months=None, now="2026-11-13 10:00"):
    """previous: önceki tablonun satırları (sözlük listesi). Dönüş: COLUMNS sırasıyla sözlük listesi."""
    settings_path, folder = Path(settings_path), Path(folder)
    departments = departments_of(settings_path)
    actor_default = user.strip() if not empty(user) else "(Kontrol Paneli'nde ad yok)"

    collected0 = []
    for path in sorted((p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".xlsx"
                        and not p.name.startswith("~$")), key=lambda p: p.name):
        try:
            collected0 += read_file(path, departments)
        except Exception as error:  # noqa: BLE001 — sorgu da her okuma hatasını satıra yazar
            collected0.append({"Anahtar": "Dosya|" + path.name, "Kaynak dosya": path.name,
                               "Dosya tarihi": datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
                               "Okuma notu": "Dosya okunamadı: " + str(error), "Kaynakta": "Var"})
    new_keys = list(dict.fromkeys(r["Anahtar"] for r in collected0))
    collected = []
    for k in new_keys:
        group = [r for r in collected0 if r["Anahtar"] == k]
        chosen = sorted(group, key=lambda r: r["Dosya tarihi"] + "|" + r["Kaynak dosya"])[-1]
        texts = {r.get("Açıklama") for r in group}
        if len(texts) > 1 and empty(chosen.get("Okuma notu")):
            files = ", ".join(r["Kaynak dosya"] for r in group)
            chosen = {**chosen, "Okuma notu": f"Çakışma: farklı açıklamalar ({files}); en yeni dosya alındı, düzeltin"}
        collected.append(chosen)

    seen = set()
    old = []
    for r in previous:
        k = r.get("Anahtar")
        if empty(k) or str(k).startswith("Dosya|") or k in seen:
            continue
        seen.add(k)
        old.append(r)
    old_keys = [str(r["Anahtar"]) for r in old]

    def line(actor, action, note):
        return f"{now} · {actor} · {action}" + ("" if empty(note) else f" · {note}")

    def merge(y, e):
        if y is not None and y["Anahtar"].startswith("Dosya|"):
            return dict(y)
        base = y if y is not None else e
        g = (lambda name: None) if e is None else e.get
        here = "Var" if y is not None else ("Aktarıldı" if g("Kaynakta") == "Aktarıldı" else "Yok")
        e_state, stored = text(g("Durum")), text(g("Kayıtlı durum"))
        changed = e is not None and y is not None and text(g("Açıklama")) != text(y.get("Açıklama"))
        carry = e is not None and e_state == "Bu döneme aktar"
        removed = e is not None and y is None and g("Kaynakta") == "Var"
        valid = e_state if e_state in STATES else "Bekliyor"
        state = "Bekliyor" if e is None or changed else ((stored or "Bekliyor") if carry else valid)
        if e is None:
            action = "Yeni açıklama"
        elif changed:
            action = "Açıklama değişti, yeniden onay gerekir"
        elif carry:
            action = f"Bu döneme aktarıldı ({period or '?'})"
        elif removed:
            action = "Kaynak dosyadan kaldırıldı"
        elif valid != (stored or "Bekliyor"):
            action = valid
        else:
            action = None
        if e is None or changed:
            actor = "Müdürlük" if empty(y.get("Yazan")) else y.get("Yazan")
        elif removed:
            actor = "Sistem"
        else:
            actor = actor_default
        note = g("Onay notu")
        row = {c: base.get(c) for c in COLUMNS if c in base}
        history = g("İşlem geçmişi")
        if action is not None:
            entry = line(actor, action, note if e is not None and not changed and not removed else None)
            history = entry if empty(history) else f"{history}\n{entry}"
        row.update({"Kaynakta": here, "Durum": state, "Onay notu": note,
                    "Son işlem": g("Son işlem") if action is None else action,
                    "İşlemi yapan": g("İşlemi yapan") if action is None else actor,
                    "İşlem zamanı": g("İşlem zamanı") if action is None else now,
                    "İşlem geçmişi": history, "Kayıtlı durum": state})
        return row

    all_keys = new_keys + [k for k in old_keys if k not in new_keys]
    rows = [merge(collected[new_keys.index(k)] if k in new_keys else None,
                  old[old_keys.index(k)] if k in old_keys else None) for k in all_keys]
    for e in old:
        new_key = f"{e['Bölüm']}|{num_text(e['Gider kodu'])}|{period}|{months}"
        if (text(e.get("Durum")) != "Bu döneme aktar" or period is None or months is None or new_key in all_keys):
            continue
        row = {c: e.get(c) for c in ["Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Açıklama", "Yazan", "Yazma tarihi"]}
        row.update({"Anahtar": new_key, "Kapanış": period, "Aylar": months,
                    "Kaynak dosya": f"Önceki dönemden: {e['Kapanış']}", "Kaynakta": "Aktarıldı", "Durum": "Onaylandı",
                    "Son işlem": "Önceki dönemden aktarıldı ve onaylandı", "İşlemi yapan": actor_default,
                    "İşlem zamanı": now,
                    "İşlem geçmişi": line(actor_default, f"Önceki dönemden aktarıldı ve onaylandı ({e['Kapanış']})", None),
                    "Kayıtlı durum": "Onaylandı"})
        rows.append(row)
    rows = [{c: r.get(c) for c in COLUMNS} for r in rows]

    def sort_key(r):  # M Table.Sort: Kapanış azalan, sonra artan; null en küçük
        def asc(v):
            return (0, "") if v is None else (1, v)
        return asc(r["Müdürlük"]), asc(r["Bölüm"]), asc(r["Gider kodu"]), asc(r["Anahtar"])

    rows.sort(key=sort_key)
    rows.sort(key=lambda r: (0, "") if r["Kapanış"] is None else (1, r["Kapanış"]), reverse=True)
    return rows


if __name__ == "__main__":
    for r in build(ROOT / "python/ornek/Ayarlar.xlsx", sys.argv[1], user="Deneme", period="2026-10",
                   months="2026:1,2,3,4,5,6,7,8,9,10"):
        print(r["Anahtar"], r["Durum"], r["Okuma notu"], r["Son işlem"])
