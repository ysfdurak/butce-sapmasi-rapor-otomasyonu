"""Power Query şablonunun veri sözleşmesi ve Python benzetimi.

`powerquery/Butce_Veri.pq` sorgusu Excel'e TEK bir tablo yükler (Veri sayfası, A1). Tablonun her satırı bir
"blok"a aittir (Tablo sütunu); şablondaki formüller satırları Anahtar sütunuyla bulur. Bu modül aynı tabloyu
Python motorunun okuyucularıyla üretir: şablon formülleri bu çıktıyla test edilir ve M sorgusunun üretmesi
gereken tablonun kesin tanımıdır.

Sütunlar (36): Anahtar, Tablo, Direktörlük, Müdürlük, Bölüm, Ad, Kod, Grup, Yıl, Kaynak, Metin, Sayı,
G1..G12 (aylık birinci değer), B1..B12 (aylık ikinci değer; genellikle bütçe).

Bloklar ve anahtarlar:
  Kapanis                        Yıl, Sayı=ay
  Ayar|<kod>                     Sayı=değer (tolerans, onemlilik_tutari, onemlilik_orani, ozet_kalem_sayisi)
  Bolum|<sıra>                   Bölüm, Müdürlük, Direktörlük, Ad=endeks anahtarı, Kod=kontrol grubu sırası,
                                 Grup=kontrol bölümü, Metin=lokasyon, Kaynak=kapanış dosyası (yoksa boş), Sayı=sıra
  Mudurluk|<sıra>                Müdürlük, Direktörlük, Sayı=sıra
  Duzen|<sıra>                   Ad=satır başlığı, Grup=satır türü, Bölüm/Müdürlük/Direktörlük, Sayı=sıra
  Kalem|<sıra>                   Bölüm, Müdürlük, Kod, Ad=gider kalemi, Grup=gider grubu, Sayı=sıra
  Satir|<sıra>                   Giderler sayfasının satırları: Metin=grup/kalem/toplam, Grup, Kod, Ad, Sayı=sıra
  GrupAd|<sıra>                  Gider grubu adları (görünüş sırası), Ad, Sayı=sıra
  KGrup|<sıra>                   Kontrol grubu: Ad=kontrol bölümü, Metin=anahtar, Grup=direktörlük anahtarı,
                                 Direktörlük, Müdürlük=müdürlükler, Bölüm=bölümler, Sayı=sıra
  Gider|<bölüm>|<kod>|<yıl>      Ad, Grup, Kaynak; G=gerçekleşen, B=bütçe (yalnız kapanış yılı)
  Toplam|<bölüm>|<yıl>           G/B: kalemlerin aylık toplamı (bir kalem boşsa ay boş)
  Adet|<bölüm>|<yıl>             G=gerçekleşen işlem adedi, B=bütçe işlem adedi
  Kontrol|<dir>|<bölüm>|<n>|<yıl> Kontrol dosyasında eşleşen n. satır; G=aylık tutar
  Ciro|<yıl>                     G=gerçekleşen ciro, B=bütçe ciro
  Endeks|<endeks>|<yıl>          G=aylık endeks
  Revizyon|<sıra>                Metin=kapanış, Grup=tür, Bölüm, Müdürlük, Kod, Ad, Yıl/Sayı=etkilenen ay,
                                 G1=eski, G2=yeni, Kaynak=eski dosya → yeni dosya
  Uyari|<sıra>                   Ad=dönem, Bölüm, Kaynak, Metin=uyarı
  Kimlik                         Ad=Direktörlük/Müdürlük, Müdürlük=müdürlük dosyasıysa adı, Metin=yenileme zamanı
  Girdi|<sıra>                   Ad=dosya, Grup=Ayarlar/Kontrol/Endeks/Maliyet, Kod=bayt, Metin=değiştirilme zamanı,
                                 Kaynak=Okundu/Kapsam dışı/Dönem bulunamadı/Bulunamadı, Yıl/Sayı=dosya dönemi

Müdürlük dosyası (mudurluk parametresi): yalnız o müdürlüğün bölümleri okunur; birleşik kontrol grubu müdürlük dışına
taşıyorsa grup alınmaz ve "Birleşik bölüm kontrolü direktörlük raporunda yer alır" uyarısı yazılır.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "python"))

COLUMNS = (
    ["Anahtar", "Tablo", "Direktörlük", "Müdürlük", "Bölüm", "Ad", "Kod", "Grup", "Yıl", "Kaynak", "Metin", "Sayı"]
    + [f"G{m}" for m in range(1, 13)]
    + [f"B{m}" for m in range(1, 13)]
)


def _num(value):
    return None if value is None else float(value)


def row(key, table, **fields):
    values = {c: None for c in COLUMNS}
    values.update(Anahtar=key, Tablo=table)
    for name, value in fields.items():
        if name in ("G", "B"):
            for m, v in enumerate(value, 1):
                values[f"{name}{m}"] = _num(v)
        else:
            values[name] = value
    return [values[c] for c in COLUMNS]


def months_of(series, year):
    return [series.get((year, m)) for m in range(1, 13)]


def strict_sum(values):
    values = list(values)
    return None if not values or any(v is None for v in values) else sum(values)


def build(settings_path, closing=None, mudurluk=None):
    """mudurluk verilirse yalnız o müdürlüğün verisi (müdürlüğe gönderilecek dosya)."""
    from bootstrap import dependencies

    dependencies()
    import butce
    from birim import read_budget_counts

    settings_path = Path(settings_path)
    config = butce.read_settings(settings_path)
    root = settings_path.parent
    cost_dir = root / config["maliyet_klasoru"]
    if closing is None:
        closing = max(butce.filename_period(p) for p in butce.cost_files(cost_dir))
    warnings = []
    snapshots = butce.read_snapshots(cost_dir, config, closing, warnings)
    all_departments = config["bolumler"]
    if mudurluk:
        departments = {n: m for n, m in all_departments.items() if butce.key(m["mudurluk"]) == butce.key(mudurluk)}
        if not departments:
            raise ValueError(f"Müdürlük bulunamadı: {mudurluk}")
        snapshots = {k: v for k, v in snapshots.items() if k[1] in departments}
        warnings = [w for w in warnings if w[1] is None or w[1] in departments or w[1] not in all_departments]
    else:
        departments = all_departments
    directors = {butce.key(m["direktorluk"]) for m in all_departments.values()}
    config = {**config, "bolumler": departments}
    company, controls, company_budget = butce.read_control(
        root / config["kontrol_dosyasi"], config["kontrol"], closing[0], directors
    )
    indices = butce.read_indices(root / config["endeks_dosyasi"], config["endeks"])
    budget_counts = read_budget_counts(root, config, snapshots, closing)
    year = closing[0]
    rows = [row("Kapanis", "Kapanis", Yıl=year, Sayı=closing[1])]
    rows.append(row("Kimlik", "Kimlik", Ad="Müdürlük" if mudurluk else "Direktörlük",
                    Müdürlük=next(iter(departments.values()))["mudurluk"] if mudurluk else None))
    rows += input_rows(butce, settings_path, config, cost_dir, closing)

    rules = {"tolerans": config["kontrol"]["tolerans"], **(config.get("rapor") or {})}
    for code, value in rules.items():
        rows.append(row(f"Ayar|{code}", "Ayar", Ad=code, Sayı=float(value)))

    # Kontrol grupları: (direktörlük anahtarı, kontrol bölümü anahtarı) görünüş sırasıyla.
    groups = {}
    for name, m in all_departments.items():
        target = m.get("kontrol_bolumu") or name
        groups.setdefault((butce.key(m["direktorluk"]), butce.key(target)), [target, m["direktorluk"], []])[2].append(
            name
        )
    cross = sorted({g[0] for g in groups.values() if set(g[2]) & set(departments) and not set(g[2]) <= set(departments)})
    groups = {k: g for k, g in groups.items() if set(g[2]) <= set(departments)}
    group_index = {k: i for i, k in enumerate(groups, 1)}
    for i, (name, m) in enumerate(departments.items(), 1):
        snap = snapshots.get((closing, name))
        target = m.get("kontrol_bolumu") or name
        rows.append(
            row(
                f"Bolum|{i}", "Bolum", Bölüm=name, Müdürlük=m["mudurluk"], Direktörlük=m["direktorluk"],
                Ad=butce.key(m.get("endeks")) if m.get("endeks") else None,
                Kod=group_index.get((butce.key(m["direktorluk"]), butce.key(target)), 0), Grup=target,
                Metin=m.get("lokasyon"), Kaynak=snap[3] if snap else None, Sayı=i,
            )
        )
    managers = []
    for m in departments.values():
        if (m["mudurluk"], m["direktorluk"]) not in managers:
            managers.append((m["mudurluk"], m["direktorluk"]))
    for i, (manager, director) in enumerate(managers, 1):
        rows.append(row(f"Mudurluk|{i}", "Mudurluk", Müdürlük=manager, Direktörlük=director, Sayı=i))
    for i, (title, level, names) in enumerate(butce.hierarchy(departments), 1):
        first = departments[names[0]]
        rows.append(
            row(
                f"Duzen|{i}", "Duzen", Ad=title, Grup=level, Bölüm=title if level == "Bölüm" else None,
                Müdürlük=first["mudurluk"] if level in ("Bölüm", "Müdürlük toplamı") else None,
                Direktörlük=first["direktorluk"] if level != "Genel toplam" else None, Sayı=i,
            )
        )
    for i, ((dkey, tkey), (target, director, names)) in enumerate(groups.items(), 1):
        rows.append(
            row(
                f"KGrup|{i}", "KGrup", Ad=target, Metin=tkey, Grup=dkey, Direktörlük=director,
                Müdürlük=", ".join(sorted({departments[n]["mudurluk"] for n in names})),
                Bölüm=", ".join(names), Sayı=i,
            )
        )

    def group_of(code):
        return next((g for lo, hi, g in config["gider_gruplari"] if lo <= code <= hi), "Eşleşmeyen")

    kalem, satir_groups = [], {}
    for name, m in departments.items():
        snap = snapshots.get((closing, name))
        if snap is None:
            continue
        items, _, counts, source = snap
        actual_years = {p[0] for item in items.values() for p in item[1]}
        for code, (item, actual, budget) in sorted(items.items()):
            kalem.append((name, m["mudurluk"], code, item, group_of(code)))
            satir_groups.setdefault(group_of(code), {})[(code, item)] = True
            for y in sorted(actual_years | {year}):
                rows.append(
                    row(
                        f"Gider|{name}|{code}|{y}", "Gider", Bölüm=name, Müdürlük=m["mudurluk"],
                        Direktörlük=m["direktorluk"], Ad=item, Kod=code, Grup=group_of(code), Yıl=y, Kaynak=source,
                        G=months_of(actual, y), B=months_of(budget, y) if y == year else [None] * 12,
                    )
                )
        for y in sorted(actual_years | {year}):
            totals_g = [strict_sum(it[1].get((y, mm)) for it in items.values()) for mm in range(1, 13)]
            totals_b = (
                [strict_sum(it[2].get((y, mm)) for it in items.values()) for mm in range(1, 13)]
                if y == year
                else [None] * 12
            )
            rows.append(row(f"Toplam|{name}|{y}", "Toplam", Bölüm=name, Müdürlük=m["mudurluk"],
                            Direktörlük=m["direktorluk"], Yıl=y, G=totals_g, B=totals_b))
            planned = budget_counts.get(name, {})
            rows.append(row(f"Adet|{name}|{y}", "Adet", Bölüm=name, Müdürlük=m["mudurluk"],
                            Direktörlük=m["direktorluk"], Yıl=y, G=months_of(counts, y),
                            B=months_of(planned, y) if y == year else [None] * 12))
    for i, (name, manager, code, item, group) in enumerate(kalem, 1):
        rows.append(row(f"Kalem|{i}", "Kalem", Bölüm=name, Müdürlük=manager, Kod=code, Ad=item, Grup=group, Sayı=i))
    k = 0
    for gi, (group, items) in enumerate(satir_groups.items(), 1):
        rows.append(row(f"GrupAd|{gi}", "GrupAd", Ad=group, Sayı=gi))
        k += 1
        rows.append(row(f"Satir|{k}", "Satir", Metin="grup", Grup=group, Ad="Grup toplamı", Sayı=k))
        for code, item in sorted(items):
            k += 1
            rows.append(row(f"Satir|{k}", "Satir", Metin="kalem", Grup=group, Kod=code, Ad=item, Sayı=k))
    k += 1
    rows.append(row(f"Satir|{k}", "Satir", Metin="toplam", Grup="TOPLAM", Ad="Tüm giderler", Sayı=k))

    for (dkey, tkey), matches in controls.items():
        if dkey not in directors or (mudurluk and (dkey, tkey) not in groups):
            continue
        for n, series in enumerate(matches, 1):
            for y in sorted({p[0] for p in series}):
                rows.append(row(f"Kontrol|{dkey}|{tkey}|{n}|{y}", "Kontrol", Ad=tkey, Grup=dkey, Kod=n, Yıl=y,
                                G=months_of(series, y)))
    for y in sorted({p[0] for p in company} | {p[0] for p in company_budget}):
        rows.append(row(f"Ciro|{y}", "Ciro", Yıl=y, G=months_of(company, y), B=months_of(company_budget, y)))
    wanted = {butce.key(m["endeks"]) for m in departments.values() if m.get("endeks")}
    for name in sorted({n for n, _ in indices if not mudurluk or n in wanted}):
        for y in sorted({p[0] for n, p in indices if n == name}):
            rows.append(row(f"Endeks|{name}|{y}", "Endeks", Ad=name, Yıl=y,
                            G=[indices.get((name, (y, mm))) for mm in range(1, 13)]))
    for i, r in enumerate(butce.revisions(snapshots, config, closing), 1):
        closing_label, affected, manager, name, code, item, kind, before, after, _, old_file, new_file = r
        ay, am = map(int, affected.split("-"))
        rows.append(row(f"Revizyon|{i}", "Revizyon", Metin=closing_label, Grup=kind, Bölüm=name, Müdürlük=manager,
                        Kod=code, Ad=item, Yıl=ay, Sayı=am, G=[before, after] + [None] * 10,
                        Kaynak=f"{old_file} → {new_file}"))
    for name in departments:
        if snapshots.get((closing, name)) is None:
            continue  # "Kapanış dosyası yok" uyarısını şablon kendisi üretir
        for before, after, _ in butce.revision_pairs(*closing):
            if not (snapshots.get((before, name)) and snapshots.get((after, name))):
                warnings.append([butce.label(after), name, None,
                                 "Ardışık dosya eksik. Bu kapanış için revizyon karşılaştırması yapılamadı."])
    for target in cross:
        warnings.append([butce.label(closing), None, None, f"{target}: Birleşik bölüm kontrolü direktörlük raporunda yer alır"])
    for i, (period_label, name, source, message) in enumerate(warnings, 1):
        rows.append(row(f"Uyari|{i}", "Uyari", Ad=period_label, Bölüm=name, Kaynak=source, Metin=message))
    return rows


def input_rows(butce, settings_path, config, cost_dir, closing):
    """Girdi kaydı: raporun okuduğu dosyalar, boyutları ve değiştirilme zamanları."""
    from datetime import datetime

    def info(path, kind, status=None, period=None):
        exists = path.is_file()
        stat = path.stat() if exists else None
        return dict(Ad=path.name, Grup=kind, Kod=stat.st_size if exists else None,
                    Metin=datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S") if exists else None,
                    Kaynak=status or ("Okundu" if exists else "Bulunamadı"),
                    Yıl=period[0] if period else None, Sayı=period[1] if period else None)

    root = settings_path.parent
    items = [info(settings_path, "Ayarlar"), info(root / config["kontrol_dosyasi"], "Kontrol"),
             info(root / config["endeks_dosyasi"], "Endeks")]
    for path in sorted(cost_dir.iterdir(), key=lambda p: p.name):
        if not path.is_file() or path.suffix.lower() != ".xlsx" or path.name.startswith("~$"):
            continue
        try:
            p = butce.filename_period(path)
        except ValueError:
            items.append(info(path, "Maliyet", "Dönem bulunamadı"))
            continue
        items.append(info(path, "Maliyet", "Okundu" if butce.in_closing_scope(p, closing) else "Kapsam dışı", p))
    return [row(f"Girdi|{i}", "Girdi", **item) for i, item in enumerate(items, 1)]


if __name__ == "__main__":
    data = build(sys.argv[1] if len(sys.argv) > 1 else ROOT / "python/ornek/Ayarlar.xlsx")
    from collections import Counter

    print(len(data), Counter(r[1] for r in data))
