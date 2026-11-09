"""Excel-independent dashboard calculations. Monetary arithmetic stays Decimal."""

from pathlib import Path
from decimal import Decimal
from datetime import datetime
from collections import defaultdict
import hashlib
import butce
from openpyxl import load_workbook
from personel import budget_values

D = Decimal
MONTHS = butce.MONTHS


def total(values):
    return butce.total(values)


def sub(a, b):
    return butce.delta(a, b)


def div(a, b):
    return butce.ratio(a, b)


def mul(a, b):
    return a * b if a is not None and b is not None else None


def million(a):
    return a / D(1000000) if a is not None else None


def positive(a):
    return a is not None and a > 0


def present(a):
    return a is not None


def change(a, b):
    gap = sub(a, b)
    return {"value": gap, "rate": div(gap, b)}


class Dataset:
    def __init__(self, settings, closing, progress=None):
        progress = progress or (lambda message: None)
        progress("Ayarlar okunuyor")
        self.settings = Path(settings).expanduser().resolve()
        self.config = butce.read_settings(self.settings)
        self.root = self.settings.parent
        self.closing = tuple(closing)
        self.budget_year, self.closing_month = self.closing
        if not (1901 <= self.budget_year <= 9998 and 1 <= self.closing_month <= 12):
            raise ValueError("Geçersiz kaynak kapanışı.")
        from workspace import inventory

        before = inventory(self)
        self.warnings = []
        progress("Maliyet dosyaları okunuyor")
        self.snapshots = butce.read_snapshots(
            self.root / self.config["maliyet_klasoru"],
            self.config,
            self.closing,
            self.warnings,
            progress=progress,
        )
        progress("Ciro ve dip toplam kaynakları okunuyor")
        self.revenue, self.controls, self.revenue_budget = butce.read_control(
            self.root / self.config["kontrol_dosyasi"],
            self.config["kontrol"],
            self.budget_year,
            {butce.key(m["direktorluk"]) for m in self.config["bolumler"].values()},
        )
        progress("Endeks okunuyor")
        self.indices = butce.read_indices(self.root / self.config["endeks_dosyasi"], self.config["endeks"])
        self.departments = list(self.config["bolumler"])
        self.years = sorted(
            {
                y
                for (closing, name), snap in self.snapshots.items()
                if closing == self.closing
                for _, a, _ in snap[0].values()
                for y, m in a
                if y <= self.budget_year
            }
        )
        if not self.years:
            raise ValueError("Seçilen kapanışta okunabilir gerçekleşen veri yok.")
        self.catalog = {}
        self.records = {}
        self.sources = []
        self.staff_mismatch = []
        self.indicator_labels = {}
        for name in self.departments:
            snap = self.snapshots.get((self.closing, name))
            if not snap:
                self.warnings.append([butce.label(self.closing), name, "", "Kapanış kaynağı eksik."])
                continue
            self.sources.append({"department": name, "file": snap[3]})
            for code, (label, a, b) in snap[0].items():
                group = next(
                    (g for lo, hi, g in self.config["gider_gruplari"] if lo <= code <= hi), "Eşleşmeyen"
                )
                if code in self.catalog and self.catalog[code] != (label, group):
                    raise ValueError(f"{code}: Bölümler arasında gider adı/grubu farklı.")
                self.catalog[code] = (label, group)
        self.catalog = dict(sorted(self.catalog.items()))
        self.groups = list(dict.fromkeys(g for _, g in self.catalog.values()))
        if not self.catalog:
            raise ValueError("Gider kalemi bulunamadı.")
        progress("Personel ve işlem göstergeleri hazırlanıyor")
        self._read_indicators()
        progress("Geriye dönük değişiklikler karşılaştırılıyor")
        self.revisions = list(butce.revisions(self.snapshots, self.config, self.closing))
        if inventory(self) != before:
            raise ValueError("Kaynaklar okunurken değişti; yeniden yenileyin.")
        self.fingerprint = before
        self.loaded_at = datetime.now().isoformat(timespec="seconds")
        self.identity = hashlib.sha256(str(self.settings).encode()).hexdigest()

    def _read_indicators(self):
        books = {}
        try:
            for name in self.departments:
                snap = self.snapshots.get((self.closing, name))
                actual_cols = {}
                budget_cols = {}
                sheet = None
                if snap:
                    path = self.root / self.config["maliyet_klasoru"] / snap[3]
                    if path not in books:
                        books[path] = load_workbook(path, data_only=True)
                    sheet = books[path][name]
                    actual_cols, budget_cols = butce.positioned_columns(
                        sheet, self.config["maliyet"], self.closing
                    )
                    indicator_row = self.config["maliyet"]["islem_adedi_satiri"]
                    label_column = self.config["maliyet"].get(
                        "gosterge_adi_sutunu", self.config["maliyet"]["gider_sutunu"]
                    )
                    labels = []
                    for r in range(indicator_row - 4, indicator_row + 1):
                        value = sheet.cell(r, label_column).value
                        label = value.strip() if isinstance(value, str) else ""
                        if not label:
                            label = f"Başlığı eksik gösterge ({sheet.cell(r,label_column).coordinate})"
                            self.warnings.append([butce.label(self.closing), name, snap[3], label])
                        labels.append(label)
                    self.indicator_labels[name] = labels
                for year in self.years:
                    for month in range(1, 13):
                        p = (year, month)
                        live = p <= self.closing
                        items = {
                            code: (
                                item[1].get(p) if live else None,
                                item[2].get(p) if year == self.budget_year else None,
                            )
                            for code, item in (snap[0] if snap else {}).items()
                        }
                        cost = total(v[0] for v in items.values())
                        budget = total(v[1] for v in items.values())
                        counts = snap[2].get(p) if snap and live else None
                        budget_count = None
                        staff = [None] * 4
                        staff_budget = [None] * 4
                        if sheet is not None:
                            row = self.config["maliyet"]["islem_adedi_satiri"]
                            if p in actual_cols and live:
                                staff = [
                                    butce.number(sheet.cell(row - 4 + i, actual_cols[p]).value)
                                    for i in range(4)
                                ]
                            if p in budget_cols and year == self.budget_year:
                                budget_count = butce.number(sheet.cell(row, budget_cols[p]).value)
                                raw = [sheet.cell(row - 4 + i, budget_cols[p]).value for i in range(4)]
                                vals = [butce.number(v) for v in raw]
                                if any(v not in (None, "") and n is None for v, n in zip(raw, vals)):
                                    raise ValueError(
                                        f"{name} {year}-{month:02d}: Personel bütçesinde sayı dışında değer var."
                                    )
                                staff_budget, mismatch, _ = budget_values(vals)
                                if mismatch:
                                    self.staff_mismatch.append([name, year, month])
                        self.records[name, year, month] = {
                            "actual": cost,
                            "budget": budget,
                            "count": counts,
                            "count_budget": budget_count,
                            "staff": staff,
                            "staff_budget": staff_budget,
                            "items": items,
                        }
        finally:
            for b in books.values():
                b.close()

    def meta(self):
        fields = {"manager": "mudurluk", "submanager": "mudurluk_alti", "location": "lokasyon"}
        options = {
            k: list(dict.fromkeys(m.get(field) or "Tanımsız" for m in self.config["bolumler"].values()))
            for k, field in fields.items()
        }
        options.update(department=self.departments, year=self.years, month=list(range(1, 13)))
        files = []
        for path in sorted((self.root / self.config["maliyet_klasoru"]).glob("*.xlsx")):
            if path.name.startswith("~$"):
                continue
            try:
                period = butce.filename_period(path)
                status = (
                    "Rapor kaynağı"
                    if period == self.closing
                    else (
                        "Revizyon karşılaştırması"
                        if period[0] == self.budget_year and period < self.closing
                        else "Seçim dışında"
                    )
                )
                files.append([path.name, butce.label(period), status])
            except ValueError as error:
                files.append([path.name, "", str(error)])
        latest = max((butce.period(r[1]) for r in files if r[1]), default=self.closing)
        mappings = [
            {"department": name, **meta, "hasSource": (self.closing, name) in self.snapshots}
            for name, meta in self.config["bolumler"].items()
        ]
        return {
            "settings": str(self.settings),
            "closing": list(self.closing),
            "latestClosing": list(latest),
            "budgetYear": self.budget_year,
            "options": options,
            "files": files,
            "mappings": mappings,
            "loadedAt": self.loaded_at,
            "sources": self.sources,
            "warnings": self.warnings,
            "revisionCount": len(self.revisions),
            "defaults": {"year": [self.budget_year], "month": list(range(1, self.closing_month + 1))},
        }

    def selection(self, state):
        options = self.meta()["options"]
        filters = state.get("filters", {})
        selected = {}
        for key, values in options.items():
            value = filters.get(key, None)
            if value is None:
                value = (
                    values
                    if key not in ("year", "month")
                    else ([self.budget_year] if key == "year" else list(range(1, self.closing_month + 1)))
                )
            if not isinstance(value, list) or any(v not in values for v in value):
                raise ValueError(f"Geçersiz seçim: {key}")
            selected[key] = list(dict.fromkeys(value))
        names = []
        for name in self.departments:
            meta = self.config["bolumler"][name]
            if (
                name in selected["department"]
                and meta["mudurluk"] in selected["manager"]
                and (meta.get("mudurluk_alti") or "Tanımsız") in selected["submanager"]
                and (meta.get("lokasyon") or "Tanımsız") in selected["location"]
            ):
                names.append(name)
        return names, sorted(selected["year"]), sorted(selected["month"])

    def aggregate(self, names, years, months, key="actual", code=None, staff=None, positive_only=False):
        vals = []
        for name in names:
            for year in years:
                for month in months:
                    rec = self.records.get((name, year, month), {})
                    if code is not None:
                        value = rec.get("items", {}).get(code, (None, None))[0 if key == "actual" else 1]
                    elif staff is not None:
                        value = rec.get(key, [None] * 4)[staff]
                    else:
                        value = rec.get(key)
                    if positive_only and not positive(value):
                        value = None
                    vals.append(value)
        return total(vals)

    def _cards(self, names, years, months, counts=False):
        key = "count" if counts else "actual"
        bkey = "count_budget" if counts else "budget"
        a = self.aggregate(names, years, months, key)
        b = (
            self.aggregate(names, years, months, bkey)
            if years == [self.budget_year] and a is not None
            else None
        )
        return {
            "actual": a if counts else million(a),
            "budget": b if counts else million(b),
            "gap": sub(b, a) if counts else million(sub(b, a)),
            "rate": div(sub(b, a), b),
            "unit": "adet" if counts else "Milyon TL",
        }

    def indicator_title(self, names, index):
        variants = {}
        for name in names:
            values = self.indicator_labels.get(name)
            if values:
                variants.setdefault(values[index], []).append(name)
        title = " / ".join(variants) if variants else f"Kaynak başlığı bulunamadı ({index+1}. gösterge)"
        return title, [{"name": label, "departments": depts} for label, depts in variants.items()]

    def _entities(self, names):
        result = []
        for manager in dict.fromkeys(self.config["bolumler"][n]["mudurluk"] for n in names):
            local = [n for n in names if self.config["bolumler"][n]["mudurluk"] == manager]
            result.extend((n, "department", [n]) for n in local)
            result.append((manager + " Toplamı", "manager", local))
        if names:
            result.append(("Direktörlük Genel Toplamı", "total", names))
        return result

    def _unit_info(self, name, year, months):
        meta = self.config["bolumler"][name]
        alias = meta.get("endeks") or ""
        mapped = bool(alias and any(k == butce.key(alias) for k, p in self.indices))
        a = self.aggregate([name], [year], months)
        b = self.aggregate([name], [year - 1], months)
        x = self.aggregate([name], [year], months, "count", positive_only=True)
        y = self.aggregate([name], [year - 1], months, "count", positive_only=True)
        f = butce.period_index(self.indices, alias, [(year, m) for m in months]) if months else None
        actual = div(a, x)
        prior = div(b, y)
        adjusted = mul(prior, f)
        rate = div(sub(adjusted, actual), adjusted)
        return {
            "actual": actual,
            "prior": prior,
            "budget": (
                div(
                    self.aggregate([name], [year], months, "budget"),
                    self.aggregate([name], [year], months, "count_budget", positive_only=True),
                )
                if year == self.budget_year
                else None
            ),
            "factor": f,
            "inflation": sub(f, D(1)),
            "adjusted": adjusted,
            "rate": rate,
            "cost": a,
            "count": x,
            "eligible": bool(
                mapped
                and a is not None
                and a >= 0
                and positive(prior)
                and positive(x)
                and positive(y)
                and rate is not None
            ),
            "visible": bool(mapped and actual is not None),
        }

    def query(self, state):
        names, years, months = self.selection(state)
        cards = self._cards(names, years, months)
        manager_rows = []
        for manager in dict.fromkeys(self.config["bolumler"][n]["mudurluk"] for n in names):
            local = [n for n in names if self.config["bolumler"][n]["mudurluk"] == manager]
            manager_rows.append({"name": manager, **self._cards(local, years, months)})
        trends = [
            {
                "year": y,
                "month": m,
                "actual": million(self.aggregate(names, [y], [m])),
                "budget": (
                    million(self.aggregate(names, [y], [m], "budget"))
                    if years == [self.budget_year]
                    else None
                ),
            }
            for y in years
            for m in months
        ]
        staff = []
        for person in range(4):
            monthly = [
                {
                    "year": y,
                    "month": m,
                    "actual": self.aggregate(names, [y], [m], "staff", staff=person),
                    "budget": (
                        self.aggregate(names, [y], [m], "staff_budget", staff=person)
                        if years == [self.budget_year]
                        else None
                    ),
                }
                for y in years
                for m in months
            ]
            totals = [
                {
                    "year": y,
                    "actual": self.aggregate(names, [y], months, "staff", staff=person),
                    "budget": (
                        self.aggregate(names, [y], months, "staff_budget", staff=person)
                        if years == [self.budget_year]
                        else None
                    ),
                }
                for y in years
            ]
            title, variants = self.indicator_title(names, person)
            staff.append({"name": title, "titleVariants": variants, "monthly": monthly, "totals": totals})
        transaction = []
        for year in sorted(years, reverse=True):
            vals = []
            for m in [*months, None]:
                span = [m] if m else months
                a = self.aggregate(names, [year], span, "count")
                prev = self.aggregate(names, [year - 1], span, "count")
                b = self.aggregate(names, [year], span, "count_budget") if year == self.budget_year else None
                vals.append(
                    {
                        "month": m,
                        "actual": a,
                        "budget": b,
                        "annual": change(a, prev),
                        "gap": change(b, a),
                        "gapRate": div(sub(b, a), b),
                    }
                )
            transaction.append({"year": year, "values": vals})
        detail = []
        for code, (label, group) in self.catalog.items():
            a = self.aggregate(names, years, months, "actual", code=code)
            b = (
                self.aggregate(names, years, months, "budget", code=code)
                if years == [self.budget_year] and a is not None
                else None
            )
            detail.append(
                {
                    "code": code,
                    "name": label,
                    "group": group,
                    "actual": million(a),
                    "budget": million(b),
                    "gap": million(sub(b, a)),
                    "rate": div(sub(b, a), b),
                    "history": [
                        {
                            "year": y,
                            "actual": million(self.aggregate(names, [y], months, "actual", code=code)),
                        }
                        for y in years
                    ],
                }
            )
        groups = []
        for group in self.groups:
            items = [r for r in detail if r["group"] == group]
            a = total(r["actual"] for r in items)
            b = total(r["budget"] for r in items)
            groups.append(
                {
                    "name": group,
                    "actual": a,
                    "budget": b,
                    "gap": sub(b, a),
                    "rate": div(sub(b, a), b),
                    "items": items,
                    "history": [
                        {"year": y, "actual": total(r["history"][i]["actual"] for r in items)}
                        for i, y in enumerate(years)
                    ],
                }
            )
        expense = []
        for label, kind, local in self._entities(names):
            vals = []
            for y in years:
                a = self.aggregate(local, [y], months)
                b = self.aggregate(local, [y], months, "budget") if y == self.budget_year else None
                vals.append(
                    {
                        "year": y,
                        "actual": million(a),
                        "budget": million(b),
                        "gap": million(sub(b, a)),
                        "rate": div(sub(b, a), b),
                    }
                )
            expense.append({"name": label, "kind": kind, "values": vals})
        opex_names = self.local_scope(state, "opexLocal", names)

        def opex_ok(n, y):
            a = self.aggregate([n], [y], months)
            b = self.aggregate([n], [y - 1], months)
            ra = total(self.revenue.get((y, m)) for m in months)
            rb = total(self.revenue.get((y - 1, m)) for m in months)
            return a is not None and positive(b) and positive(ra) and positive(rb)

        active = self.budget_year in years
        if active:
            opex_names = [n for n in opex_names if opex_ok(n, self.budget_year)]
        else:
            opex_names = [n for n in opex_names if any(opex_ok(n, y) for y in years)]
        opex = []
        for label, kind, local in self._entities(opex_names):
            vals = []
            for y in years:
                scope = local if active else [n for n in local if opex_ok(n, y)]
                a = div(self.aggregate(scope, [y], months), total(self.revenue.get((y, m)) for m in months))
                b = (
                    div(
                        self.aggregate(scope, [y], months, "budget"),
                        total(self.revenue_budget.get((y, m)) for m in months),
                    )
                    if y == self.budget_year
                    else None
                )
                prior = div(
                    self.aggregate(scope, [y - 1], months),
                    total(self.revenue.get((y - 1, m)) for m in months),
                )
                vals.append(
                    {
                        "year": y,
                        "actual": a,
                        "budget": b,
                        "annualImprovement": div(sub(prior, a), prior) if y == self.budget_year else None,
                        "budgetImprovement": div(sub(b, a), b) if y == self.budget_year else None,
                    }
                )
            opex.append({"name": label, "kind": kind, "values": vals})
        bm_names = self.local_scope(state, "bmLocal", names)
        infos = {
            (n, y): self._unit_info(n, y, months) for n in bm_names for y in set(years + [self.budget_year])
        }
        if active:
            bm_names = [n for n in bm_names if infos[n, self.budget_year]["eligible"]]
        else:
            bm_names = [n for n in bm_names if any(infos[n, y]["visible"] for y in years)]
        bm = []
        active_cost = total(infos[n, self.budget_year]["cost"] for n in bm_names) if active else None
        for label, kind, local in self._entities(bm_names):
            vals = []
            for y in years:
                visible = [n for n in local if infos[n, y]["visible"]]
                cost = total(infos[n, y]["cost"] for n in visible)
                counts = total(infos[n, y]["count"] for n in visible)
                info = {"year": y, "actual": div(cost, counts)}
                if y == self.budget_year:
                    info.update(
                        prior=div(
                            total(self.aggregate([n], [y - 1], months) for n in local),
                            total(
                                self.aggregate([n], [y - 1], months, "count", positive_only=True)
                                for n in local
                            ),
                        ),
                        budget=div(
                            total(self.aggregate([n], [y], months, "budget") for n in local),
                            total(
                                self.aggregate([n], [y], months, "count_budget", positive_only=True)
                                for n in local
                            ),
                        ),
                        inflation=infos[local[0], y]["inflation"] if kind == "department" else None,
                        adjusted=div(
                            total(
                                mul(
                                    infos[n, y]["adjusted"],
                                    self.aggregate([n], [y - 1], months, "count", positive_only=True),
                                )
                                for n in local
                            ),
                            total(
                                self.aggregate([n], [y - 1], months, "count", positive_only=True)
                                for n in local
                            ),
                        ),
                        rate=div(total(mul(infos[n, y]["cost"], infos[n, y]["rate"]) for n in local), cost),
                        weight=div(cost, active_cost),
                        managerWeight=(
                            div(
                                cost,
                                total(
                                    infos[n, y]["cost"]
                                    for n in bm_names
                                    if self.config["bolumler"][n]["mudurluk"]
                                    == self.config["bolumler"][local[0]]["mudurluk"]
                                ),
                            )
                            if kind == "department"
                            else None
                        ),
                    )
                vals.append(info)
            bm.append({"name": label, "kind": kind, "values": vals})
        priority = []
        priority_incomplete = 0
        if years == [self.budget_year]:
            for name in names:
                for code, (label, group) in self.catalog.items():
                    a = self.aggregate([name], years, months, "actual", code=code)
                    b = self.aggregate([name], years, months, "budget", code=code)
                    if a is None or b is None:
                        priority_incomplete += 1
                        continue
                    over = a - b
                    if round(over, 2) > 0:
                        priority.append(
                            {
                                "department": name,
                                "manager": self.config["bolumler"][name]["mudurluk"],
                                "code": code,
                                "item": label,
                                "group": group,
                                "actual": million(a),
                                "budget": million(b),
                                "gap": million(b - a),
                                "overrun": million(over),
                                "rate": div(b - a, b),
                            }
                        )
            if priority_incomplete:
                priority = []
        priority.sort(key=lambda r: (-r["overrun"], r["department"], r["code"]))
        gross = total(r["overrun"] for r in priority) or D(0)
        net_count = (
            sum(
                1
                for n in names
                if (lambda a, b: a is not None and b is not None and a > b)(
                    self.aggregate([n], years, months), self.aggregate([n], years, months, "budget")
                )
            )
            if years == [self.budget_year]
            else None
        )
        checks = []
        if months:
            for selected_year in years:
                _, year_checks, _ = butce.make_report(
                    self.config,
                    self.closing,
                    (selected_year, months[-1]),
                    months,
                    self.snapshots,
                    self.revenue,
                    self.controls,
                    self.indices,
                    [],
                )
                checks.extend(r[:10] for r in year_checks if set(r[-1]) & set(names))
        messages = []
        if not names or not years or not months:
            messages.append("Seçime uygun bölüm ve dönem yok.")
        elif cards["actual"] is None:
            messages.append(
                "Seçili gerçekleşenlerde eksik veri veya kapanıştan sonraki ay var; toplam boş bırakıldı."
            )
        if len(years) > 1:
            messages.append(
                "Birden fazla yıl seçili. Genel kartlar ve Detay Durum yalnız gerçekleşenleri toplar; yıl bazlı tablolarda bütçe yılı korunur."
            )
        elif years != [self.budget_year]:
            messages.append("Seçili yılın bütçesi bu kaynakta bulunmuyor; bütçe alanları boş.")
        if self.staff_mismatch:
            messages.append(
                "İlk iki personel göstergesinin bütçesi farklı olan kayıtlar var; ilgili son iki boş bütçe hücresi tamamlanmadı."
            )
        if any(len(s["titleVariants"]) > 1 for s in staff):
            messages.append(
                "Seçili bölümlerde aynı satırdaki personel göstergelerinin adları farklı. Grafiklerde kaynak adları birlikte gösteriliyor; tek bölüm seçerek ilgili adı görebilirsiniz."
            )
        invalid = sum(r[7] != "Doğrulandı" for r in checks)
        if invalid:
            messages.append(
                f"{invalid} bölüm/ay dip toplamı doğrulanamadı. Doğrulama ve revizyonlar bölümünde eşleşmeyen bölümü veya tutar farkını inceleyin."
            )
        return {
            "selection": {"departments": names, "years": years, "months": months},
            "messages": messages,
            "general": {"cards": cards, "managers": manager_rows, "trend": trends},
            "staff": staff,
            "transactions": {
                "name": self.indicator_title(names, 4)[0],
                "cards": self._cards(names, years, months, True),
                "rows": transaction,
            },
            "detail": {
                "groups": groups,
                "total": {
                    "actual": cards["actual"],
                    "budget": cards["budget"],
                    "gap": cards["gap"],
                    "rate": cards["rate"],
                },
            },
            "departments": {"expense": expense, "opex": opex, "opexEligible": opex_names},
            "unit": {
                "rows": bm,
                "eligible": bm_names,
                "rate": (
                    next(
                        (
                            r["values"][years.index(self.budget_year)].get("rate")
                            for r in bm
                            if r["kind"] == "total"
                        ),
                        None,
                    )
                    if active
                    else None
                ),
            },
            "priority": {
                "rows": priority,
                "gross": gross if years == [self.budget_year] and not priority_incomplete else None,
                "netDepartments": net_count,
                "top5": div(total(r["overrun"] for r in priority[:5]), gross),
                "incomplete": priority_incomplete,
            },
            "audit": {
                "checks": checks,
                "warnings": self.warnings,
                "revisions": [r for r in self.revisions if r[3] in names],
            },
        }

    def local_scope(self, state, key, names):
        values = state.get(key)
        if values is None:
            return names[:]
        if not isinstance(values, list) or any(n not in self.departments for n in values):
            raise ValueError("Geçersiz yerel bölüm seçimi.")
        return [n for n in names if n in values]
