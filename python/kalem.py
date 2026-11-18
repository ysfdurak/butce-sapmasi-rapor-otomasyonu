"""Kalem ayrıntısı: bir gider satırının bütçe, gerçekleşen, hacim / birim maliyet etkisi ve kaynak hücreleri.

Hesap kuralları analiz.volume_price_rows ile aynıdır (bkz. docs/VERI_VE_HESAP.md), yalnız bölüm toplamı
yerine tek gider satırına uygulanır:

* Bölümün bütçe işlem adedi (PA) ve gerçekleşen işlem adedi (A) seçili ayların toplamıdır. Bir ayda adet
  boş, sıfır veya negatifse o bölümde etkiler boş kalır (bölüm düzeyindeki kuralla aynı).
* Kalemin bütçe birim maliyeti = kalem bütçesi / PA.
* Hacim etkisi = (A − PA) × kalem bütçe birim maliyeti.
* Birim maliyet etkisi = kalem gerçekleşeni − kalem bütçe birim maliyeti × A.
* Hacim + birim maliyet etkisi = kalem gerçekleşeni − kalem bütçesi (sapma). Bir bölümün bütün kalemlerinin
  etkileri toplandığında "Hacim ve Birim Maliyet" sayfasındaki bölüm satırı elde edilir.

Tutarlar TL ve Decimal'dir. Kaynak hücreler workspace.source_cells ile aynı okumadan gelir; hücre
toplamlarının hesaplanan tutarla tuttuğu ayrıca kontrol edilir (mutabakat).
"""

from decimal import Decimal

import analiz
import butce
import workspace

ZERO = Decimal(0)
KURUS = Decimal("0.01")


def _sum(values):
    values = list(values)
    return sum(values, ZERO) if values and None not in values else None


def _count(data, name, months, key):
    """Seçili ayların işlem adedi toplamı; bir ay boş/sıfır/negatifse None (analiz kuralı)."""
    values = []
    for m in months:
        v = data.records.get((name, data.budget_year, m), {}).get(key)
        values.append(v if v is not None and v > ZERO else None)
    return _sum(values)


def _effects(data, name, months, code):
    year = [data.budget_year]
    actual = data.aggregate([name], year, months, code=code)
    budget = data.aggregate([name], year, months, "budget", code=code)
    count = _count(data, name, months, "count")
    planned = _count(data, name, months, "count_budget")
    unit_budget = butce.ratio(budget, planned)
    volume = price = None
    reason = ""
    if planned is None:
        reason = "Bütçe işlem adedi eksik, sıfır veya negatif bir ay var; etkiler hesaplanmadı."
    elif count is None:
        reason = "Gerçekleşen işlem adedi eksik, sıfır veya negatif bir ay var; etkiler hesaplanmadı."
    elif actual is None or budget is None:
        reason = "Kalemin gerçekleşen veya bütçe verisi eksik."
    else:
        volume = (count - planned) * unit_budget
        price = actual - unit_budget * count
    return {
        "actual": actual,
        "budget": budget,
        "variance": butce.delta(actual, budget),
        "count": count,
        "plannedCount": planned,
        "unitActual": butce.ratio(actual, count),
        "unitBudget": unit_budget,
        "volume": volume,
        "price": price,
        "reason": reason,
    }


def detail(data, state, code, department=None, notes=None):
    """Tek gider satırının ayrıntısı. department verilirse yalnız o bölüm, yoksa seçimdeki bütün bölümler."""
    if workspace.inventory(data) != data.fingerprint:
        raise ValueError("Kaynak dosyaları değişmiş; kalem ayrıntısını açmadan önce kaynakları yenileyin.")
    names, years, months = data.selection(state)
    if department:
        if department not in names:
            raise ValueError("Bölüm seçim kapsamında değil.")
        names = [department]
    if type(code) is not int or code not in data.catalog:
        raise ValueError("Geçersiz gider kodu.")
    label, group = data.catalog[code]
    budget_ready = years == [data.budget_year] and bool(months)
    notes = notes or {}

    departments = []
    for n in names:
        if budget_ready:
            row = _effects(data, n, months, code)
        else:
            row = {
                "actual": data.aggregate([n], years, months, code=code),
                "budget": None,
                "variance": None,
                "count": None,
                "plannedCount": None,
                "unitActual": None,
                "unitBudget": None,
                "volume": None,
                "price": None,
                "reason": "Bütçe ve etkiler yalnız bütçe yılı tek başına seçildiğinde hesaplanır.",
            }
        note = notes.get((n, code), {})
        row.update(
            department=n,
            manager=data.config["bolumler"][n]["mudurluk"],
            rate=butce.ratio(row["variance"], row["budget"]),
            direction=analiz.direction(row["variance"]),
            material=analiz.is_material(row["variance"], row["budget"], analiz.settings(data.config)),
            note=note.get("note", ""),
            noteStale=bool(note.get("stale")),
            noteStatus=note.get("status", ""),
            noteSource=note.get("source", ""),
        )
        departments.append(row)

    actual = _sum(r["actual"] for r in departments)
    budget = _sum(r["budget"] for r in departments) if budget_ready else None
    variance = butce.delta(actual, budget)
    volume = _sum(r["volume"] for r in departments) if budget_ready else None
    price = _sum(r["price"] for r in departments) if budget_ready else None
    residual = None
    if None not in (variance, volume, price):
        residual = (variance - volume - price).quantize(KURUS)

    sources, missing = workspace.source_cells(data, names, years, months, code, counts=budget_ready)
    if workspace.inventory(data) != data.fingerprint:
        raise ValueError("Kaynaklar kalem ayrıntısı okunurken değişti; yeniden yenileyin.")

    def cell_total(kind):
        return _sum(r[9] for r in sources if r[4] == kind)

    reconciliation = []
    for kind, expected in (("Gerçekleşen", actual), ("Bütçe", budget)):
        if kind == "Bütçe" and not budget_ready:
            continue
        found = cell_total(kind)
        reconciliation.append(
            {
                "kind": kind,
                "cells": sum(1 for r in sources if r[4] == kind),
                "cellTotal": found,
                "calculated": expected,
                "ok": found is not None and expected is not None and found == expected,
            }
        )

    rules = analiz.settings(data.config)
    return {
        "code": code,
        "name": label,
        "group": group,
        "scope": {"departments": names, "years": years, "months": months, "single": bool(department)},
        "budgetReady": budget_ready,
        "actual": actual,
        "budget": budget,
        "variance": variance,
        "rate": butce.ratio(variance, budget),
        "direction": analiz.direction(variance),
        "material": analiz.is_material(variance, budget, rules),
        "threshold": analiz.threshold_text(rules),
        "volume": volume,
        "price": price,
        "residual": residual,
        "departments": departments,
        "sources": sources,
        "missingDepartments": missing,
        "reconciliation": reconciliation,
        "formula": (
            "Sapma = gerçekleşen − bütçe (pozitif: Aleyhte). Kalem bütçe birim maliyeti = kalem bütçesi / "
            "bölümün bütçe işlem adedi. Hacim etkisi = (gerçekleşen adet − bütçe adet) × kalem bütçe birim "
            "maliyeti. Birim maliyet etkisi = kalem gerçekleşeni − kalem bütçe birim maliyeti × gerçekleşen adet. "
            "İki etkinin toplamı sapmaya eşittir; bölümün bütün kalemleri toplandığında Hacim ve Birim Maliyet "
            "sayfasındaki bölüm satırı çıkar."
        ),
    }
