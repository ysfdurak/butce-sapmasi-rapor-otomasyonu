"""Personel bütçe tamamlama kuralı: 3. ve 4. personel göstergesi boşsa 1./2. göstergenin bütçesi kullanılır.

VBA karşılığı: MakroYaz.IndicatorBudget.
"""


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
