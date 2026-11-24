"""Makro: müdürlük raporunun Giderler sayfası başlıkları açıklama okuyucusuyla uyumlu olmalı.

MakroYaz her bölüm için "… Gerçekleşen TL", "… Bütçe TL", "Bütçe - Gerçekleşen TL" yazar.
MakroAciklama.ExpensePositions gerçekleşen sütununu başlıkta "gerceklesen" arayarak bulur;
fark sütunu da bu sözcüğü taşıdığı için her bölüm iki kez sayılıyor, rapor oluşturma
"Geçersiz bölüm başlığı" ile duruyordu (Windows Excel, 24 Kasım 2026). Bu test, yazılan
başlıklardan yalnız birinin okuyucunun koşuluna uyduğunu denetler.
"""

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
VBA = ROOT / "makro" / "vba"


def k(s: str) -> str:
    s = s.strip().replace("İ", "i").replace("I", "ı").lower()
    for a, b in zip("ışğüöç", "isguoc"):
        s = s.replace(a, b)
    return s


def okuyucu_secer(title: str) -> bool:
    t = k(title)
    return "gerceklesen" in t and t[:5] != "butce"


class GiderBasliklari(unittest.TestCase):
    def test_kosul_kaynakta(self):
        kaynak = (VBA / "MakroAciklama.bas").read_text(encoding="utf-8")
        self.assertIn('InStr(title, "gerceklesen") > 0 And Left$(title, 5) <> "butce"', kaynak)

    def test_bolum_basina_tek_gerceklesen(self):
        yaz = (VBA / "MakroYaz.bas").read_text(encoding="utf-8")
        satir = next(l for l in yaz.splitlines() if 'RowValues ws, 5, Array(SelectionCaption() & " Gerçekleşen TL"' in l)
        basliklar = ["2026 / Aylar 10" + m if m.startswith(" ") else m for m in re.findall(r'"([^"]*)"', satir)]
        self.assertEqual(sum(okuyucu_secer(b) for b in basliklar), 1, basliklar)


class IlkAcilis(unittest.TestCase):
    def test_auto_open_paneli_hazirlar(self):
        panel = (VBA / "MakroPanel.bas").read_text(encoding="utf-8")
        govde = panel.split("Public Sub Auto_Open()", 1)[1].split("End Sub", 1)[0]
        self.assertIn('ListBoxes("MRV_Aylar")', govde)
        self.assertIn("PanelHazirla", govde)
        self.assertIn("PanelDurum", govde)


if __name__ == "__main__":
    unittest.main()
