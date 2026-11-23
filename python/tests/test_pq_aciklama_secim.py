"""Denetim D19: dönen açıklama dosyalarında "en yeni" özgün değiştirilme zamanıyla seçilir.

Aynı dakika içinde gelen iki dosyada dakikalık metin eşittir; önceden dosya adı sırası karar veriyordu ve adı
sonra gelen ama daha eski dosya seçilebiliyordu. Farklı metinler için çakışma uyarısı sürmeli.
"""

import os
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))

import pq_aciklama  # noqa: E402


def sahte_oku(path, departments):
    metin = {"A_yeni.xlsx": "Yeni açıklama", "Z_eski.xlsx": "Eski açıklama"}[path.name]
    return [{"Anahtar": "Giriş|1|2026-10|2026:1", "Müdürlük": "Mal Kabul", "Bölüm": "Giriş", "Gider kodu": 1,
             "Açıklama": metin, "Kaynak dosya": path.name,
             "Dosya tarihi": datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
             "Okuma notu": None, "_zaman": path.stat().st_mtime_ns}]


class AyniDakikaSecimi(unittest.TestCase):
    def test_ozgun_zaman_karar_verir_cakisma_korunur(self):
        with tempfile.TemporaryDirectory() as t:
            d = Path(t)
            taban = int(datetime(2026, 11, 20, 10, 0, 0).timestamp()) * 10**9
            for ad, sn in (("Z_eski.xlsx", 5), ("A_yeni.xlsx", 40)):  # aynı dakika: 10:00:05 ve 10:00:40
                (d / ad).write_bytes(b"x")
                os.utime(d / ad, ns=(taban + sn * 10**9, taban + sn * 10**9))
            with patch.object(pq_aciklama, "read_file", sahte_oku), patch.object(pq_aciklama, "departments_of", lambda p: {}):
                rows = pq_aciklama.build(Path("ayar.xlsx"), d)
        satir = [r for r in rows if r.get("Anahtar") == "Giriş|1|2026-10|2026:1"]
        self.assertEqual(len(satir), 1)
        self.assertEqual(satir[0]["Kaynak dosya"], "A_yeni.xlsx")
        self.assertEqual(satir[0]["Açıklama"], "Yeni açıklama")
        self.assertIn("Çakışma", satir[0]["Okuma notu"])
        self.assertNotIn("_zaman", satir[0])


if __name__ == "__main__":
    unittest.main()
