# Doğrulama

| Alan | Durum | Kanıt / kabul adımı |
| --- | --- | --- |
| Python hesaplama, kaynak, V5 analizleri ve rapor akışı | 94 otomatik test geçti | `cd python` ve `python run_tests.py`; GitHub Actions her gönderimde çalıştırır |
| Python JavaScript/API | V4'te 72 benzetim ve yerel API kontrolü geçti (V5'te arayüz hesabı değişmedi) | [Python doğrulama kaydı](../python/Dogrulama_Durumu.json) |
| VBA kaynakları | `vba_elle` kopyaları güncel; ayrılmış sözcük, Windows-1254 kodlaması ve kurulum modül listesi denetimi geçti | `python tools/vba_araclari.py kontrol`; GitHub Actions |
| VBA ↔ Python eşdeğerliği | Aynı sentetik beklenen değerler iki tarafta da test edilir | `python/tests/test_analiz.py` ve makrodaki **Sentetik hesap testi** |
| Python Windows başlangıcı, dosya pencereleri ve gerçek tarayıcı yerleşimi | Canlı test bekliyor | [Canlı kontrol listesi](CANLI_KONTROL.md) |
| VBA derleme, sentetik hesap düğmesi ve Excel rapor çıktısı | Windows masaüstü Excel'de canlı test bekliyor | [Canlı kontrol listesi](CANLI_KONTROL.md) · [VBA doğrulama kaydı](../makro/Dogrulama_Durumu.json) |

Önce [sentetik demoyu](BASLANGIC.md) çalıştırın. Ekim kapanışında beş raporu ve beklenen tutarları karşılaştırın. Ardından yalnız Ocak, Mart ve Mayıs'ı seçip önceki yıl kapsamının da aynı kaldığını kontrol edin.

Yorum senaryosunda iki Eylül müdürlük dosyasından 10 yorum alınır. İkisi aynen, sekizi düzenlenerek topluca onaylanır. Bekleyen yorumların kırmızı gösterimini ve aynı dosyalar tekrar alındığında mevcut metinlerin korunmasını sınayın. [Örnek yorumlar ve adımlar](../makro/ornek/onceki_yorumlar/TEST_ADIMLARI.md)
