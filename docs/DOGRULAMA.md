# Doğrulama

| Alan | Durum | Kanıt / kabul adımı |
| --- | --- | --- |
| Python hesaplama, kaynak ve rapor akışı | 77 otomatik test geçti | `cd python` ve `python run_tests.py` |
| Python JavaScript/API | V4'te 72 benzetim ve yerel API kontrolü geçti | [V4 doğrulama kaydı](../python/Dogrulama_Durumu.json) |
| Python Windows başlangıcı, dosya pencereleri ve gerçek tarayıcı yerleşimi | Canlı test bekliyor | [Python canlı kontrol listesi](../python/V4_CANLI_KONTROL.md) |
| Makro modülleri ve paket | Kaynak ve paket kontrolleri geçti | [VBA doğrulama kaydı](../makro/Dogrulama_Durumu.json) |
| VBA derleme, sentetik hesap düğmesi ve Excel rapor çıktısı | Windows masaüstü Excel'de canlı test bekliyor | [VBA canlı kontrol listesi](../makro/V4_CANLI_KONTROL.md) |

Önce [sentetik demoyu](BASLANGIC.md) çalıştırın. Ekim kapanışında beş raporu ve beklenen tutarları karşılaştırın. Ardından yalnız Ocak, Mart ve Mayıs'ı seçip önceki yıl kapsamının da aynı kaldığını kontrol edin.

Yorum senaryosunda iki Eylül müdürlük dosyasından 10 yorum alınır. İkisi aynen, sekizi düzenlenerek topluca onaylanır. Bekleyen yorumların kırmızı gösterimini ve aynı dosyalar tekrar alındığında mevcut metinlerin korunmasını sınayın. [Örnek yorumlar ve adımlar](../makro/ornek/onceki_yorumlar/TEST_ADIMLARI.md)
