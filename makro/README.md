# Bütçe Sapması — Excel Makro Rapor V5

Windows masaüstü Excel üzerinde bağımsız VBA rapor uygulaması. Python veya PowerShell gerekmez. On bir modül içerir. `Kurulum.vbs` bilgisayarınızda `Kontrol_Sablonu.xlsx`'ten `Butce_Kontrol_Makro_V5.xlsm` oluşturur; dosya varsa yedek alıp modülleri günceller. Hazır derlenmiş xlsm bu pakette yoktur.

[Kurulum ve kullanım](KULLANIM.md) ile başlayın. VBA proje erişimi kapalıysa aynı belgedeki elle kurulum kullanılabilir (`vba_elle/` klasörü). Şirket güvenlik ayarları otomatik değiştirilmez.

Raporlar Python sürümüyle aynı hesap kurallarını kullanır: Yönetici Özeti, Genel Özet, Giderler, Birim Maliyet, Önemli Sapmalar, Yıl Sonu Tahmini, Hacim ve Birim Maliyet, Kontrol ve Revizyon. Kurallar: [Veri ve hesap](../docs/VERI_VE_HESAP.md). Açıklama akışı ve onay izi: [Açıklamalar](../docs/YORUM_KULLANIMI.md).

Çalışma Merkezi; kontrol listesi, hesap izi, eşleştirme değişiklikleri ve kalıcı kimlikler, adlandırılmış rapor seçimleri, kaynak alanı işaretleme, girdi arşivi (son 20 çalışma), yeniden üretim ve güvenli durdurmayı sunar.

**Geliştirme:** Tek kaynak `vba/` klasörüdür (UTF-8). `vba_elle/` kopyaları `python tools/vba_araclari.py uret` ile üretilir; GitHub Actions `python tools/vba_araclari.py kontrol` ile kopyaların güncel olduğunu ve ayrılmış sözcük/kodlama hatası olmadığını denetler. Windows Excel'de derleme ve çalıştırma [canlı kontrol listesiyle](../docs/CANLI_KONTROL.md) yapılır; [test durumu](Dogrulama_Durumu.json) bu sınırı kaydeder.
