# V3 Windows Excel kabul kontrolü

1. Kurulum.vbs ile V3 xlsm oluşturun; Alt+F11 > Debug > Compile VBAProject çalıştırın. Dokuz modül görünmeli.
2. PanelHazirla ile Çalışma Merkezi ve yeni düğmeler açılmalı. Kaynaklar/ay/çıktı seçimi eski panelle aynı çalışmalı.
3. SentetikTest'i ve önceki CANLI_KONTROL.md adımlarını tamamlayın.
4. Kontrolleri yenile: kaynaklar, dip toplam farkları, kapsam dışı BM/OPEX bölümleri görünmeli.
5. Eşleştirme incelemesini onaylayın. Kopya kaynakta bölüm adı değiştirerek eski/yeni çifti deneyin; kimlik aynı kalmalı. Eski ve yeni adları içeren aynı dönem/tutarlı açıklama dosyaları birleştirilebilmeli.
6. Hesap İzleme D4 / D5 seçimlerini değiştirip yenileyin; gerçek kaynak hücreleri ile tutarlar eşit olmalı. İlk beş adet satırı gider hücrelerine karışmamalı.
7. Tablo Düzeni satırını seçin, Kaynak hücresinden seç düğmesiyle kaynak dosyada bir alan işaretleyin. D sütunu uygun satır/kolon/adresle dolmalı. İptal de çalışmalı. Geçersiz yeni düzen orijinal ayarı değiştirmemeli; geçerli değişiklik yedek oluşturmalı.
8. İki görünüm kaydedin/açın; tam seçili ay kümesi korunmalı.
9. Rapor oluşturun: dört müdürlük + genel beş çıktı. Çalıştırma geçmişi ve Calisma_Arsivi oluşmalı. Önceki aynı ayarlı çalışmayla kaynak farklarını gösterin. Arşivden yeniden üretim yeni çıktı klasörüne yazmalı.
10. Güvenli durdur düğmesini kaynak okuma ve raporlar arasında deneyin; önceki çıktılar kalmalı, geçici rapor klasörü yayımlanmamalı. Başka iş düğmesi devam eden işi değiştirememeli.
11. Dosyayı kapatıp tekrar açın: eşleştirme kayıtları, görünümler ve geçmiş kaybolmamalı. Başarısız işlemden sonra yeni çalışma önceki geçmiş satırını değiştirmemeli.

Bu adımlar burada çalıştırılmadı; VBA kaynak/paket incelemesi canlı Excel testinin yerine geçmez.
