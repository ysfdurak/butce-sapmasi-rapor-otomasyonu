# V4 canlı kabul kontrolü

Örnek Ayarlar.xlsx, Ekim 2026 kapanışı ve Ocak–Ekim aylarıyla deneyin.

1. ornek/onceki_yorumlar klasöründeki iki Onceki_2026_09 dosyasını getirin: 10 kayıt kontrol beklemeli, önceki ay kümesi Ocak–Eylül görünmeli.
2. Beklenen senaryo Excel'ine göre 2 yorumu aynen topluca onaylayın; kalan 8'i tek tabloda düzenleyip topluca onaylayın. 10 onaylı kayıt olmalı.
3. Farklı ay kümesinde tekrar deneyin: ortak metni yalnız seçilen iki taslağa uygulayın, taslakları onaysız kaydedin. Raporlarda kırmızı ve etiketli görünmeli. Onay sonrası normal renk olmalı.
4. İki yorumu bu dönem geçersiz sayın: rapora alınmamalı, geçmiş metinler inceleme listesinde durmalı.
5. Bu döneme ait bir açıklama girin/yükleyin, önceki dosyaları yeniden getirin: mevcut açıklama ve düzenlenen taslaklar ezilmemeli. Başka bir satırı onaylarken kaydedilmemiş diğer düzenlemeler kaybolmamalı.
6. Kaynak veya ayar değişmişken eski inceleme ekranından onaylamayı deneyin: güncel olmayan kayıt reddedilmeli. Normal renkteki bir açıklamanın tutarı değişirse tekrar inceleme gerekmeli.
7. Genel ve müdürlük çıktılarında açıklama kapsamı, etiket, renk ve Açıklama Durumu sayfasını kontrol edin. İlgisiz müdürlüğün açıklaması bulunmamalı. Boş durum sayfası oluşmamalı.
8. Kaydedip kapatın, yeniden açın. Onaylar, taslaklar ve geçersiz kararları korunmalı.

Python: gerçek tarayıcı düzeni, Windows dosya pencereleri ve küçük ekran kaydırmalarını kontrol edin. İki pencereyle aynı yorumu düzenleyip eski pencerenin güncel kaydı ezemediğini doğrulayın.

Makro: önce Alt+F11 > Debug > Compile VBAProject çalıştırın; on modül yüklenmiş olmalı. Excel çıktısına elle girilen mevcut yorumlar önce Bu dönem yorumlarını yükle ile okunmalıdır. Taslak sütunu G'dir. Yeni sürümü boş kontrol dosyasında deneyin; eski xlsm'yi koruyun. Burada VBA canlı çalıştırması yapılmadı.
