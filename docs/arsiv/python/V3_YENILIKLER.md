# V3 — Çalışma merkezi

Bu sürüm mevcut yedi Python dashboard sayfasını ve rapor hesaplarını korur. Yerel çalışır; şirket kaynaklarını dış servise göndermez. Başlat.bat veya Başlat.pyw yerine paket içindeki mevcut ASCII adlarıyla **Baslat.bat / Baslat.pyw** kullanılır.

## Kullanım

1. Baslat.bat ile açın. Üstte **Çalışma merkezi** düğmesine basın.
2. Kaynaklar bölümünde ayar dosyasını ve kapanışı seçin; kaynakları yenileyin.
3. Kontrol merkezinde beklenen/okunan bölümleri, sorunları ve işlem önerilerini inceleyin. Mevcut dip toplam doğrulaması ve revizyon sayfaları da durur. Uyarı olması raporun tümünün doğru olduğu anlamına gelmez; özellikle eksik kaynaklar düzeltilmelidir.
4. Eşleştirme değişikliklerini son onayla karşılaştırın. Bir bölümün adı değişmişse yeni bölümün karşısında eski adını seçin ve onaylayın. Otomatik benzerlik önerisi tek başına eşleştirme oluşturmaz. Onaylı ad değişiminde açıklamalar aynı kalıcı kimlikle taşınır; tutar değişmişse önceki açıklama güncel sayılmaz.
5. **Hesap nasıl oluştu?** bölümünde tüm seçili giderler veya tek bölüm/kalem için dosya, sayfa, hücre ve değerleri görün. Ham gider birleşmeleri ayrı kaynak satırlarıyla izlenir; ilk beş adet satırı parasal toplamın dışındadır. Birim maliyet, endeks oranları ve ağırlıklı iyileşme yöntemi açıklanır. Kaynak dosyaları değiştiyse eski toplamla yeni hücreleri karıştırmamak için önce yenileme istenir.
6. **Tablo düzeni** bölümünde örnek kaynak hücresine tıklayarak seçili parametreyi doldurabilir veya parametre tablosunda elle değiştirebilirsiniz. Yalnız mevcut ayar anahtarları düzenlenir. Kaydetmeden önce kaynakların tamamı okunur; geçersiz adayda eski ayar korunur, geçerli değişiklikte yedek alınır.
7. Birden fazla seçim kümesini adlandırarak **Kayıtlı görünümler** altında saklayın. Genel filtreler ve bağımsız BM/OPEX bölüm seçimleri birlikte kaydedilir. Her ay için yeni bir görünüm oluşturmak gerekmez; kayıtlı görünüm seçimini kontrol ederek güncelleyebilirsiniz.
8. **Rapor oluştur** ile ayar Excel'indeki bütün bölümler için çıktı alın. Dashboard'un daraltılmış organizasyon seçimi rapor kapsamını daraltmaz; rapor yılı ve ayları uygulanır. Dört müdürlük örneğinde beş rapor çıkar.
9. İşlem sırasında aşama ve dosya adı gösterilir. **Güvenli durdur** mevcut dosya/Excel yazım adımı bittikten sonra durur. Tamamlanmamış çıktı klasörü yayımlanmaz. Ayar kaydı atomik bir adımdır; kaydetme tamamlanmışsa geri alınmış gibi gösterilmez.

## Geçmiş ve yeniden üretim

Her rapor çalışmasında `kullanici/calismalar/<kimlik>` altında kullanılan ayarlar, maliyet dosyaları, ciro/kontrol ve endeks dosyaları kopyalanır. Kaynak imzaları, organizasyon eşleştirmeleri, tam ay kümesi, hesap motoru imzaları ve o anda geçerli açıklamalar kayda girer. İlk çalışma bütün kaynakları yeni gösterir; sonraki çalışma aynı ayar yolunun önceki çalışmasıyla karşılaştırılır.

**Bu girdilerle yeniden üret** güncel dosyalar yerine saklanan girdileri ve o çalışmanın açıklamalarını kullanır. Veri veya hesap motoru imzası değişmişse sessizce farklı sonuç üretmez. Eski yazılım sürümü farklıysa aynı eski paketle yeniden üretim gerekir. Her çalışma ayrı çıktı klasörüne yazılır; eski raporlar değiştirilmez. Girdi kopyaları disk alanı kullanır; arşivleri ihtiyaca göre yedekleyin.

## V2'den geçiş

V3'ü yeni klasöre açın. Açıklamalarınız ve kayıtlı kaynak/seçimleriniz varsa eski paketin `kullanici` klasörünü uygulama kapalıyken V3'e kopyalayın. Asıl ayar Excel dosyasını aynı konumdan seçin; kaynak ayar yolu farklılaşırsa ayrı veri kümesi kimliği oluşur. Eski klasörü silmeden yeni sürümü kontrol edin.

## Yeni açık talepler

Tekrarlayan davranış için otomatik açıklama ve ortak sayfadan eşik üstü gider detayı aktarımı henüz etkin değildir. `ORTAK_KARARLAR.md` kararları ve açık alanları içerir. Yeni gider detayı için TL VE yüzde koşulu birlikte kullanılacak; benzer davranışın tanımını kullanıcı verecek.
