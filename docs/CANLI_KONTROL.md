# Windows canlı kabul kontrolü

Kutular Windows'ta kullanıcı tarafından işaretlenmek üzere bilerek boştur. Örnek `Ayarlar.xlsx`, Ekim 2026 kapanışı ve Ocak–Ekim aylarıyla deneyin; beklenen değerler [Başlangıç](BASLANGIC.md#beklenen-demo-değerleri) sayfasındadır.

## V5 — iki sürüm için ortak

- [ ] Rapor klasöründe beş dosya var; her dosya **Yönetici Özeti** sayfasıyla açılıyor ve tek sayfa (yatay) PDF olarak yazdırılabiliyor.
- [ ] Genel raporda Önemli Sapmalar 150 satır (75 aleyhte); ilk satır Toplama / Ücret giderleri, −20.764.946,98 TL, **Lehte** (yeşil).
- [ ] Depolama müdürlük dosyasında Önemli Sapmalar'ın **Açıklama durumu** sütunu `Açıklama gerekli`; Giderler'deki ilgili açıklama hücreleri turuncu. Bir hücreye açıklama yazınca durum `Açıklandı` oluyor.
- [ ] Yıl Sonu Tahmini direktörlük satırı: yıllık bütçe 4.108.586.425,59; tahmin 4.117.030.835,77; aylık ortalama gidiş 4.139.260.068,37.
- [ ] Hacim ve Birim Maliyet direktörlük satırında hacim etkisi + birim maliyet etkisi = gerçekleşen − bütçe (8.444.410,18).
- [ ] Ayarlar.xlsx'te `onemlilik_tutari` 250.000 ve `onemlilik_orani` 0,05 yapınca önemli sapma sayısı artıyor; `onemlilik_orani` 10 yapınca açık hata mesajı çıkıyor.
- [ ] Bir maliyet dosyasında bir tutarı metin olarak (`'1.234,56`) yazın: Kontrol ve Revizyon › Uyarılar'da hücre adresiyle "sayıya çevrildi" uyarısı var ve toplam değişmiyor.
- [ ] Maliyet dosyasının en üstüne iki satır ekleyin: uyarıda `BF61` / `EZ60` önerisi görünüyor.
- [ ] Bir açıklamayı onaylayın: Açıklama Durumu sayfasında Son işlem, İşlemi yapan (Windows kullanıcınız) ve İşlem zamanı dolu.
- [ ] 21'den fazla rapor üretince girdi arşivinde en fazla 20 çalışma kalıyor.

- [ ] Raporun ikinci sayfası **Pano**: filtreler Tümü iken gerçekleşen 3.449.383.390,31. Müdürlük = Depolama ve Gider grubu = Personel Giderleri seçince gerçekleşen 500.718.327,50, bütçe 415.183.233,99. Grafikler filtreyle değişiyor.

## Makrosuz Power Query yolu

- [ ] [Rehberdeki](POWER_QUERY.md) adımlarla sorgu şablonun `Veri!$A$1` hücresine yükleniyor; Kontrol Paneli'nde veri satırı > 0.
- [ ] Genel Özet direktörlük toplamı (Ocak–Ekim): gerçekleşen 3.449.383.390,31, bütçe 3.440.938.980,13.
- [ ] Kapsam = Depolama, ay seçimi değiştirilince tüm sayfalar yenilemeden güncelleniyor; Pano grafikleri değişiyor.
- [ ] Yeni ay dosyası klasöre konup **Tümünü Yenile** yapılınca kapanış ayı ve Revizyonlar güncelleniyor.

## Excel makro

- [ ] `Kurulum.vbs` Türkçe mesajlarla `Butce_Kontrol_Makro_V5.xlsm` oluşturuyor. Tekrar çalıştırınca "güncelle" seçeneği yedek alıp modülleri değiştiriyor; panel seçimleri ve yorumlar korunuyor.
- [ ] Alt+F11 › Debug › Compile VBAProject hatasız; on bir modül (MakroAnaliz dahil) yüklü.
- [ ] **Sentetik hesap testi** tüm kontrolleri geçiriyor (V5 analiz değerleri dahil).
- [ ] Türkçe Excel'de Genel Özet ve Giderler'in renkli sapma biçimleri hatasız uygulanıyor.
- [ ] Rapor sırasında bir hata oluşturun (ör. çıktı klasörünü silin): hata mesajından sonra Excel otomatik hesaplama ve ekran güncellemesi geri gelmiş.
- [ ] Çalıştırma Geçmişi'nden bir satırı yeniden üretin: paneldeki yıl/ay seçimi değişmeden kalıyor.
- [ ] Yorum akışı: [Açıklamalar](YORUM_KULLANIMI.md) sayfasındaki 10 yorumluk senaryo; taslak sütunu G. Kaydedip kapatıp açınca onaylar ve geçersiz kararları korunuyor.

## Python dashboard

ZIP'i tamamen çıkarıp Baslat.bat ile açın. Standart başlangıç Ekim 2026, Ocak–Ekim, bütün bölümlerdir. Python kurulu değilken Baslat.bat kurulum seçeneklerini Türkçe açıklıyor.

- [ ] Personel grafik başlıkları kaynak Excel'deki ilk dört adla, işlem tablosu/kart adı beşinci satır başlığıyla eşleşiyor. Aynı satırın bölüm bazında adı farklıysa tek bölüm seçiminde doğru ad görünüyor.
- [ ] Kaynaklar'da ayar Excel'ini ve üç kaynak konumunu Gözat ile seç. Konumları kaydet ve yenile sonrasında yedek yolu gösteriliyor; kapatıp açınca aynı yollar korunuyor.
- [ ] Kaynaklar penceresinde yedi bölüm eşleştirme alanı, ana gider grupları, gider birleştirmeleri ve ortak sayfa kayıtlarının tamamı görünüyor.
- [ ] Ayara yeni bölüm/müdürlük ekleyip kaynak sayfasını tamamla; yenileyince yeni kayıt Tümü kapsamına giriyor. Silme ve ad değiştirme de listeye yansıyor. Eşleştirmesiz kaynak adı açıkça uyarılıyor.

- [ ] İlk ekranda 7 sayfa ve 6 ortak listbox var; kartlar 3.440,94 / 3.449,38 / −8,44 / −%0,25 gösteriyor. Tutarlar milyon TL.
- [ ] Ortak listelerde seçim yapıp sayfa değiştirince seçim korunuyor. Ocak/Mart/Mayıs için gerçekleşen 913,19 ve bütçe 1.094,64 milyon TL.
- [ ] 2025 + 2026 seçince genel kartlarda yalnız gerçekleşen var. İşlem/Bölüm Analizi tablolarında 2026'nın bütçe metrikleri duruyor.
- [ ] Genel çizgi grafiğinde her ay kendi değerini gösteriyor. Alt/üst sınır ve +/−/↑/↓ çalışıyor; uzun grafik yatay kayıyor.
- [ ] Personel 4 ayrı grafik. Bütçe çizgi, gerçekleşenler farklı yıl renkleriyle sütun. Sayılar ondalıksız. Çok yıl seçince grafik genişliyor.
- [ ] İşlem adedinde yıllar +/− ile açılıyor; geçen yıl değişimi mevcut. Başlangıç gerçekleşen 69.216.031, bütçe 67.920.000.
- [ ] Detay Durum ana giderleri + ile açıyor, toplam en altta; halkada yalnız yüzdeler var. 2024/2025/2026 seçip gider gerçekleşeninin üzerine gelince o giderin üç yıllık çizgisi görünüyor.
- [ ] Bölüm gider tablosunda milyon TL birimi açık, müdürlük toplamı kendi bölümlerinden sonra geliyor.
- [ ] BM genel iyileşme −%1,52. BM yerel bölümünde yalnız Giriş seçilince ağırlık %100, iyileşme yaklaşık −%0,93; genel kart değişmiyor.
- [ ] OPEX yerel bölümünde yalnız Tartım seçince kendi tablosu daralıyor; BM'nin Giriş seçimi korunuyor. Üst ortak filtreler değişmiyor.
- [ ] Öncelik sayfasında brüt aşım 230,82 milyon TL, 9 aşan bölüm, 234 kayıt, ilk beş payı %23,56.
- [ ] İlk aşım kaydına açıklama yazıp Seçimi kaydet; uygulamayı kapat/aç. Aynı ay kümesinde açıklama ve seçim korunuyor; başka ay kümesinde o açıklama taşınmıyor.
- [ ] Rapor oluştur'da masaüstünde seçtiğin klasöre 5 Excel kaydediliyor; tam klasör yolu doğru. Aynı ay kümesinde yazdığın açıklama müdürlük ve genel dosyaya taşınıyor.
- [ ] Açıklama yükle ile ornek/aciklama klasöründeki dört müdürlük dosyası seçilince 24 açıklama alınıyor; yanlış ay kümesinde reddediliyor.
- [ ] Doğrulama/revizyonda 26 değişiklik, bölüm ve alt gider adlarıyla görünüyor.
- [ ] Kendi kaynak kopyalarında yeni bölüm/müdürlük ekleyip Kaynakları yenile; listeler, satırlar ve rapor sayıları otomatik artıyor.
- [ ] 2027 kaynaklarıyla kapanışı 2027 seçince bütçe 2027, önceki yıl 2026 oluyor; konumları sağa kayan ciro/bütçe/endeks okunuyor.
- [ ] Pencereyi daraltınca kartlar ve tablolar kullanılabilir; taşan tablolar kendi içinde kayıyor.

Son iki veri senaryosu için tüm kaynakların tarihleri/dosya dönemleri/eşleştirmeleri birbiriyle tutarlı olmalı.
