# V4 — Önceki ayın yorumlarını topluca inceleme

Bu sürüm önceki rapor kapanışının yorumlarını kontrol bekleyen taslaklar olarak getirir. Yeni gider nedeni veya fatura yorumu uydurmaz. Fatura detayı çıkarma ikinci aşamada ele alınacaktır.

## Python: günlük akış

1. Kaynakları ve bütçe yılını seçin. Ortak bölüm ve ay seçimleri yorum incelemesinde de geçerlidir.
2. Üstte **Önceki ayın yorumlarını getir** düğmesine basın. Yerel kayıtta onaylı önceki kapanış yorumları varsa getirilir. Yoksa açılan tabloda **Önceki ayın açıklamalı dosyalarını seç** ile müdürlük raporlarını seçin.
3. Önceki yorum ve bu dönemin taslağı yan yanadır. Kaynak kapanışı ve eski ay kümesi gösterilir. Farklı ay seçimi veya tutar değişimi işaretlenir. Tutarların aynı olması otomatik onay değildir.
4. Aynı kalacak iki satırı seçip **Seçileni onayla** düğmesine bir kez basabilirsiniz.
5. Diğer sekiz satırın taslaklarını aynı tablodaki metin kutularında düzenleyin; yazdığınız satırlar otomatik seçilir. **Seçileni onayla** ile birlikte kaydedin. Başka satırları onaylamak, henüz kaydetmediğiniz diğer taslakları silmez.
6. Aynı açıklama birkaç kayıt için geçerliyse satırları seçin, **Seçili satırlara ortak açıklama uygula** alanına metni yazın. **Metni seçili taslaklara uygula** yalnız taslakları doldurur; onay ayrı yapılır.
7. **Düzenlenen taslakları kaydet** onay vermeden metni saklar. **Seçileni bu dönem geçersiz say** açıklamayı rapordan çıkarır; önceki metin ve karar kaydı durur.
8. Bekleyen/onaylı/geçersiz/verisi değişen/tümü görünümleri ve bölüm/gider/metin araması kullanılabilir. Toplu onay ve geçersiz sayma yalnız görünür, seçili satırlara uygulanır.

Bu dönemde zaten girilmiş açıklamalar korunur; geçmiş dosyayı yeniden getirmek onları veya düzenlediğiniz taslakları değiştirmez. İki pencereden aynı kayıt değiştirilirse eski pencere güncel kaydı ezemez. Onaylandıktan sonra tutar değişirse yorum yeniden kontrol bekleyen hale gelir.

## Makro: günlük akış

Kurulum.vbs V4 xlsm dosyasını oluşturur. Elle kurulumda **10 modülün tamamını** (MakroYorum dahil) yükleyin. PanelHazirla ile Çalışma Merkezi açılır.

- **Önceki ayın yorumlarını getir:** kontrol dosyasına kaydedilmiş önceki kapanış yorumlarını getirir.
- **Açıklamaları toplu düzenle:** Yorum İnceleme tablosunu açar. Dosyalardan ilk yükleme için **Önceki dosyaları seç** düğmesini kullanın.
- A sütununa `Evet` yazarak satırları seçin. **F: önceki yorum**, **G: bu dönem taslağı**. G sütununda doğrudan düzenleme yapın. H/I sütunları kaynak kapanışı ve eski ay seçimidir.
- İki satırı seçip onaylayın. Bekleyenleri göster düğmesiyle kalan sekiz satıra geçin; G sütununda düzenleyip Görünürleri seç ve Seçileni onayla ile topluca kaydedin.
- Ortak metni B3 alanına yazıp seçili satırlara uygulayın. Bu işlem onay vermez. Taslakları kaydet düğmesi değişen taslakları onaysız saklar.
- Excel çıktısına elle yazdığınız **bu dönem** açıklamalarını korumak için önce **Bu dönem yorumlarını yükle** düğmesiyle aynı kapanış ve aynı ay kümesindeki müdürlük raporlarını alın. Uygulamanın okumadığı dış dosyadaki elle yazılmış metni kontrol dosyası kendiliğinden bilemez.
- Kaydedilmemiş G sütunu değişiklikleri varken rapor üretimi durur. Kontrol dosyasını kaydedin; geçmiş yorumlar ve inceleme kararları bu xlsm içinde tutulur.

## Raporda görünüm

- Bekleyen yorum kırmızı ve `[Kontrol bekliyor — önceki dönemden]` etiketiyle, kaynak kapanışı ve ay kümesiyle gösterilir.
- Onaylanan açıklama normal renktedir. Geçersiz sayılan açıklama rapora alınmaz.
- İlgili müdürlük dosyasında bölüm/alt gider açıklama hücresine yazılır. Genel dosyada aynı gider için `Bölüm A: ...` biçiminde birleştirilir.
- Açıklama bulunan raporlara **Açıklama Durumu** sayfası eklenir. İlgisiz müdürlüğün yorumları dosyaya taşınmaz; boş durum sayfası oluşturulmaz.
- Genel dosyada aynı hücrede bekleyen ve onaylı bölüm metinleri birlikteyse hücre kırmızıdır; etiket ve durum sayfası hangi bölümün beklediğini açıklar.
- Bekleyen etiketli bir çıktı, Python'daki normal açıklama yükleme yoluyla onaylı metne dönüştürülemez; inceleme yolu kullanılmalıdır.

## 10 yorumluk hazır test

Her iki pakette `ornek/onceki_yorumlar` klasörü vardır:

- `Onceki_2026_09_Mudurluk_01.xlsx`: 5 Eylül yorumu.
- `Onceki_2026_09_Mudurluk_02.xlsx`: 5 Eylül yorumu.
- `10_Yorum_Beklenen_Senaryo.xlsx`: hangi iki yorumun aynı kalacağı ve hangi sekizinin nasıl değiştirileceği.
- `beklenen.json`: otomatik testlerin kullandığı aynı beklenen metinler.

Örnek Ayarlar.xlsx ile **Ekim 2026** kapanışını açın, **Ocak–Ekim** aylarını seçin. İki Eylül dosyasını birlikte yükleyin. 10 kayıt bekleyen olmalı; eski ay seçiminin Ocak–Eylül olduğu görünmelidir. İki kaydı aynen, sekiz kaydı beklenen çalışma dosyasındaki metinlerle düzenleyerek onaylayın. Sonuç 10 onaylı yorumdur. Dosyaları tekrar getirin: mevcut 10 kayıt korunmalıdır.

Ayrıca iki satırı geçersiz sayma, taslakları onaysız kaydetme, farklı ay seçimi ve mevcut bu dönem açıklamasını koruma durumlarını deneyin. Ayrı durum denemelerinde başka ay kümesi veya temiz örnek kullanıcı kaydı kullanın; kişisel kayıtlarınızı silmeyin.

## Test sınırı

Python hesap/servis testleri ve JavaScript etkileşim/API simülasyonları otomatik çalıştırılmıştır. Gerçek tarayıcı düzeni ve Windows dosya pencereleri burada canlı test edilmemiştir. VBA kaynak/paket kontrolleri yapılmıştır; Windows Excel'de derleme ve makroların çalıştırılması henüz doğrulanmamıştır. `Dogrulama_Durumu.json` ayrıntıları içerir.
