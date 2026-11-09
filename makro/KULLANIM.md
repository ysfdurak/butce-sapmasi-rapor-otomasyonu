> **V4:** Önce [Önceki yorumları toplu inceleme](V4_YORUM_KULLANIMI.md) belgesini okuyun. Yeni kurulum adı V4, makro modül sayısı 10’dur. Önceki kaynak, rapor ve dashboard kuralları korunur.

> **V3 güncellemesi:** Önce [V3_YENILIKLER.md](V3_YENILIKLER.md) ve [V3_CANLI_KONTROL.md](V3_CANLI_KONTROL.md) dosyalarını okuyun. Aşağıdaki V2 akışı korunur; sürüm dosya adlarında V3 kullanılır. Yeni özellikler ve sınırları yukarıdaki belgede açıklanır.

# Bütçe raporu — bağımsız makro sürümü V2

Bu paket **ilk raporlama motorunun VBA karşılığıdır**. Mevcut Python / dashboard V11 paketinden ayrıdır; onu değiştirmez. Veri okuma, hesaplama, rapor kaydetme ve açıklama birleştirme Windows masaüstü Excel içinde çalışır. Çalıştırmak için Python, pip, openpyxl veya PowerShell gerekmez. İnternet bağlantısı ve veri yükleme işlemi yoktur.

## Önemli durum

VBA kaynak kodu, kurulum aracı, ayar Excel'i ve sentetik kaynaklar hazırdır. Bu ortamda Windows Excel bulunmadığından **VBA derleme, düğme/listbox etkileşimi ve gerçek dosya üretimi henüz Windows Excel üzerinde doğrulanmadı**. Bu paket canlı kullanım öncesi sentetik pilot sürümdür. Sentetik beklenen sonuçlar mevcut hesap motorundan doğrulandı; bu, VBA'nın çalıştırılmış olduğu anlamına gelmez.

Paketin içinde önceden derlenmiş, hazır `.xlsm` bulunmaz. `Kurulum.vbs`, gerçek Excel'i kullanarak `Kontrol_Sablonu.xlsx` ve on VBA modülünden **Butce_Kontrol_Makro_V4.xlsm** oluşturur. Şablonu doğrudan açınca çalışan makro veya listbox beklemeyin.

## 1. İlk kurulum

1. ZIP'i yeni bir klasöre **tamamen çıkarın**. ZIP'in içinden çalıştırmayın.
2. `Kurulum.vbs` dosyasına çift tıklayın.
3. Windows masaüstü Excel açılır; modüller eklenir ve aynı klasöre `Butce_Kontrol_Makro_V4.xlsm` kaydedilir. Ay listesi ve düğmeler oluşturulur.
4. Sonraki açılışlarda yalnız bu `.xlsm` dosyasını açmanız yeterli olur. Kurulum, var olan kontrol dosyasını değiştirmez.

Kurumsal makro politikası geçerlidir. Kurulum güvenlik ayarı değiştirmez. VBA proje erişimi veya Windows betikleri şirket politikasınca engellenirse aşağıdaki elle kurulum kullanılabilir; makro çalıştırmanın da engelli olduğu ortamda BT desteği gerekir.

### Elle kurulum — VBS veya VBA proje erişimi engelliyse

1. `Kontrol_Sablonu.xlsx` dosyasını Excel'de açın; **Farklı Kaydet → Excel Makro İçerebilen Çalışma Kitabı (.xlsm)** ile `Butce_Kontrol_Makro_V4.xlsm` olarak aynı klasöre kaydedin.
2. `Alt+F11` ile VBA editörünü açın. Bu çalışma kitabının projesini seçin.
3. **File → Import File** yoluyla `vba_elle` klasöründeki on `.bas` modülünün tamamını içeri alın. Bu kopyalar Türkçe Windows (1254) içindir. Türkçe harfler bozulursa UTF-8 `vba` kopyalarını bir metin editöründe açıp metni yeni standart modüllere yapıştırın; `Attribute VB_Name` satırını yapıştırmayın ve modül adlarını dosya adlarıyla aynı yapın.
4. **Debug → Compile VBAProject** seçin. Hata varsa hata metnini ve işaretlenen satırı kaydedin.
5. Excel'e dönün; `Alt+F8 → PanelHazirla → Çalıştır`. Kaydedin.

Bu yöntem programatik VBA proje erişimi gerektirmez. İmzalı dağıtım gerekiyorsa BT'nin mevcut imzalama süreci kullanılmalıdır.

## Kaynak yollarını bir kez seçme ve değiştirme

**Ayar dosyası seç** ile kendi Ayarlar Excel'ini seç. Sonra **Kaynaklar ve eşleştirmeler** düğmesi mevcut **Kaynak Önizleme** sayfasını açar.

- Üstte maliyet klasörü, ciro/tutar kontrol dosyası ve endeks dosyası için ayrı seçme düğmeleri vardır. İstersen tam yolu hücreye de yazabilirsin.
- **Konumları kaydet** yeni kaynakların okunabildiğini kontrol eder, önceki ayarı `Ayar_Yedekleri` klasörüne kopyalar, yalnız üç yol hücresini değiştirir ve eşleştirmeleri tekrar okur. Ayar Excel'i işlem sırasında değişirse üzerine yazmaz. Excel'de açık/kilitliyse kapatıp yeniden dene.
- Ayar yolu, çıktı yolu ve ay listesi kontrol `.xlsm` dosyasında saklanır. **Seçimleri kaydet** veya Ctrl+S ile kaydet. Kaynak seçimi ve başarılı rapor üretiminde de kontrol dosyası kaydedilmeye çalışılır; yazılamazsa ekranda bilgi verilir.
- Sonraki ay aynı klasöre yeni dosyaları ekleyip kapanışı değiştir; **Kaynakları yenile** tüm kayıtları tekrar okur. Yollar değişmediyse yeniden seçmek gerekmez.
- Bu sayfada bölüm–müdürlük–direktörlük–müdürlük altı–lokasyon–kontrol adı–endeks adı alanlarının tamamı, ana gider grupları, gider birleştirme kayıtları ve dışlanan sayfalar gösterilir. Sabit bir bölüm sayısı yoktur.
- Yeni bölümün kaynak sayfasını ve ayar eşleştirmesini ekleyip yenileyince otomatik dahil olur. Silinen veya adı değişen eşleştirmeler de yeniden okunur. Eşleştirmesiz kaynak bölüm adıyla uyarılır; müdürlük/lokasyon/endeks ilişkisi tahmin edilmez. Ayarda olup kaynağı olmayan bölüm de eksik kaynak uyarısıyla görünür.

**Kaynak başlıkları:** İlk dört personel satırının ve beşinci işlem adedi satırının adları, gider adı sütunundan (örnekte BE60:BE64) okunur. `Maliyet / gosterge_adi_sutunu` ayarı verilirse o sütun numarası kullanılır. Kaynaklar önizlemesinde bölüm bazında bu beş gerçek ad görünür. Kontrol ve Revizyon sayfasının sonundaki adet tablosunda aynı adlarla önceki yıl, gerçekleşen ve bütçe adetleri bulunur; parasal gider toplamına girmez. Birim maliyet işlem adedi başlığı da kaynaktan gelir. Kaynak adları bölüm bazında farklıysa her bölüm kendi adıyla gösterilir; birleşik başlıkta farklı adlar birlikte yazılır.

Bu güncelleme bağımsız **makro rapor motoruna** uygulanmıştır. Mevcut yedi sayfalık Excel dashboard'unu bu paket sıfırdan oluşturan bir üreticiye dönüştürmez; Python dashboard ayrı sürümdür.

V1'den geçişte ZIP'i yeni klasöre çıkar ve yeni `Butce_Kontrol_Makro_V4.xlsm` dosyasını oluştur. Kendi mevcut Ayarlar Excel'ini yeniden seçebilirsin; maliyet ve açıklama kaynakların değiştirilmez. Eski `.xlsm` içine yeni/eskileri karıştırarak modül ekleme.

## 2. İlk deneme

Önce **Sentetik hesap testi** düğmesine basın. Bu düğme panelde seçtiğiniz gerçek ayar dosyasını kullanmaz; paketin `ornek/Ayarlar.xlsx` verisiyle kontrollü hesap yapar. Paneldeki ayarları değiştirmez. Sonuç `Test Sonuçları` sayfasında görünür. Bu test rapor dosyası üretimini ve görsel yerleşimi sınamaz.

Sonra kontrol panelinde:

| Alan | Sentetik başlangıç değeri |
|---|---|
| Ayar dosyası | ornek\Ayarlar.xlsx |
| Kaynak kapanış yılı | 2026 |
| Kaynak kapanış ayı | 10 |
| Rapor yılı | 2026 |
| Ay listesi | Ocak–Ekim seçili |
| Çıktı klasörü | Klasör seç düğmesiyle istediğiniz konum |

**Kaynakları önizle**, dosya adındaki yıl/ay bilgisini ve dosyanın seçim sonucunu gösterir. **Rapor oluştur**, kaynakları okuyup hesaplar ve çıktıları seçtiğiniz klasörün altında `Raporlar_yyyymmdd_hhnnss` adlı yeni bir alt klasöre kaydeder. Tam yol panelde ve bitiş mesajında görünür. Eski çıktılar değiştirilmez. Yarım kalan bu çalışmanın geçici dosyaları kaldırılır.

Örnekte her kapanış için 5 maliyet Excel'i, bu dosyalarda 1/2/3/3/3 bölüm; toplam 12 bölüm, 4 müdürlük ve tek direktörlük vardır. Eylül ve Ekim kapanışları revizyon kontrolü için birlikte gelir. Çıktı **5 Excel** olur: bir direktörlük geneli ve dört müdürlük dosyası. Bölüm/müdürlük sayısı ayarlardan okunur; kodda sabit değildir.

Ay listesi Ctrl gerektirmeden tıklayarak çoklu seçim yapar. Örneğin Ocak, Mart, Mayıs seçerseniz yalnız bu üç ay toplanır; önceki yılda da aynı üç ay kullanılır. Rapor dosyaları seçimin anlık çıktısıdır. **Seçimi kontrol panelinde değiştirip yeniden Rapor oluştur'a basın.** Bu varyantta eski raporun kendi içinde formüllü seçim paneli yerine merkezi makro kontrol paneli kullanılır. 7 sayfalık etkileşimli dashboard bu pakette yeniden üretilmez; mevcut V11 dashboard sürümünde kalır.

## 3. Beklenen tutarlar

2026 Ocak–Ekim, Ekim kapanışı, tüm bölümler:

| Kontrol | Beklenen sonuç |
|---|---:|
| Gerçekleşen gider | 3.449.383.390,31 TL |
| Bütçe gider | 3.440.938.980,13 TL |
| Bütçe − gerçekleşen | −8.444.410,18 TL |
| Karşılaştırılabilir bölüm | 12 |
| Direktörlük ağırlıklı reel iyileşme | yaklaşık −%1,52 |
| Geçmiş gider/bütçe revizyon kaydı | 26 |
| Dip toplam kontrol kaydı | 110; hepsi Doğrulandı |

Ocak, Mart, Mayıs birlikte seçildiğinde gerçekleşen **913.186.126,67 TL**, bütçe **1.094.635.989,61 TL** olur. Böylece kesintisiz kümülatif yerine tam seçili ay kümesinin kullanıldığı kontrol edilebilir. Ayrıntılar `Beklenen_Sonuclar.json` ve `Sentetik_Beklenen_Sonuclar.xlsx` içindedir. Bunlar karşılaştırma referansıdır; VBA çıktısı değildir.

## 4. Ayarlar ve hesap kuralları

Eşleştirmeler yine `ornek/Ayarlar.xlsx` içinde düzenlenir:

| Sayfa | İçerik |
|---|---|
| Bölüm Eşleştirme | Bölüm, müdürlük, direktörlük, endeks adı, lokasyon, kontrol bölümü, müdürlük altı |
| Gider Grupları | Alt gider indeks aralıkları ve ana gider adları |
| Gider Eşleştirme | Birden fazla kaynak giderin tek rapor kalemine bağlanması |
| Ortak Sayfalar | Okunmayacak iki ortak sayfa |
| Dosya ve Konum Ayarları | Kaynak yolları, hücre/satır konumları ve tolerans |

- Maliyet gerçekleşen tarihleri BF59'dan sağa okunur. Bütçe, EZ58 ipucundan sağa doğru **Bütçe** etiketiyle ayrılır; altında kapanış yılının tek 12 aylık bloğu bulunur. Aynı tarihin gerçekleşende ve bütçede bulunması karışıklık yaratmaz. Daha yeni yılda etiket sağa kayabilir.
- 60–63 personel sayıları, 64 işlem adedidir. Giderler 65'ten itibaren 27 satırdır. Tutar hesaplarına ilk beş adet satırı dahil edilmez. Personel gideri adı taşıyan parasal gider kalemleri gider hesabına dahildir.
- Son kapanış dosyası geçmiş ayları güncel haliyle sağlar. Önceki kapanışlar, aynı yılın geriye dönük alt gider revizyonlarını ve bütçe değişikliklerini yakalamak içindir. Eksik ardışık kaynaklar uyarılır. Eylül–Ekim örneği bu karşılaştırmayı gösterir; Ocak–Ağustos dosyaları olmadığı için o geçişler için uyarı beklenir.
- Giriş ve Tartım'ın toplamı, kontrol dosyasında Ürün Kabul Toplamı ile kıyaslanır. Eşleşmeyen bölüm, eksik dönem, mükerrer bölüm/kapanış ve toleransı aşan dip toplam farkı raporlanır. Uyarı varken oluşan raporun onaylı/doğrulanmış olduğu varsayılmaz.
- Kontrol/ciro: D18'den gerçekleşen dönemler; 19. satır ciro. CW18 ipucunun sağında, gerçekleşenden sonraki tek kapanış yılı bütçe bloğu aranır. Yıl değişimindeki 12 sütun kayması desteklenir. Şirket cirosu bölümlerle çarpılmaz. Müdürlüklerin kendi raporunda da aynı şirket cirosu paydası kullanılır.
- Endeksin ikinci tablosu C17'den başlar. D17'den ardışık tarihler ilk boşluğa kadar okunur; sağdaki iki hazır oran sütunu kullanılmaz. Aylık endeks = bu yıl / önceki yıl aynı ay. Seçili aylarda bu oranların aritmetik ortalaması alınır.
- Birim maliyet = seçili ay gider toplamı / seçili ay pozitif işlem adetleri toplamı. Önceki yıl BM × dönem endeksi, endekslenmiş BM'dir. Reel iyileşme = (endekslenmiş önceki BM − bu yıl BM) / endekslenmiş önceki BM. Karşılaştırılabilir ve endeksle eşleşmiş bölümler gösterilir. Güncel gider ağırlıkları müdürlük ve direktörlük için ayrı normalize edilir; grup iyileşmesi ağırlıklı toplamdır.
- Eksik tutar sıfır sayılmaz. Toplamın bir parçası eksikse toplam boş kalır. Sıfır/negatif işlem adedi BM hesabına alınmaz. Eski yıl raporunda bütçe boş kalır. Bütçe yılı kaynak kapanışından gelir; kodda 2026 yalnız sentetik testte sabittir.

Bu varyant güncel konumlu tablo yapısını ve tarih içeren açık gerçekleşen/bütçe aralıklarını destekler. Tarih içermeyen eski serbest blok keşfi veya eski üç sütunlu endeks listesi bu varyantta kullanılmaz; eski Python sürümü ayrıca korunmuştur.

## 5. Raporlar ve açıklamalar

Her çıktıda dört görünür sayfa vardır:

1. **Genel Özet:** Bölümler, ait oldukları müdürlük bitince müdürlük toplamı; genelde en altta direktörlük toplamı. Önceki yıl, gerçekleşen, bütçe, bütçe farkı, aşım oranı, OPEX/şirket cirosu. OPEX oranı için önceki yıl verisi olan bölüm kapsamı kullanılır.
2. **Giderler:** Ana giderlerin altında genişletilebilir alt giderler; bölümler yatay bloklar halinde. Bütçe − gerçekleşen negatifse kırmızı, pozitifse yeşil. Aşım yüzdesi (gerçekleşen − bütçe)/bütçe olduğundan pozitifse kırmızıdır. Ana gider alt toplamları kendi grubunun altında, genel toplam en altta. Müdürlüklerde beşinci alan Açıklama.
3. **Birim Maliyet:** Karşılaştırılabilir bölümler, müdürlük ve genel toplam düzeni; gider ağırlıkları ve reel iyileşme.
4. **Kontrol ve Revizyon:** Bölüm bazlı dip toplamlar, hangi alt giderde hangi önceki ayın değiştiği, önce/sonra/fark ve kaynak dosyalar, ciro ve uyarılar. Sonunda ilk beş adet göstergesi kaynak adlarıyla ve ondalıksız sayılarla gösterilir.

Müdürlük raporlarının **Giderler** sayfasında Açıklama sütununa sabit metin girip kaydedin. Kontrol panelinde **Açıklamaları birleştir** ile önce aynı dönemin genel raporunu, sonra o rapordaki tüm müdürlükleri kapsayan dosyaları seçin. Dönem/kapanış/ay kümesi, gider kodları ve gerçekleşen/bütçe tutarları kontrol edilir. Kaynaklar değiştirilmeden yeni bir genel dosya kaydedilir. Alt Gider Açıklamaları sayfasında her kalem tek hücreye şu düzende gelir:

```
Giriş: Açıklama metni
Tartım: Diğer açıklama
```

## 6. Windows canlı kabul kontrolü

- [ ] VBA editöründe Compile VBAProject hatasız; on modül yüklü.
- [ ] Kaynakları seçip Konumları kaydet çalıştır; üç yol ayar Excel'inde güncellenmiş, diğer eşleştirmeler aynı ve yedek dosyası oluşmuş.
- [ ] Kontrol dosyasını kapat/aç; ayar yolu, çıktı yolu ve çoklu ay seçimi korunmuş.
- [ ] Kaynak önizlemesinde personel/işlem adları gerçek kaynak hücreleriyle aynı; kontrol raporunun sonundaki gösterge tablosunda da bunlar var.
- [ ] Bir eşleştirme ekle/sil/adını değiştir; yenileyince tüm kayıtlar ve müdürlük dosya sayıları yeni yapıdan geliyor.
- [ ] Sentetik hesap testi tüm kontrolleri geçiriyor.
- [ ] Ocak–Ekim seçili Rapor oluştur: seçilen klasör altında beş çıktı var; tutarlar yukarıdaki referansla eşleşiyor.
- [ ] Ocak/Mart/Mayıs seçiminde tam bu ayların toplamları ve önceki yıl aynı aylar hesaplanıyor.
- [ ] 2025 rapor yılında bütçe alanları boş; 2026 kapanışından sonra ay seçilince hata açık.
- [ ] Giderlerde eksi işareti, iki ondalık, kırmızı/yeşil koşullu biçim ve bölüm dış kenarlıkları görünür.
- [ ] Alt gider satırları +/− ile açılıyor; ana gider toplamları altta.
- [ ] Kontrol sayfasında 26 revizyonun müdürlük/bölüm/alt gider/dönem bilgileri var.
- [ ] Bir müdürlük dosyasında not girip tüm müdürlükleri yükleyince yeni genel dosyada bölüm adıyla birleşiyor.
- [ ] Ayarı/sentetik kaynağı kopyalayarak yeni bölüm veya müdürlük ekleyince çıktı sayısı ve kapsamı değişiyor.
- [ ] 2027 kaynaklarıyla başlık ve bütçe yılı 2027'ye kayıyor; sağa kayan bütçe/endeks başlıkları doğru okunuyor.

Son iki madde için gerçek 2027 verisi gerekmez; kaynakların kopyalarında tarihleri ve dosya dönemlerini tutarlı kaydırabilirsiniz. Bu Windows kontrolleri burada tamamlanmış sayılmamıştır.
