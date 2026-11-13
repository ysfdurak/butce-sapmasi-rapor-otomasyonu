# Makrosuz yol: Power Query + rapor şablonu

Şirket bilgisayarında **Python kurulamıyor ve makrolar kapalıysa** bu yolu kullanın. Yalnız Excel'in yerleşik özellikleri kullanılır: Power Query ve formüller (Excel 2016, 2019, 2021, Microsoft 365). Lisans, eklenti, makro veya kurulum gerekmez. Formüllerde yeni sürümlere özgü işlevler (FILTER, XLOOKUP, LET…) kullanılmadı; şablon tüm bu sürümlerde aynı çalışır.

Üç dosya:

| Dosya | Görevi |
| --- | --- |
| [`powerquery/Butce_Rapor_Sablonu.xlsx`](../powerquery/Butce_Rapor_Sablonu.xlsx) | Raporlar ve panolar (formüllerle). Kontrol Paneli'ndeki seçimler değişince her şey anında güncellenir. |
| [`powerquery/Butce_Veri.pq`](../powerquery/Butce_Veri.pq) | Ana sorgu. Ayarlar, maliyet, kontrol ve endeks dosyalarını okur; şablonun **Veri** sayfasına tek tablo yükler. Müdürlük dosyası için yalnız o müdürlüğün verisini yükleyebilir. |
| [`powerquery/Butce_Aciklama.pq`](../powerquery/Butce_Aciklama.pq) | Açıklama sorgusu. Müdürlüklerden dönen dosyalardaki açıklamaları **Açıklama Onayı** sayfasına toplar; onay durumunu ve işlem geçmişini korur. |

## Python / makro sürümüyle karşılaştırma

| | Python / makro | Power Query şablonu |
| --- | --- | --- |
| Yönetici Özeti, Pano (grafikler), Genel Özet, Giderler | ✓ | ✓ |
| Birim maliyet ve endeks, Hacim ve birim maliyet | ✓ | ✓ |
| Önemli sapmalar, Yıl sonu tahmini | ✓ | ✓ |
| Kontrol dosyası mutabakatı (bölüm / birleşik / Opex / ciro) | ✓ | ✓ (Veri Kontrolü) |
| Revizyonlar (ardışık kapanış dosyaları) ve uyarılar | ✓ | ✓ |
| Kapsam: direktörlük veya tek müdürlük, ay seçimi, önceki yıl | ✓ | ✓ (Kontrol Paneli) |
| Müdürlük dosyaları (yalnız kendi verisi) | Tek çalıştırmada hepsi | Müdürlük başına bir kez: kopya + sorguda `MudurlukDosyasi` (Dağıtım sayfası) |
| Açıklama toplama | Dönen dosyalardan, tutar kontrolüyle | ✓ Butce_Aciklama sorgusu; kapanış, ay seçimi ve tutar kontrolü |
| Onay ve onay izi | Onay, kim/ne zaman, silinmeyen günlük | ✓ Durum sütunu; her işlem zaman · kişi · işlem olarak satırın geçmişine eklenir |
| Önceki dönem açıklamalarını taşıma | Taslak olarak, toplu onay | ✓ Durum = **Bu döneme aktar** |
| Girdi arşivi | Girdi dosyalarının kopyası | Girdi Kaydı (dosya adı, boyut, değiştirilme zamanı) + aylık **Farklı Kaydet**; Excel dosya kopyalayamaz |

Yalnız bu yolda olanlar: **hızlı mod**, **revize bütçe** karşılaştırması (Bütçe Sürümleri) ve **Trend** (üst üste aleyhte giden kalemler). Şirkette ilk deneme için: [kontrol listesi](PQ_DENEME.md).

Bilinçli fark: müdürlük dosyasında diğer bölümlerin verisi bulunmadığından Birim Maliyet'teki **Direktörlük gider ağırlığı** boş kalır (direktörlük raporunda vardır).

Hesap kuralları aynıdır: bir ay boşsa toplam boş kalır (sıfır sayılmaz), Lehte/Aleyhte yönü, önemlilik eşiği (tutar **ve** oran), endeksli önceki yıl karşılaştırması, kontrol toleransı.

## Kurulum (bir kez, yaklaşık 5 dakika)

1. `Butce_Rapor_Sablonu.xlsx`'i istediğiniz klasöre kopyalayıp açın. Uyarı çubuğu çıkarsa **Düzenlemeyi Etkinleştir**.
2. **Veri › Veri Al › Diğer Kaynaklardan › Boş Sorgu**. Power Query düzenleyicisi açılır.
3. **Giriş › Gelişmiş Düzenleyici**. İçindeki metni silin, `Butce_Veri.pq` dosyasının tamamını yapıştırın.
4. En üstteki iki satırı düzenleyin:
   - `AyarDosyasi`: Ayarlar.xlsx'in tam yolu, ör. `"C:\Butce\ornek\Ayarlar.xlsx"`. Dosya Gezgini'nde dosyaya Shift + sağ tık › **Yol olarak kopyala**.
   - `KapanisDonemi`: boş = klasördeki en son kapanış; belirli bir ay için `"2026-09"`.
5. **Bitti**. Sağdaki **Ad** kutusuna `Butce_Veri` yazın.
6. "Gizlilik düzeyleri" sorulursa tüm kaynaklar için **Kuruluş** seçin.
7. **Giriş › Kapat ve Yükle › Kapat ve Şuraya Yükle…** › **Tablo** › **Var olan çalışma sayfası** › `=Veri!$A$1` › **Tamam**.
8. **Kontrol Paneli**'ne dönün. "Veri satırı" sıfırdan büyükse kurulum tamam. Dosyayı kaydedin.

Daha önce boş bir dosyaya kurduğunuz eski `Butce_Veri` sorgusu bu şablonla çalışmaz; şablona yeni sorguyu yapıştırın.

## Kullanım

| Ne zaman | Ne yapılır |
| --- | --- |
| Yeni kapanış dosyaları klasöre konduğunda | **Veri › Tümünü Yenile** (tek tuş). Tüm sayfalar ve grafikler güncellenir. |
| Kapsam, ay veya eşik değiştirmek için | Kontrol Paneli'ndeki sarı hücreler. Yenileme gerekmez, anında hesaplanır. |
| Müdürlüğe ayrı dosya göndermek için | Kapsam'da müdürlüğü seçin › **Dosya › Farklı Kaydet**. |
| Açıklama eklemek için | **Açıklamalar** sayfası: bölüm + gider kodu + açıklama. Giderler ve Önemli Sapmalar'da görünür; sağdaki liste açıklama bekleyen önemli sapmaları gösterir. |

Kontrol Paneli seçimleri:

- **Rapor yılı:** boş = kapanış yılı. Önceki bir yıl seçilirse bütçe karşılaştırması yapılmaz.
- **Kapsam:** `Tümü` = direktörlük raporu; bir müdürlük seçilirse yalnız o müdürlük.
- **Ay seçimi:** Yıl başından kapanışa · Yalnız kapanış ayı · Elle seçim (Elle seçimde ayların altına `1`).
- **Önemlilik tutarı / oranı, Özet kalem sayısı:** boş bırakılırsa Ayarlar.xlsx'teki değerler.

## Hızlı mod

Revizyonları bulmak için sorgu kapanış yılının bütün kapanış dosyalarını okur. Klasörde çok ay varsa yenileme uzayabilir. `Butce_Veri` sorgusunda `HizliMod = true` yapılırsa yalnız kapanış ve önceki ayın dosyaları okunur; revizyonlar ve "ardışık dosya eksik" uyarısı yalnız son ay için hesaplanır. Diğer bütün raporlar aynıdır. Kontrol Paneli › Son yenileme satırında "hızlı mod" yazar.

## Revize bütçe

Kaynak maliyet sayfasında **Bütçe** başlığıyla aynı satırda `Revize Bütçe` başlığı ve altında kapanış yılının 12 ayı varsa sorgu onu da okur (etiket farklıysa Ayarlar › Dosya ve Konum Ayarları'na `Maliyet · revize_butce_etiketi` satırı ekleyin). Gider eşleştirmesi ve "bir ay boşsa toplam boş" kuralı aynen uygulanır.

- **Kontrol Paneli › Karşılaştırma bütçesi:** `Revize bütçe` seçilirse Genel Özet, Giderler, Önemli Sapmalar, Yıl Sonu Tahmini, Pano ve Trend revize bütçeye göre hesaplanır (bütçe işlem adedi ve bütçe cirosu orijinal kalır). Yönetici Özeti'nin başında "Karşılaştırma bütçesi: revize bütçe" yazar.
- **Bütçe Sürümleri** sayfası: gerçekleşen, bütçe, revize bütçe, revize − bütçe, gerçekleşen − bütçe, gerçekleşen − revize ve revize bütçeye göre sapma oranı (bölüm, müdürlük, direktörlük).
- Kaynakta revize blok yoksa sayfa bunu yazar; seçim yapılırsa Kontrol Paneli uyarır.

## Trend

**Trend** sayfası iki şey gösterir (karşılaştırma bütçesine göre, yalnız bütçe yılında):

1. Her bölüm, müdürlük ve direktörlük için ay ay gerçekleşen − bütçe ve kapsam toplamının grafiği.
2. **Üst üste aleyhte giden kalemler:** kapanış ayında biten ve her ay *aylık eşiği* aşan aleyhte sapma serisi en az *N* ay olan kalemler; seri uzunluğu ve seri toplamına göre sıralı. Aylık eşik boşsa önemlilik tutarı / 12, N boşsa 3 (sayfanın üstündeki sarı hücreler). Yönetici Özeti'nde sayısı görünür.

Tek seferlik sapmayı kalıcı olandan ayırmak için kullanın: 3+ ay süren seri, bütçenin veya kaynak tahsisinin gözden geçirilmesi gerektiğini gösterir.

## Müdürlük dosyaları (dağıtım)

Müdürlük dosyası yalnız o müdürlüğün bölümlerini içerir; diğer bölümlerin tutarları dosyaya hiç girmez. **Dağıtım** sayfası müdürlükleri, önerilen dosya adlarını ve sorguya yazılacak satırı listeler. Her müdürlük için bir kez:

1. Ana raporu **Dosya › Farklı Kaydet** ile önerilen adla kaydedin (ör. `Depolama_2026-10.xlsx`).
2. Yeni dosyada **Veri › Sorgular ve Bağlantılar** › `Butce_Veri` sağ tık › **Düzenle** › **Gelişmiş Düzenleyici**.
3. `MudurlukDosyasi = "Depolama",` yazın › **Bitti** › **Kapat ve Yükle**. Kontrol Paneli'nde "Müdürlük dosyası: Depolama" görünür.
4. Kaydedip müdürlüğe gönderin. Müdürlük açıklamalarını **Açıklamalar** sayfasına yazar, dosyayı kaydedip geri gönderir. Müdürlüğün **Tümünü Yenile**'ye basması gerekmez.

Sonraki aylarda aynı müdürlük dosyasını açıp **Tümünü Yenile** yapmanız yeterli (sorgu müdürlük ayarını hatırlar).

## Açıklama toplama ve onay

Bir kez kurulum (ana raporda):

1. Ayarlar.xlsx'in yanında `pq_aciklama` klasörü açın (başka yer için sorgudaki `AciklamaKlasoru` satırı).
2. **Veri › Veri Al › Boş Sorgu › Gelişmiş Düzenleyici** › `Butce_Aciklama.pq`'yi yapıştırın, `AyarDosyasi` satırını düzenleyin.
3. Sorgunun adı `Butce_Aciklama` › **Kapat ve Şuraya Yükle… › Tablo › Var olan çalışma sayfası** › `='Açıklama Onayı'!$A$1`.
4. Kontrol Paneli › **Onaylayan (adınız)** hücresine adınızı yazın.

Her ay:

1. Müdürlüklerden dönen dosyaları `pq_aciklama` klasörüne koyun › **Veri › Tümünü Yenile**.
2. **Açıklama Onayı** sayfasında her satır için **Kontrol** sütununa bakın:
   - `Güncel`: açıklama bu kapanış ve ay seçimi için yazılmış, tutarlar değişmemiş.
   - `Tutar değişti`: açıklama yazıldıktan sonra tutar değişmiş (revizyon); müdürlükten güncel açıklama isteyin.
   - `Önceki kapanış (…)` / `Ay seçimi farklı`: başka dönem için yazılmış. Uygunsa **Bu döneme aktar**.
   - `Bu bölümde bu kod yok`, `Çakışma…`, `Dosya okunamadı…`: dosyayı düzeltin.
3. **Durum** sütununda seçin: `Onaylandı`, `Reddedildi`, `Bekliyor` veya `Bu döneme aktar`; gerekirse **Onay notu** yazın.
4. **Tümünü Yenile**. İşlem, adınız ve zamanla **İşlem geçmişi** sütununa eklenir; Giderler ve Önemli Sapmalar güncellenir.

Kurallar:

- Rapor yalnız `Onaylandı` + `Güncel` açıklamaları gösterir. `Bekliyor` olanlar `[Onay bekliyor]` önekiyle görünür ve Önemli Sapmalar'da durum **Onay bekliyor** olur.
- Ana raporun kendi Açıklamalar sayfasına yazılan açıklama (finans) her zaman önceliklidir.
- Müdürlük açıklama metnini değiştirip yeniden gönderirse durum `Bekliyor`'a döner ve geçmişe yazılır.
- Kaynak dosyası klasörden kaldırılan açıklama tablodan silinmez (Kaynakta = Yok); onay izi korunur.
- Sorgu kendi önceki tablosunu okuyarak durumları korur. Tablo satırlarını elle silmeyin; **Kayıtlı durum** sütununu değiştirmeyin.
- İlk yenilemede gizlilik düzeyi sorulursa **Kuruluş** seçin (Geçerli Çalışma Kitabı dahil).

## Girdi kaydı ve arşiv

**Girdi Kaydı** sayfası her yenilemede raporun okuduğu dosyaları listeler: tür, dönem, boyut, değiştirilme zamanı ve durum (Okundu / Kapsam dışı / Dönem bulunamadı / Bulunamadı). Kontrol Paneli'nde **Son yenileme** zamanı görünür. Excel dosya kopyalayamadığı için girdi arşivi şöyle tutulur: kapanış dosyaları klasörde kalır (revizyonlar zaten bunlardan hesaplanır) ve her ayın raporu **Farklı Kaydet** ile ayrı saklanır; o raporun Girdi Kaydı hangi dosya sürümlerinden üretildiğini gösterir.

## Nasıl çalışır

```
Ayarlar.xlsx ─┐
maliyet/*.xlsx├─► Butce_Veri (Power Query) ─► Veri sayfası (tek tablo) ─► formüller ─► raporlar, pano
Kontrol.xlsx  │                                                             ▲
Endeks.xlsx  ─┘                                                Kontrol Paneli seçimleri
```

- Sorgu her satırı bir **bloğa** yazar (`Tablo` sütunu): `Gider` (bölüm × kalem × yıl, 12 ay gerçekleşen G1–G12 ve bütçe B1–B12), `Toplam`, `Adet`, `Kontrol`, `Ciro`, `Endeks`, `Revizyon`, `Uyari` ve rapor düzeni (`Bolum`, `Duzen`, `Kalem`, `Satir`…). Her satırın tekil bir `Anahtar`ı vardır (ör. `Gider|Giriş|5|2026`).
- Şablondaki gizli yardımcı sayfalar (`_H`, `_Bolum`, `_Kalem`…) satırları `Anahtar` ile bulur (INDEX/MATCH), seçili ayları toplar (SUMPRODUCT) ve raporlar bunlardan beslenir.
- Hücre konumları Ayarlar.xlsx'ten okunur. Maliyet sayfası **Bütçe** başlığına, kontrol dosyası ilk ardışık tarih başlığına, endeks dosyası **Bölüm Gider Endeksi** başlığına göre hizalanır; boş üst satırlar/sütunlar sonucu etkilemez.
- Tablonun tam tanımı ve Python karşılığı: [`tools/pq_veri.py`](../tools/pq_veri.py). Şablonu üreten kod: [`tools/pq_sablon.py`](../tools/pq_sablon.py).

## Doğrulama

Her değişiklikte CI üç şeyi otomatik denetler:

1. **Gerçek M motoru:** İki sorgu da Excel'deki Power Query ile aynı motorla (Microsoft Power Query SDK, PQTest) örnek verilerle çalıştırılır ve Python benzetimiyle satır satır karşılaştırılır: Butce_Veri için en son kapanış, 2026-09, Depolama müdürlük dosyası, hızlı mod ve revize bütçeli kaynak; Butce_Aciklama için üç ardışık yenileme (yeni açıklamalar → onay, ret, aktarma, metni değişen ve kaynağı kaldırılan satırlar → değişiklik yokken geçmişe bir şey yazılmaması).
2. **Şablon formülleri:** Şablon bu veriyle LibreOffice'te hesaplanır; Genel Özet, Giderler, Birim Maliyet, Önemli Sapmalar, Yıl Sonu Tahmini, Hacim ve Uyarılar Python motorunun raporlarıyla karşılaştırılır (direktörlük yıl başından, yalnız kapanış ayı, Depolama · Ocak/Mart/Mayıs, Depolama müdürlük dosyası). Açıklama Onayı'nın kontrol sütunu ve onaylı açıklamaların rapora yansıması, revize bütçe seçimi, Bütçe Sürümleri ve Trend listesi de veriden hesaplanan beklenen değerlerle denetlenir.
3. **Birim testleri:** blok sayıları, tekil anahtarlar, boş ay kuralı.

Elle kontrol: örnek verilerle Ekim 2026, Ocak–Ekim, Tümü: gerçekleşen **3.449.383.390,31 TL**, bütçe **3.440.938.980,13 TL**.

## Sorun giderme

| Belirti | Neden / çözüm |
| --- | --- |
| Kontrol Paneli: "Veri yok" | Sorgu `Veri!$A$1`'e yüklenmedi. **Veri › Sorgular ve Bağlantılar** › `Butce_Veri` sağ tık › **Yükle…** › Var olan sayfa `=Veri!$A$1`. |
| `DataFormat.Error` / dosya bulunamadı | `AyarDosyasi` yolu veya Ayarlar'daki `maliyet_klasoru`, `kontrol_dosyasi`, `endeks_dosyasi` yanlış. |
| Açıklama Onayı'ndaki durumlar yenilemede kayboluyor | Sorgu Açıklama Onayı tablosunu okuyamıyor: tablo `'Açıklama Onayı'!$A$1`'de olmalı; satır ve sütunları silinmemeli. |
| `Formula.Firewall` hatası | **Veri › Veri Al › Sorgu Seçenekleri › Geçerli Çalışma Kitabı › Gizlilik › Gizlilik düzeylerini yoksay**. |
| Uyarılar: `Bütçe başlığı bulunamadı` | Sayfada "Bütçe" başlığı yok. Ayarlar'daki `butce_etiketi` ile sayfayı karşılaştırın. |
| Uyarılar: `tek 12 aylık bütçe bloğu bulunamadı` | Bütçe başlığının altındaki 12 ay dosyanın yılıyla eşleşmiyor veya eksik. |
| Uyarılar: `Gider adı bulunamadı` | `ilk_gider_satiri` / `gider_kalem_sayisi` kaynakla uyuşmuyor; satır eklenmiş olabilir. |
| Uyarılar: `Aynı bölüm ve kapanış için birden fazla kaynak` | Bölüm iki dosyada var; ikisi de alınmaz. Fazlasını kaldırın. |
| Uyarılar: `Kontrol/Endeks dosyası okunamadı` | Dosya adı veya sayfa adı Ayarlar'la uyuşmuyor; diğer raporlar yine üretilir. |
| Yenileme çok uzun sürüyor | Maliyet klasöründe yalnız kapanış dosyaları olsun; eski yılların dosyalarını alt klasöre taşıyın (alt klasörler okunmaz). |

Şablon sınırları: 2.000 açıklama satırı, 60 bölüm, Genel Özet'te 120 satır, Giderler'de 80 satır × 40 bölüm, 2.000 bölüm-kalem, 1.500 revizyon, 1.000 uyarı. Gerekirse `tools/pq_sablon.py` başındaki sabitler büyütülüp şablon yeniden üretilir.
