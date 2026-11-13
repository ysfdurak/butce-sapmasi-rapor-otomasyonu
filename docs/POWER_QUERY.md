# Makrosuz yol: Power Query + rapor şablonu

Şirket bilgisayarında **Python kurulamıyor ve makrolar kapalıysa** bu yolu kullanın. Yalnız Excel'in yerleşik özellikleri kullanılır: Power Query ve formüller (Excel 2016, 2019, 2021, Microsoft 365). Lisans, eklenti, makro veya kurulum gerekmez. Formüllerde yeni sürümlere özgü işlevler (FILTER, XLOOKUP, LET…) kullanılmadı; şablon tüm bu sürümlerde aynı çalışır.

İki dosya:

| Dosya | Görevi |
| --- | --- |
| [`powerquery/Butce_Rapor_Sablonu.xlsx`](../powerquery/Butce_Rapor_Sablonu.xlsx) | Raporlar ve panolar (formüllerle). Kontrol Paneli'ndeki seçimler değişince her şey anında güncellenir. |
| [`powerquery/Butce_Veri.pq`](../powerquery/Butce_Veri.pq) | Tek Power Query sorgusu. Ayarlar, maliyet, kontrol ve endeks dosyalarını okur; şablonun **Veri** sayfasına tek tablo yükler. |

## Python / makro sürümüyle karşılaştırma

| | Python / makro | Power Query şablonu |
| --- | --- | --- |
| Yönetici Özeti, Pano (grafikler), Genel Özet, Giderler | ✓ | ✓ |
| Birim maliyet ve endeks, Hacim ve birim maliyet | ✓ | ✓ |
| Önemli sapmalar, Yıl sonu tahmini | ✓ | ✓ |
| Kontrol dosyası mutabakatı (bölüm / birleşik / Opex / ciro) | ✓ | ✓ (Veri Kontrolü) |
| Revizyonlar (ardışık kapanış dosyaları) ve uyarılar | ✓ | ✓ |
| Kapsam: direktörlük veya tek müdürlük, ay seçimi, önceki yıl | ✓ | ✓ (Kontrol Paneli) |
| Açıklama girişi | Müdürlük dosyalarından toplanır | Açıklamalar sayfasına yazılır (dosyalardan toplama: sonraki aşama) |
| Müdürlük dosyalarını ayrı ayrı üretme, onay izi, girdi arşivi | ✓ | — (Kapsam seçip **Farklı Kaydet**) |

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
| Açıklama eklemek için | **Açıklamalar** sayfası: bölüm + gider kodu + açıklama. Giderler ve Önemli Sapmalar'da görünür. |

Kontrol Paneli seçimleri:

- **Rapor yılı:** boş = kapanış yılı. Önceki bir yıl seçilirse bütçe karşılaştırması yapılmaz.
- **Kapsam:** `Tümü` = direktörlük raporu; bir müdürlük seçilirse yalnız o müdürlük.
- **Ay seçimi:** Yıl başından kapanışa · Yalnız kapanış ayı · Elle seçim (Elle seçimde ayların altına `1`).
- **Önemlilik tutarı / oranı, Özet kalem sayısı:** boş bırakılırsa Ayarlar.xlsx'teki değerler.

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

1. **Gerçek M motoru:** Sorgu, Excel'deki Power Query ile aynı motorla (Microsoft Power Query SDK, PQTest) örnek verilerle çalıştırılır; çıkan tablo Python benzetimiyle satır satır karşılaştırılır (en son kapanış ve 2026-09).
2. **Şablon formülleri:** Şablon bu veriyle LibreOffice'te hesaplanır; Genel Özet, Giderler, Birim Maliyet, Önemli Sapmalar, Yıl Sonu Tahmini, Hacim ve Uyarılar Python motorunun raporlarıyla karşılaştırılır (direktörlük yıl başından, yalnız kapanış ayı, Depolama · Ocak/Mart/Mayıs).
3. **Birim testleri:** blok sayıları, tekil anahtarlar, boş ay kuralı.

Elle kontrol: örnek verilerle Ekim 2026, Ocak–Ekim, Tümü: gerçekleşen **3.449.383.390,31 TL**, bütçe **3.440.938.980,13 TL**.

## Sorun giderme

| Belirti | Neden / çözüm |
| --- | --- |
| Kontrol Paneli: "Veri yok" | Sorgu `Veri!$A$1`'e yüklenmedi. **Veri › Sorgular ve Bağlantılar** › `Butce_Veri` sağ tık › **Yükle…** › Var olan sayfa `=Veri!$A$1`. |
| `DataFormat.Error` / dosya bulunamadı | `AyarDosyasi` yolu veya Ayarlar'daki `maliyet_klasoru`, `kontrol_dosyasi`, `endeks_dosyasi` yanlış. |
| `Formula.Firewall` hatası | **Veri › Veri Al › Sorgu Seçenekleri › Geçerli Çalışma Kitabı › Gizlilik › Gizlilik düzeylerini yoksay**. |
| Uyarılar: `Bütçe başlığı bulunamadı` | Sayfada "Bütçe" başlığı yok. Ayarlar'daki `butce_etiketi` ile sayfayı karşılaştırın. |
| Uyarılar: `tek 12 aylık bütçe bloğu bulunamadı` | Bütçe başlığının altındaki 12 ay dosyanın yılıyla eşleşmiyor veya eksik. |
| Uyarılar: `Gider adı bulunamadı` | `ilk_gider_satiri` / `gider_kalem_sayisi` kaynakla uyuşmuyor; satır eklenmiş olabilir. |
| Uyarılar: `Aynı bölüm ve kapanış için birden fazla kaynak` | Bölüm iki dosyada var; ikisi de alınmaz. Fazlasını kaldırın. |
| Uyarılar: `Kontrol/Endeks dosyası okunamadı` | Dosya adı veya sayfa adı Ayarlar'la uyuşmuyor; diğer raporlar yine üretilir. |
| Yenileme çok uzun sürüyor | Maliyet klasöründe yalnız kapanış dosyaları olsun; eski yılların dosyalarını alt klasöre taşıyın (alt klasörler okunmaz). |

Şablon sınırları: 60 bölüm, Genel Özet'te 120 satır, Giderler'de 80 satır × 40 bölüm, 2.000 bölüm-kalem, 1.500 revizyon, 1.000 uyarı. Gerekirse `tools/pq_sablon.py` başındaki sabitler büyütülüp şablon yeniden üretilir.
