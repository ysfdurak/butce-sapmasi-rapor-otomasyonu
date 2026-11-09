# Makrosuz yol: Excel Power Query + PivotTable

Şirket bilgisayarında **Python kurulamıyor ve makrolar kapalıysa** bu yolu kullanın. Yalnız Excel'in yerleşik özellikleri kullanılır (Excel 2016 ve sonrası, Microsoft 365). Lisans, eklenti, makro veya kurulum gerekmez.

| | Python / makro sürümü | Power Query yolu |
| --- | --- | --- |
| Gerçekleşen, bütçe, sapma, Lehte/Aleyhte | ✓ | ✓ |
| Müdürlük / bölüm / gider grubu / ay filtreleri, grafik | ✓ (Pano sayfası) | ✓ (dilimleyici + PivotChart) |
| Önemli sapma eşiği, yıl sonu tahmini, hacim/birim maliyet | ✓ | Elle (pivot hesaplanan alan) |
| Dip toplam mutabakatı, revizyon takibi, birim maliyet/endeks | ✓ | — |
| Müdürlük dosyaları, açıklama toplama ve onay izi | ✓ | — |

Bu yol hızlı bir başlangıç ve yönetim panosu içindir; kapanış kontrollerinin tamamı için Python veya makro sürümü gerekir.

## Kurulum (bir kez, yaklaşık 10 dakika)

1. Yeni boş bir Excel dosyası açın ve `Butce_Pano.xlsx` adıyla kaydedin.
2. **Veri › Veri Al › Diğer Kaynaklardan › Boş Sorgu**. Power Query düzenleyicisi açılır.
3. **Giriş › Gelişmiş Düzenleyici**. İçindeki metni silin ve [`powerquery/Butce_Veri.pq`](../powerquery/Butce_Veri.pq) dosyasının tamamını yapıştırın.
4. En üstteki iki satırı düzenleyin:
   - `AyarDosyasi`: Ayarlar.xlsx'in tam yolu. Örnek: `"C:\Butce\ornek\Ayarlar.xlsx"`. Dosya Gezgini'nde dosyaya Shift + sağ tık › **Yol olarak kopyala**.
   - `KapanisDonemi`: boş bırakırsanız klasördeki en son kapanış okunur; belirli bir ay için `"2026-10"`.
5. **Bitti**. Sağdaki **Ad** kutusuna `Butce_Veri` yazın.
6. "Gizlilik düzeyleri" sorulursa iki kaynak için de **Kuruluş** seçin.
7. **Giriş › Kapat ve Yükle**. Veri yeni bir sayfaya tablo olarak gelir.

Maliyet klasörü, hücre konumları, bölüm/müdürlük eşleştirmeleri ve gider grupları Ayarlar.xlsx'ten okunur; sorgunun içini değiştirmeniz gerekmez. Tablo, maliyet sayfasındaki **Bütçe** başlığına göre bulunur.

## Pano (PivotTable + dilimleyici)

1. Tablodaki bir hücreyi seçin › **Ekle › PivotTable › Yeni çalışma sayfası**.
2. Alanlar:
   - **Satırlar:** Müdürlük, Bölüm (veya Gider grubu)
   - **Değerler:** Gerçekleşen, Bütçe, Sapma (gerçekleşen − bütçe). Toplam olarak; sayı biçimi `#.##0`.
3. **PivotTable Çözümle › Dilimleyici Ekle:** Müdürlük, Bölüm, Gider grubu, Ay, Yıl.
   - Bütçe yalnız kapanış yılı için bulunur. Bütçe karşılaştırmasında Yıl dilimleyicisinde kapanış yılını seçin.
   - Önceki yıl karşılaştırması için Yıl'ı **Sütunlar**'a alın.
4. **Sapma oranı** için: **PivotTable Çözümle › Alanlar, Öğeler ve Kümeler › Hesaplanan Alan**.
   - Ad: `Sapma %`, formül: `=(Gerçekleşen-Bütçe)/Bütçe`. Yüzde biçimi verin.
5. **PivotChart** ekleyin. Ay ekseni için Satırlar'a **Ay no** ve **Ay** alanlarını koyun.
6. **Önemli sapmalar:** Satırlara Bölüm ve Gider kalemi'ni alın, Sapma'ya göre büyükten küçüğe sıralayın.
   - **Değer Filtreleri › Büyük veya Eşit** ile TL eşiğinizi uygulayın (örnek 1.000.000).

## Her ay

1. Yeni ayın maliyet dosyalarını aynı klasöre koyun. Dosya adında yıl ve ay yan yana olmalı (`Kasım 2026 - Dosya 1.xlsx`).
2. `Butce_Pano.xlsx`'i açın › **Veri › Tümünü Yenile**. Sorgu en son kapanışı okur; pivotlar ve grafikler güncellenir.

## Sorun giderme

| Mesaj | Neden / çözüm |
| --- | --- |
| `Kaynak okunamadı: <dosya> / <sayfa>: Bütçe başlığı bulunamadı` | Sayfada "Bütçe" yazan başlık yok. Ayarlar'daki `butce_etiketi` ile sayfayı kontrol edin. |
| `... tek 12 aylık bütçe bloğu bulunamadı` | Bütçe başlığının altındaki 12 ay, dosyanın yılıyla eşleşmiyor veya eksik. |
| `... Gider adı bulunamadı: satır N` | `ilk_gider_satiri` / `gider_kalem_sayisi` kaynakla uyuşmuyor; satır eklenmiş olabilir. |
| `DataFormat.Error` / dosya bulunamadı | `AyarDosyasi` yolu veya Ayarlar'daki `maliyet_klasoru` yanlış. |
| Gerçekleşen boş | Kaynak hücre metin içeriyor. Türkçe biçimdeki metin tutarlar (`1.234,56`) çevrilir; diğerleri boş kalır. |

Aynı bölüm iki dosyada bulunursa iki satır da tabloya girer ve toplamlar ikiye katlanır. Kaynak dosyada her bölüm bir kez olmalıdır (Python/makro sürümü bu durumu uyarır).

**Doğrulama sınırı:** Bu sorgu Power Query motorunda çalıştırılarak denenmedi. Yazım ve mantık kontrolü yapıldı; hesap kuralları Python sürümüyle aynıdır. İlk kurulumda örnek verileri yükleyip toplamları karşılaştırın: Ekim 2026, Ocak–Ekim, gerçekleşen **3.449.383.390,31 TL**, bütçe **3.440.938.980,13 TL**. Farklıysa sorun bildirin.
