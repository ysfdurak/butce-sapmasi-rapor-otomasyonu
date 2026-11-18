# Veri ve hesap akışı

Bu sayfa iki sürümün (Python ve Excel VBA) ortak hesap tanımıdır. Bir kural değişirse iki sürüm de birlikte değişir; Python `tests/test_analiz.py` ve makro `MakroTest` aynı sentetik beklenen değerleri kontrol eder.

| Kaynak | İşlev | Örnekteki konum |
| --- | --- | --- |
| Maliyet Excel'leri | Kapanışa ait bölüm sayfaları, gerçekleşen ve bütçe | Gerçekleşen tarihleri BF59. Bütçe başlangıç ipucu EZ58. İlk dört personel satırı 60–63, işlem adedi 64, parasal giderler 65 ve sonrası |
| Kontrol ve ciro Excel'i | Bölüm dip toplamı mutabakatı ve tek şirket cirosu | Gerçekleşen tarihler D18, ciro D19, bütçe başlangıç ipucu CW18 |
| Endeks Excel'i | Bölüm maliyet endeksinin tarih serisi | İkinci tablo C17, tarihler D17'den ilk boşluğa kadar |
| Ayarlar.xlsx | Bölüm/müdürlük/lokasyon, gider grupları, gider birleştirme, atlanacak sayfalar, kaynak konumları ve rapor eşikleri | Her sürümün `ornek/` klasörü |

Birden fazla bölüm aynı kontrol grubuna bağlanabilir. Örnekte Giriş ve Tartım, Ürün Kabul Toplamı'na bağlanır. Endeks bölüm adları ayrı eşleştirilir. Kaynakları yenilemek yeni eşleştirmeleri sisteme alır. Son kapanış hesap kaynağıdır, önceki kapanışlar revizyon karşılaştırmasına girer. Gerçekleşen ve bütçedeki aynı tarih, iki ayrı kaynak bloğundan okunur. Sayfa ve müdürlük adları büyük/küçük harf ve Türkçe harf farkı gözetilmeden eşleştirilir.

## İşaret kuralı

- **Sapma = gerçekleşen − bütçe.** Pozitif sapma bütçe aşımıdır ve **Aleyhte** yazılır; negatif sapma bütçe altıdır ve **Lehte** yazılır.
- Genel Özet ve Giderler sayfalarındaki eski **Bütçe − gerçekleşen** sütunu ters işaretlidir (pozitif = bütçe altı, yeşil). **Bütçe aşım oranı** = (gerçekleşen − bütçe) / bütçe; pozitif = aşım, kırmızı. Bu iki sütun önceki raporlarla karşılaştırılabilirlik için korunur.

## Göstergeler

- **OPEX/ciro** = kapsamdaki gider / tek şirket cirosu. Karşılaştırma gerektiren bölüm tablolarında önceki yıl verisi bulunan bölümler kullanılır.
- **Birim maliyet** = seçili ayların gider toplamı / ilgili işlem adedi toplamı.
- **Dönem endeksi** = seçili aylarda (bu yılın ay endeksi / önceki yılın aynı ay endeksi) oranlarının aritmetik ortalaması.
- **Reel iyileşme** = (endekslenmiş önceki yıl birim maliyeti − güncel birim maliyet) / endekslenmiş önceki yıl birim maliyeti. Grup iyileşmesi, karşılaştırılabilir bölümlerin güncel gider paylarıyla ağırlıklandırılır.

## V5 analizleri

| Sayfa | Kural |
| --- | --- |
| **Yönetici Özeti** (ilk sayfa, tek sayfaya yazdırılır) | Kapsamın gerçekleşen, bütçe, sapma, sapma oranı ve yönü; yıl sonu tahmini; veri kontrolü (doğrulanan / toplam); uyarı sayısı; önemli sapma sayısı; en büyük *N* aleyhte ve *N* lehte önemli sapma (`ozet_kalem_sayisi`, varsayılan 5). Kapanış kaynağı eksik bölüm varsa toplamlar boş kalır. |
| **Pano** (ikinci sayfa) | Müdürlük, bölüm ve gider grubu açılır listeleri (B4:B6, varsayılan Tümü); toplam kartları, aylık tablo, gider grubu ve bölüm tabloları, aylık gerçekleşen/bütçe ve gider grubu sapma grafikleri. Formüller gizli `_Pano Veri` sayfasındaki bölüm × gider kalemi × seçili ay satırlarını SUMIFS ile toplar; pivot yenilemesi veya makro gerekmez. Eksik aylık tutar 0 sayılır (kesin toplamlar Genel Özet'tedir). |
| **Önemli Sapmalar** | Bölüm × alt gider × seçili ay toplamında \|sapma\| ≥ `onemlilik_tutari` **ve** \|sapma / bütçe\| ≥ `onemlilik_orani`. Bütçe sıfırsa oran tanımsızdır; yalnız tutar eşiği uygulanır. Sıralama \|sapma\| büyükten küçüğe. Müdürlük dosyasında açıklama hücresi turuncu olur ve **Açıklama durumu** formülü `Açıklandı` / `Açıklama gerekli` gösterir. |
| **Yıl Sonu Tahmini** | Kapanış yılı için. Yöntem 1: Ocak–kapanış ayı gerçekleşen + kalan ayların bütçesi. Yöntem 2 (aylık ortalama gidiş): Ocak–kapanış gerçekleşen / kapanış ayı × 12. Her ikisi yıllık bütçeyle karşılaştırılır. Seçili ay listesinden bağımsızdır. |
| **Hacim ve Birim Maliyet** | Seçili aylarda: **hacim etkisi** = (gerçekleşen adet − bütçe adet) × bütçe birim maliyeti; **birim maliyet etkisi** = gerçekleşen gider − bütçe birim maliyeti × gerçekleşen adet. İki etkinin toplamı gerçekleşen − bütçe farkıdır. Sıfır/negatif adetli ay bulunan bölümde etkiler boş kalır. |

Seçili yıl bütçe yılı değilse önemlilik uygulanmaz; tahmin ve hacim sayfaları oluşturulmaz.

**Kalem ayrıntısı (V5.6, Python):** hacim ve birim maliyet etkisi tek gider satırına da ayrılır. Kalem bütçe birim maliyeti = kalem bütçesi / bölümün bütçe işlem adedi; formüller yukarıdakiyle aynıdır. Bir bölümün bütün kalemlerinin etkileri toplamı bölüm satırına eşittir. [Kalem ayrıntısı](KALEM_AYRINTISI.md)

## V5.4 bütçe sürümleri ve trend

Python (`surum.py`), makro (`MakroSurum`) ve Power Query şablonu aynı kuralları uygular.

| Konu | Kural |
| --- | --- |
| Güncel bütçe | Kapanış ayı dosyasındaki bütçe. Revize bütçe ayrı satırda gelmez; aynı hücrelerdeki değerin değişmesinden anlaşılır. |
| Önceki ay bütçesi | Aynı yılın bir önceki kapanış dosyası. Ocak kapanışında karşılaştırma yapılmaz. |
| İlk bütçe | Bölümün kapanış yılında okunan ilk kapanış dosyası (Ocak yoksa ilk bulunan ay; Girdi Kaydı'nda görünür). |
| Karşılaştırma bütçesi | `karsilastirma_butcesi = İlk bütçe` ise tüm sapma sayfaları ilk bütçeyi kullanır; kalem bütçesi ilk dosyadan, toplam yalnız kapanış dosyasındaki kalemlerden (eksik varsa boş). Bütçe işlem adedi ve ciro güncel kalır; revizyonlar kaynak dosyalardan hesaplanır. |
| Bütçe - Bölüm / Müdürlük / Direktörlük | Güncel bütçe ay ay; yıl toplamı yalnız 12 ay doluysa; ilk bütçe yıl toplamı ve fark. |
| Bütçe Değişimi | Güncel − önceki ay bütçesi; yalnız \|fark\| > `tolerans` olan satır ve kalemler. Kalem, iki dosyada da olmalı. |
| Trend | Aylık gerçekleşen − karşılaştırma bütçesi. Kalem, kapanış ayında biten seride üst üste `trend_en_az_ay` ay boyunca fark ≥ `trend_aylik_esik` ise listelenir; sıra: seri uzunluğu, sonra seri toplamı (büyükten küçüğe). |
| Yönetici Özeti | İki satır: bütçesi değişen kalem sayısı ve yıllık etkisi; üst üste aleyhte giden kalem sayısı. |

## Konumların otomatik bulunması (V5.5)

Python (`konum.py`), makro (`MakroKonum`) ve Power Query (`Butce_Veri.pq` › Otomatik konum) aynı kuralları uygular. Ayarlar'da değer varsa o kullanılır; yoksa dosyadan bulunur ve Uyarılar'a "Konum otomatik bulundu" yazılır (Power Query yazmaz). Bulunamazsa hata konum satırının girilmesini ister.

| Dosya | Bulunan ayarlar | Kural |
| --- | --- | --- |
| Maliyet (en yeni dosyanın ilk bölüm sayfası) | `butce_etiketi`, `tarih_baslangici`, `gider_sutunu`, `islem_adedi_satiri`, `ilk_gider_satiri`, `gider_kalem_sayisi` | "Bütçe" yazan ve altı tarih/ay olan hücre bütçe etiketidir (birden fazla satırda varsa belirsiz sayılır). Altındaki satır başlıktır; başlıkta etiketin solundaki ilk tarih gerçekleşen başlangıcıdır. Bunun solunda, altındaki 59 satırda en az üç yazı olan en yakın sütun gider adı sütunudur. Bu sütunda "İşlem Adedi" geçen satır işlem adedi satırı, bir altı ilk gider satırıdır; giderler ilk boş ada kadar sayılır. Personel satırları işlem adedinin üstündeki dört satırdır. |
| Kontrol | `ciro_baslangici`, `butce_ciro_baslangici`, `etiket_sutunu` | En az 12 tarihi olan ve ilk tarihin altında sayı bulunan en üst satır ciro başlığıdır. Tarihler kesildikten sonraki ilk tarih bütçe cirosu başlangıcıdır. "Opex Toplamı" yazan tek sütun etiket sütunudur. |
| Endeks | `endeks_araligi` | "Bölüm Gider Endeksi" veya "Bölüm Adı" yazan ve sağında tarih olan tek hücre tablonun sol üstüdür. |

Ayardaki konum bir maliyet sayfasında tutmazsa (başlık veya gider adı bulunamazsa) o sayfa için konum yeniden bulunur ve uyarılır; diğer sayfalar ayardaki konumla okunur. Eski satır/sütun düzeni ayarlarından (`yil_satiri`, `ciro_satiri`, `ilk_veri_satiri` vb.) biri girilmişse eski düzen kullanılır ve konum aranmaz. Sayfa adları (`sayfa`) ve `tolerans` ayar olarak kalır.

## Rapor ayarları

Ayarlar.xlsx › **Dosya ve Konum Ayarları** sayfasına `Rapor` alanıyla eklenir. Satır yoksa varsayılan kullanılır.

| Ayar kodu | Varsayılan | Anlamı |
| --- | ---: | --- |
| `onemlilik_tutari` | 1.000.000 | TL eşiği |
| `onemlilik_orani` | 0,10 | Oran eşiği (0,10 = %10; 0 ile 1 arası) |
| `ozet_kalem_sayisi` | 5 | Yönetici özetindeki aleyhte/lehte kalem sayısı |
| `arsiv_saklama_adedi` | 20 | Girdi arşivinde tutulan son çalışma sayısı (0 = sınırsız) |
| `karsilastirma_butcesi` | Güncel bütçe | `İlk bütçe` yazılırsa raporlar yıl başı bütçesiyle karşılaştırır |
| `trend_aylik_esik` | önemlilik tutarı / 12 | Trend: bir ayın aleyhte sayılması için gereken TL |
| `trend_en_az_ay` | 3 | Trend: listeye girmek için üst üste aleyhte ay sayısı (1–12) |

Konumlu düzen ayarları (`tarih_baslangici`, `ciro_baslangici`, `baslik_satiri`/`endeks_araligi`) tanımlıysa eski satır/sütun düzeninin ayarları (`yil_satiri`, `ciro_satiri`, `ilk_veri_satiri` vb.) zorunlu değildir; yalnız yedek olarak okunur.

## Kaynak sağlamlığı

- **Metin tutarlar:** `1.234,56`, `1234,56`, `(1.234)`, `1234.56`, `7 500 TL` sayıya çevrilir; nokta + üçlü basamak grubu Türkçe binlik ayırıcı sayılır (`1.234` = 1234). Çevrilen ve çevrilemeyen hücreler hücre adresiyle uyarılır; çevrilemeyen hücre boş sayılır ve ilgili toplam boş kalır.
- **Formül sonuçları:** Kaynak dosya Excel dışında kaydedildiyse formül sonucu dosyada bulunmayabilir. Bu hücreler adresleriyle uyarılır; dosyayı Excel'de açıp kaydedin.
- **Düzen kayması:** Tarih başlığı veya Bütçe etiketi ayardaki hücrede bulunamazsa yakın satır/sütunlarda aranır ve yeni hücre önerilir (ör. `BF59 yerine BF61`). Ayar kendiliğinden değiştirilmez; Tablo düzeni ekranından onaylanır.
- **Yıl dönümü revizyonu:** Ocak kapanışında önceki yılın Aralık dosyası okunur ve geçen yılın 12 ayının gerçekleşenlerinde sonradan yapılan değişiklikler revizyon tablosuna girer. Diğer aylarda ardışık iki kapanış karşılaştırılır (gerçekleşen: önceki aylar; bütçe: 12 ay).

## Uygulama yapısı

Python'da `app.py` yerel arayüz ve API'yi sunar. `butce.py` kaynak okuma ve rapor motoru, `model.py` dashboard hesapları, `analiz.py` V5 analizleri, `birim.py` bütçe işlem adetleri, `notes.py` ve `note_review.py` açıklama kaydı ve onay izidir. `web/` arayüzü, `tests/` otomatik kontrolleri, `wheels/` çevrimdışı bağımlılıkları içerir.

Excel VBA'da `vba/` okunabilir UTF-8 kaynakları, `vba_elle/` Excel'in Türkçe Windows içe aktarma kopyalarını (Windows-1254, CRLF), `Kontrol_Sablonu.xlsx` kontrol kitabının başlangıç sayfalarını içerir. Kurulum 11 VBA modülünü çalışma kitabına ekler. `vba_elle/` elle düzenlenmez: `python tools/vba_araclari.py uret` ile `vba/` klasöründen üretilir; `python tools/vba_araclari.py kontrol` kopyaların güncel olduğunu ve VBA kaynaklarında ayrılmış sözcük, kodlama ve modül listesi hatası olmadığını denetler.
