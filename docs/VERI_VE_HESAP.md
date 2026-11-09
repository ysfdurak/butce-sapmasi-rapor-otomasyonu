# Veri ve hesap akışı

| Kaynak | İşlev | Örnekteki konum |
| --- | --- | --- |
| Maliyet Excel'leri | Kapanışa ait bölüm sayfaları, gerçekleşen ve bütçe | Gerçekleşen tarihleri BF59. Bütçe başlangıç ipucu EZ58. İlk dört personel satırı 60–63, işlem adedi 64, parasal giderler 65 ve sonrası |
| Kontrol ve ciro Excel'i | Bölüm dip toplamı mutabakatı ve tek şirket cirosu | Gerçekleşen tarihler D18, ciro D19, bütçe başlangıç ipucu CW18 |
| Endeks Excel'i | Bölüm maliyet endeksinin tarih serisi | İkinci tablo C17, tarihler D17'den ilk boşluğa kadar |
| Ayarlar.xlsx | Bölüm/müdürlük/lokasyon, gider grupları, gider birleştirme, atlanacak sayfalar ve kaynak konumları | Her sürümün `ornek/` klasörü |

Birden fazla bölüm aynı kontrol grubuna bağlanabilir. Örnekte Giriş ve Tartım, Ürün Kabul Toplamı'na bağlanır. Endeks bölüm adları ayrı eşleştirilir. Kaynakları yenilemek yeni eşleştirmeleri sisteme alır. Son kapanış hesap kaynağıdır, önceki kapanışlar revizyon karşılaştırmasına girer. Gerçekleşen ve bütçedeki aynı tarih, iki ayrı kaynak bloğundan okunur.

## Göstergeler

- **Bütçe sapması** = bütçe − gerçekleşen. Sapma oranı, farkın bütçeye bölünmesiyle bulunur. Raporun özellikle *aşım oranı* olarak adlandırdığı alan ters yöndedir.
- **OPEX/ciro** = kapsamdaki gider / tek şirket cirosu. Karşılaştırma gerektiren bölüm tablolarında önceki yıl verisi bulunan bölümler kullanılır.
- **Birim maliyet** = seçili ayların gider toplamı / ilgili işlem adedi toplamı.
- **Dönem endeksi** = seçili aylarda (bu yılın ay endeksi / önceki yılın aynı ay endeksi) oranlarının aritmetik ortalaması.
- **Reel iyileşme** = (endekslenmiş önceki yıl birim maliyeti − güncel birim maliyet) / endekslenmiş önceki yıl birim maliyeti. Grup iyileşmesi, karşılaştırılabilir bölümlerin güncel gider paylarıyla ağırlıklandırılır.

## Uygulama yapısı

Python'da `app.py` yerel arayüz ve API'yi sunar. `butce.py`, `model.py` ve `birim.py` kaynak okuma, hesaplama ve raporlama akışını oluşturur. `web/` arayüzü, `tests/` otomatik kontrolleri, `wheels/` çevrimdışı bağımlılıkları içerir.

Excel VBA'da `vba/` okunabilir UTF-8 kaynakları, `vba_elle/` Excel'in Türkçe Windows içe aktarma kopyalarını, `Kontrol_Sablonu.xlsx` kontrol kitabının başlangıç sayfalarını içerir. Kurulum 10 VBA modülünü çalışma kitabına ekler.
