# Bütçe Sapması Analizi ve Aylık Kapanış Raporlama Otomasyonu

**Excel kaynaklarından bütçe sapması, maliyet ve aylık kapanış analizi**

Uygulama, farklı dosyalardaki bölüm giderlerini tek bir kapanışta toplar. Gerçekleşenleri bütçe ve önceki yılın aynı aylarıyla karşılaştırır, kaynak dip toplamlarını kontrol eder ve geçmiş aylarda sonradan değişen kalemleri gösterir. Direktörlük ve müdürlük raporları üretir.

**Kurulumsuz dene:** [Tarayıcı demosu](https://ysfdurak.github.io/butce-sapmasi-rapor-otomasyonu/demo/) — Python sürümü sentetik veriyle doğrudan tarayıcıda çalışır; yüklediğiniz dosyalar bilgisayarınızdan çıkmaz. [Nasıl çalışıyor?](docs/TARAYICI_DEMOSU.md)

## Çalıştırma seçenekleri

| Sürüm | Çıktı | İlk adım |
| --- | --- | --- |
| Excel VBA | Excel kontrol paneli, direktörlük ve müdürlük raporları | [Windows Excel kurulumu](docs/BASLANGIC.md#excel-vba) |
| Python | Yedi sayfalık yerel dashboard ve Excel raporları | [Python kurulumu](docs/BASLANGIC.md#python) |
| Excel Power Query (makrosuz) | Python ve makro kullanılamayan bilgisayarlarda aynı raporlar ve pano: hazır şablon + tek sorgu, tek tuşla yenileme | [Power Query rehberi](docs/POWER_QUERY.md) |

Her iki sürümde kaynak yolları, bölüm ve gider eşleştirmeleri ve rapor eşikleri Excel'deki ayar dosyasından yönetilir. İki sürüm aynı hesap kurallarıyla aynı rapor sayfalarını üretir ([Veri ve hesap](docs/VERI_VE_HESAP.md)). Python arayüzü tarayıcıda kendi bilgisayarınızda açılır. VBA paketi masaüstü Excel'de rapor üretir.

## Uygulama akışı

1. Maliyet dosyalarının bulunduğu klasörü, ciro/dip toplam kontrol dosyasını ve endeks dosyasını tanımlayın.
2. Kapanış dönemini ve rapor aylarını seçin. Örneğin Ocak, Mart ve Mayıs seçildiğinde, önceki yıl da yalnız aynı üç ayla hesaplanır.
3. Kaynak önizlemesinde bölüm eşleştirmelerini ve dönemleri kontrol edin. Dip toplam farklarını ve önceki kapanıştan gelen revizyonları inceleyin.
4. Raporları seçtiğiniz klasöre üretin. Önceki dönemden gelen açıklamaları topluca inceleyip düzenleyin veya onaylayın.

## Neleri gösterir?

- **Pano** — Müdürlük, bölüm ve gider grubu açılır listeleriyle filtrelenen tablolar ve grafikler; formül tabanlı olduğundan her Excel'de makrosuz çalışır.
- **Yönetici özeti** — Raporun ilk sayfası; toplam sapma ve yönü (Lehte/Aleyhte), yıl sonu tahmini, veri kontrolü ve en büyük önemli sapmalar. Tek sayfaya yazdırılır.
- **Bütçe sapması ve gider yapısı** — Gerçekleşen, bütçe, bütçe eksi gerçekleşen, yüzde sapma ve ana/alt gider kırılımı.
- **Önemli sapmalar** — TL ve yüzde eşiğini birlikte aşan kalemler; müdürlük dosyalarında açıklanması gereken hücreler işaretlenir ve açıklama durumu izlenir.
- **Yıl sonu tahmini ve hacim/birim maliyet etkisi** — Gerçekleşen + kalan bütçe ve aylık ortalama gidiş tahmini; bütçe farkının işlem adedi (hacim) ve birim maliyet etkisine ayrıştırılması.
- **Operasyon göstergeleri** — Kaynaktaki başlıklarıyla personel ve işlem adetleri. Bu adet satırları gider toplamının dışında tutulur.
- **Birim maliyet ve OPEX/ciro** — Endekslenmiş önceki yıl birim maliyeti, gider ağırlıklı reel iyileşme ve şirket cirosu üzerinden OPEX oranı.
- **Kapanış kontrolleri** — Bölüm dip toplamı mutabakatı, geçmiş ay gider/bütçe değişiklikleri ve eksik eşleştirme uyarıları.
- **Açıklama takibi** — Önceki dönem metinlerini inceleme kuyruğuna alma, toplu onay, onay izi (kim, ne zaman) ve müdürlük metinlerini bölüm adlarıyla genel raporda birleştirme.

Python dashboard sayfaları Genel Durum, Personel, İşlem Adedi, Detay Durum, Bölüm Analizi, Birim Maliyet ve Sapma Öncelikleri. Altı ortak seçim müdürlük, müdürlük altı, lokasyon, bölüm, yıl ve ay kapsamını sayfalar arasında taşır.

## Sentetik demo

Örnek set tek direktörlük, dört müdürlük ve 12 bölüm içerir. Eylül ve Ekim 2026 kapanışlarının her birinde beş maliyet dosyası bulunur. Ekim kapanışı ve Ocak–Ekim seçimiyle beklenen gerçekleşen **3.449,38 milyon TL**, bütçe **3.440,94 milyon TL**, fark **−8,44 milyon TL**. Rapor akışı bir direktörlük geneli ve dört müdürlük dosyası üretir. [Başlangıç ve beklenen sonuçlar](docs/BASLANGIC.md)

<details>
<summary>Excel kontrol paneli önizlemesi</summary>

![Excel kontrol paneli tasarım önizlemesi](makro/Kontrol_Paneli_Onizleme.png)

Görsel, kontrol paneli tasarım önizlemesidir.

</details>

## Doğrulama durumu

- Python'un **94 otomatik testi** ve VBA kaynak denetimi (`tools/vba_araclari.py`) GitHub Actions'ta her gönderimde çalışır.
- Aynı sentetik beklenen değerler Python testlerinde ve makrodaki **Sentetik hesap testi**nde kontrol edilir.
- V4'te kaydedilen **72 JavaScript/API kontrolü**, arayüzün çizim ve olay mantığının benzetimini ve canlı yerel API'yi kapsar.
- Windows Excel'de VBA derlemesi, gerçek tarayıcı yerleşimi ve Windows dosya pencereleri [canlı kabul listesiyle](docs/CANLI_KONTROL.md) doğrulanır.

[Hesap ve veri akışı](docs/VERI_VE_HESAP.md) · [Açıklamalar](docs/YORUM_KULLANIMI.md) · [Doğrulama](docs/DOGRULAMA.md) · [Canlı kontrol](docs/CANLI_KONTROL.md) · [Sürüm notları](docs/SURUM_NOTLARI.md) · [Ortak kararlar](docs/ORTAK_KARARLAR.md)

Depodaki çalışma kitapları sentetik örnekler ve şablonlardır. Gerçek kaynakları, ayar kopyalarını, raporları, günlükleri ve yerel kayıtları kendi bilgisayarınızda veya kurumunuzun onaylı depolama alanında tutun.
