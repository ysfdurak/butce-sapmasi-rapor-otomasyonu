# Bütçe Sapması — Python V5

> **V5:** Raporlara Yönetici Özeti, Önemli Sapmalar, Yıl Sonu Tahmini ve Hacim ve Birim Maliyet sayfaları eklendi; rapor üretimi yaklaşık 12 kat hızlandı. Hesap kuralları: [Veri ve hesap](../docs/VERI_VE_HESAP.md). Açıklama akışı: [Açıklamalar ve onay izi](../docs/YORUM_KULLANIMI.md). Önceki sürüm notları [docs/arsiv](../docs/arsiv/) klasöründedir.

Bu sürüm, raporlama ve yedi sayfalık dashboard'u **tek yerel Python uygulamasında** birleştirir. Önceki Excel/makro sürümleri ayrıca korunmuştur. Bu paket Excel COM, VBA, PowerShell, pip veya internetten grafik kütüphanesi indirme kullanmaz. Dashboard için masaüstü Excel'in açık olması gerekmez.

## Açılış

1. ZIP'i bilgisayarında yeni bir klasöre tamamen çıkar.
2. **Baslat.bat** dosyasına çift tıkla. İstersen Python dosya ilişkisi kurulmuşsa **Baslat.pyw** ile de açabilirsin.
3. Python yerel işlemi başlatır, varsayılan tarayıcıda dashboard açılır. Tarayıcı adresi `127.0.0.1` olur. Bu ekran bir internet sitesi değildir; hesaplar bilgisayarındaki Python işleminde yapılır.
4. İlk açılış sentetik örnekle gelir. 12 bölüm, dört müdürlük, tek direktörlük vardır; her kapanış için beş maliyet Excel'i kullanılır.
5. İşin bitince sağ üstteki **⏻** düğmesiyle uygulamayı durdur; sonra sekmeyi kapat. Yalnız sekmeyi kapatmak Python işlemini durdurmaz. Açılışta görünen komut penceresi de açık kalmalıdır.

Python **3.9 ve üzeri** gerekir. Başlatıcı önce 3.13, sonra 3.14, sonra bilgisayardaki diğer Python 3 sürümünü arar. Gerekli openpyxl 3.1.5 ve et_xmlfile dosyaları `wheels` klasöründedir; doğrudan buradan yüklenir. Kurulum ve pip komutu çalışmaz; SSL üzerinden paket indirme yapılmaz. `wheels` klasörünü silme.

Bu paket hazır bir Windows `.exe` değildir. Daha önce kurduğun Python kullanılacaktır. Windows'ta dosya seçme pencereleri için Python'un standart Tk bileşeni gerekir; standart Windows kurulumunda bulunur. Dosya penceresi açılmazsa kaynak/çıktı yolunu ekrana elle yazabilirsin.

## İlk kontrol

Başlangıç: Ekim 2026 kapanışı, yıl 2026 ve Ocak–Ekim seçili; bütün bölümler açık.

| Gösterge | Beklenen |
|---|---:|
| Gerçekleşen gider | 3.449,38 milyon TL |
| Bütçe gideri | 3.440,94 milyon TL |
| Bütçe − gerçekleşen | −8,44 milyon TL |
| İşlem adedi gerçekleşen | 69.216.031 |
| İşlem adedi bütçe | 67.920.000 |
| Genel reel BM iyileşmesi | yaklaşık −%1,52 |
| Brüt gider aşımı | 230,82 milyon TL |
| Net bütçeyi aşan bölüm | 9 |
| Aşım kaydı | 234 |
| İlk beş aşımın brüt aşım içindeki payı | %23,56 |
| Geriye dönük gider/bütçe değişikliği | 26 kayıt |

Ay seçiminde yalnız Ocak, Mart ve Mayıs açık olursa gerçekleşen **913,19 milyon TL**, bütçe **1.094,64 milyon TL** olur. Aradaki Şubat ve Nisan toplanmaz. Önceki yıl karşılaştırması da aynı üç ayı kullanır.

## Ortak seçimler

Üstteki altı listbox: **Müdürlük, Müdürlük altı, Lokasyon, Bölüm, Yıl ve Ay**. Kutulara tıklayarak çoklu seçim yapılır; Ctrl gerekmez. Sayfa değiştirdiğinde bütün ortak seçimler korunur. Liste başındaki Tümü/Temizle, ilgili listeyi değiştirir. **Sıfırla**, kaynak bütçe yılı ve kapanışa kadar aylarla tüm bölümleri getirir.

**Seçimi kaydet**, ortak seçimleri ve BM/OPEX yerel kapsamlarını sonraki açılışa saklar. Kayıtlar yalnız `kullanici/kayitlar.sqlite3` dosyasındadır. Yeni sürüme geçerken açıklamaları ve seçimleri korumak için `kullanici` klasörünü yedekle. Bu dosya JSON ayar dosyası değildir; senin düzenlemen gerekmez. İş kuralları ve eşleştirmeler yine Excel'dedir.

Genel kartlarda bütçe ancak **bütçe yılı tek başına** seçiliyken görünür. Birden fazla yıl seçersen gerçekleşenler toplanır; genel bütçe/sapma kartları ve Detay Durum bütçe alanları boş kalır. İşlem ve Bölüm Analizi gibi yıl bazlı tablolarda, seçili bütçe yılının kendi bütçe sütunları korunur.

## Yedi sayfa

1. **Genel durum:** Dört kart; müdürlük gerçekleşen ve bütçe sapma yatay grafikleri; ayların ayrı ayrı bütçe/gerçekleşen çizgileri. Çizgi grafiği birikimli değildir. Alt/üst değer girerek veya +/− ve ↑/↓ ile dikey aralığı değiştirirsin. Uzun grafikte yatay kaydırma vardır.
2. **Personel:** İlk dört adet satırı kullanılır; grafik adları kaynak Excel'deki gerçek başlıklardan okunur. Örneğin Kadrolu, Taşeron, Fiili ve Planlı çalışan adları kaynakta neyse aynen gelir. Gerçekleşenler yıl renkleriyle sütun, bütçe ayrı renkte çizgi. İki görünüm: yıl/ay detayı veya tam seçili ayların toplamı. Toplam görünümü aylık personel sayılarının toplamıdır; ortalama kişi sayısı değildir. Yeni yıllar veri kaynağından eklenir, grafik yatay genişler. Son iki bütçe hücresi boşsa, uyumlu ilk iki göstergenin bütçesiyle tamamlanır; ilk iki değer çelişirse boş alan tamamlanmaz. Seçili bölümlerde aynı satırın adı farklıysa adlar birlikte gösterilir ve hangi bölüme ait oldukları belirtilir; tek bölüm seçilince o bölümün başlığı gelir. Kaynak başlığı boşsa bu durum açıkça uyarılır.
3. **İşlem adedi:** Beşinci adet satırı; bu göstergenin adı da kaynak başlık hücresinden gelir. Dört kart; yıl başlıklarının altında açılan gerçekleşen, bütçe, yıllık değişim ve bütçe sapması. Yıllık değişim adet ve yüzdeyi birlikte gösterir. Sapma = bütçe − gerçekleşen; işlem adedinde hedefi aşan gerçekleşen yeşil, hedef altında kalan kırmızıdır. Bu, gider ekranlarının renk yönünden farklıdır.
4. **Detay durum:** Ana gider +/− düğmesinden alt giderler açılır. Gerçekleşen, bütçe, fark, oran ve en altta toplam. Sağdaki halka grafiği gerçekleşen giderlerden hesaplanır; etiketleri yalnız yüzdedir. Gerçekleşen hücresinin üzerine geldiğinde veya klavyeyle odaklandığında, aynı kalemin seçili yıllardaki değerini aynı seçili aylar için çizgiyle görürsün.
5. **Bölüm analizi:** Bölüm giderleri / OPEX-ciro görünümleri. Bölümler, hemen altında ait oldukları müdürlük toplamı ve en altta genel toplam. Gider tutarları açıkça milyon TL gösterilir. OPEX oranlarında önceki yıl karşılaştırması yapılamayan bölümler çıkarılır. Payda şirketin tek ciro toplamıdır.
6. **Birim maliyet:** Endeksle eşleştirilen ve kıyaslanabilir bölümler. Önceki yıl, bütçe yılı gerçekleşen, bütçe, enflasyon, endekslenmiş önceki BM, reel iyileşme ve iki gider ağırlığı. Müdürlük/direktörlük iyileşmeleri güncel giderlerle ağırlıklandırılır. Alt geniş kart seçili karşılaştırılabilir bölümlerin genel iyileşmesidir. Eski yıl seçiminde yalnız gerçekleşen BM gösterilir.
7. **Sapma öncelikleri:** Brüt aşım, net aşan bölüm, ilk beş payı, aşım sayısı; en büyük on gider aşımı ve tüm sıralama. Bölüm/gider araması ve açıklama girişleri vardır. Bir gider altındaki olumlu fark başka giderdeki aşımı gizlemesin diye brüt aşım ayrı hesaplanır. Bütçe yılını tek başına seçmek gerekir.

**BM ve OPEX tablolarında kendi “Tablo bölümleri” listesi bulunur.** Bu seçim üstteki ortak bölüm listesiyle kesişir; diğer sayfaları veya birbirini değiştirmez. Hariç kalan bölümlerden sonra ağırlıklar yeniden normalize edilir. Örneğin BM yerel listesinde yalnız Giriş seçersen ağırlığı %100, genel BM iyileşmesi yaklaşık −%0,93 olur; Genel Durum kartı değişmez.

Tutarlar iki ondalık ve milyon TL; BM TL/işlem, oranlar iki ondalık yüzdedir. Personel ve işlem adetleri ondalıksız, binlik ayıraçlıdır. Negatif tutarlar eksi işaretiyle yazılır.

## Kendi kaynaklarını bağlamak / ayı yenilemek

1. **Kaynaklar** düğmesini aç. Ayar Excel'ini Gözat ile seç veya tam yolunu yaz. Gözat ile seçimde ayarlar otomatik okunur; elle yol yazdıysan **Seçilen ayarları oku** düğmesine bas.
2. **Maliyet dosyaları klasörü**, **ciro ve tutar kontrol dosyası** ve **endeks dosyası** için ayrı Gözat düğmeleri vardır. Yolları ilk kullanımda bir kez seç; ihtiyaç halinde daha sonra değiştir.
3. Kaynak kapanış yılı/ayını girip **Konumları kaydet ve yenile** de. Yeni konumlarla kaynaklar doğrulanır; sonra yalnız üç konum hücresi ayar Excel'ine kaydedilir. Eşleştirme sayfalarına dokunulmaz. Önceki dosyanın birebir yedeği yanındaki Ayar_Yedekleri klasörüne alınır ve yolu ekranda gösterilir. Dosya Excel'de açık/kilitliyse kaydı kapatıp tekrar dene. Hatalı kaynak veya sonradan değişen ayar dosyasının üzerine yazılmaz.
4. Sonraki aylar konum değişmiyorsa dosyaları aynı klasöre ekle, kapanışı değiştir ve **Yalnız kaynakları yenile** düğmesine bas. Yeni dosyalar ve ayar Excel'indeki tüm kayıtlar yeniden okunur. Başarısız okumada çalışan önceki veri korunur. Ayar Excel'i yolu da sonraki açılışta hatırlanır.
5. Yeni bölüm/müdürlük/lokasyon eklerken Ayarlar'daki eşleştirmeleri, maliyet sayfasını, kontrol grubunu ve gerekiyorsa endeks eşleştirmesini tamamla. Yenile; listeler, satırlar, grafik serileri ve rapor dosya sayıları yeni yapıdan üretilir. Silinen veya adı değişen kayıtlar da yeniden okunur; uygulamada ayrıca sabit bir bölüm listesi tutulmaz. Kaynakta olup ayarda bulunmayan bölüm adıyla uyarılır; müdürlük/lokasyon/endeks ilişkisi tahmin edilmez. Ayarda olup kaynakta bulunmayan bölüm de eksik kaynak uyarısıyla görünür.
6. Bütçe yılı 2027 olduğunda kaynak kapanışını 2027 seçersin. Başlıklar, bütçe alanları ve 2026–2027 kıyasları aynı parametreden kayar. Kod değiştirilmez. Sonraki açılış son kullanılan kaynak/kapanışı hatırlar; tarih kendiliğinden takvim yılından alınmaz.

Kaynak penceresinde dosya adındaki yıl/ay kontrolünü ve her bölümün hangi dosyadan okunduğunu açabilirsin. Kapanış dosyası rapora temel olur; önceki kapanış dosyaları revizyon karşılaştırması içindir. Kaynaklar yalnız okunur. Excel formüllü kaynaklarda değerlerin Excel'de hesaplanmış ve kaydedilmiş olması gerekir; Python Excel'in hesap motorunu çalıştırmaz.

Bu pencerede bütün bölüm eşleştirmelerinin yedi alanı, ana gider grupları, gider birleştirme kayıtları ve dışlanan ortak sayfalar da gösterilir. Bunları ayar Excel'inde güncellediğinde yenileme ile tamamı yeniden kullanılır. “Tümü” seçimi yeni bölüm eklenince onu da kapsar; bilinçli olarak daralttığın bölüm seçimi ise kendi kapsamını korur.

### Konumlar ve kurallar

Ayar Excel'inin sayfaları: Bölüm Eşleştirme; Gider Grupları; Gider Eşleştirme; Ortak Sayfalar; Dosya ve Konum Ayarları.

- Gerçekleşen tarihleri BF59'dan sağa; giderler 65. satırdan itibaren. 60–63 personel, 64 işlem adedidir. **Hiçbir parasal toplama bu ilk beş adet satırı girmez.** Personel gideri adlı parasal gider kalemleri ise dahil edilir.
- Personel/işlem adları varsayılan olarak gider adlarıyla aynı sütundan, örnekte BE60:BE64'ten okunur. Başlık sütunu farklıysa Maliyet / gosterge_adi_sutunu ayarına sütun numarası eklenebilir; yoksa gider_sutunu değeri kullanılır. Alt gider adları kaynaktan; ana gider grupları ve birleştirilen giderlerin rapor adları ayar Excel'indeki tanımdan gelir.
- Bütçe EZ58 başlangıç ipucundan sağa aranan Bütçe etiketi ve altındaki kapanış yılının 12 aylık bloğuyla ayrılır. Gerçekleşenle bütçedeki yıl/ayın aynı olması karışıklık yaratmaz.
- Kontrol/ciro D18 başlığı, 19. satır değer; CW18 ipucunun sağındaki tek kapanış yılı bütçesi. Yıl geçişinde sağa kayabilir. Kontrol bölüm eşleştirmesinde çok bölüm tek kontrol adına toplanabilir.
- İkinci endeks tablosu C17, ilk tarih D17. İlk boş sütuna kadar tarihler okunur; sağdaki iki hazır oran sütunu kullanılmaz.
- Seçili ay endeksi = ilgili yılın her ayı / önceki yıl aynı ayı oranlarının aritmetik ortalaması. BM = seçili gider toplamı / pozitif seçili işlem adetleri toplamı. Reel iyileşme = (endekslenmiş önceki BM − güncel BM) / endekslenmiş önceki BM.
- Eksik değer sıfır sayılmaz. Eksik parçası olan toplam boş kalır. Sıfır/negatif işlem adetleri, eksik önceki yıl veya eksik endeks BM karşılaştırmasını kapsam dışında bırakır. Müdürlük ve genel ağırlıklar yalnız geçerli bölümlerin güncel giderlerinden hesaplanır.

## Excel raporları ve açıklamalar

**Rapor oluştur** penceresinde çıktı klasörü ve rapor yılını seç. Aylar üstteki ortak ay seçiminden gelir. Bu düğme ilk rapor aşamasını çalıştırır: **ayar dosyasındaki tüm bölüm/müdürlükler** için dosya üretir; üstteki organizasyon daraltması rapor kapsamını daraltmaz. Bu kapsam ekranda da yazılıdır.

Çıktılar seçtiğin konumda `Raporlar_tarih_saat` alt klasörüne kaydedilir. Beş dosyalık örnek, 1 genel ve 4 müdürlük dosyası oluşturur. İşlem sırasında kaynaklar yeniden okunur. Tüm dosyalar tamamlanınca klasör son adına taşınır; bu çalışmanın yarım dosyaları normal hata durumunda kaldırılır. İşlem sonunda tam klasör yolu ve dosya indirme bağlantıları görünür. Eski çıktılar ve kaynaklar değiştirilmez.

Dashboard'da girdiğin ve aynı yıl/kapanış/tam ay kümesine ait olan geçerli açıklamalar yeni raporların müdürlük Açıklama sütunlarına yazılır. Genel raporda Alt Gider Açıklamaları sayfasında `Bölüm adı: açıklama` biçiminde tek hücrede birleşir.

Önceden Excel'e yazılan açıklamaları da **Sapma öncelikleri → Açıklama yükle** ile alabilirsin. Müdürlük dosyaları veya birleştirilmiş genel dosya kabul edilir. Dönem, kapanış, ay kümesi, bölüm/gider adı ve tutarlar doğrulanır. Bir dosyada hata varsa hiçbir açıklama kısmen yüklenmez. Excel'deki zararsız sayısal hassasiyet farkları ayardaki kontrol toleransı içinde değerlendirilir.

**Rapor açıklamalarını birleştir**, mevcut akışı korur: genel dosyayı, tüm müdürlük dosyalarını ve çıktı klasörünü seçersin; yeni bir genel Excel oluşur. Kaynak dosyalar değiştirilmez. Yeni genel Excel'i ardından Açıklama yükle ile dashboard'a alabilirsin.

Uygulamada girilen açıklamalar `kullanici/kayitlar.sqlite3` içinde saklanır. Ay kümesi değişirse diğer dönemin açıklaması gösterilmez. İlgili gider/bütçe tutarı sonradan değişirse eski açıklama o tutara otomatik taşınmaz; yeniden girme bilgisi görünür.

## Doğrulama ve hata takibi

Sol alttaki **Doğrulama ve revizyonlar** bölümünde dip toplam karşılaştırmaları, önce/sonra/değişiklik tutarları ve kaynak uyarıları vardır. Dip toplam kontrolü birleşik bölüm eşleştirmelerini dikkate alır. Revizyonlar müdürlük, bölüm, alt gider, etkilenen ay ve yeni kapanışla gösterilir.

Açılış sorunu `Baslatma_Hatasi.txt`, iş hataları `kullanici/hatalar.log` dosyasına yazılır. Hata aktarırken yaptığın işlemi, kapanışı ve seçili ayları da belirt. Gerçek şirket verileriyle oluşan kayıt dosyaları şirket verisi içerebilir; dosya paylaşımı kendi şirket kurallarına bağlıdır.

Bu yerel sürüm tek kullanıcı içindir; başka bilgisayarlara sunucu olarak açılmaz. Yalnız `127.0.0.1` dinlenir, her açılışta ayrı oturum anahtarı üretilir. Kullanıcının seçtiği raporlar dışında klasör indirme/gezme servisi yoktur; dış sitelerden gelen yazma istekleri reddedilir. Ortak kurumsal sunucuya dağıtım bu paketin kapsamında değildir.

## Test durumu

V5'te Python test sayısı 93'tür (`python run_tests.py`); V4'te kaydedilen 72 JavaScript/API kontrolü arayüz değişmediği için geçerlidir. Canlı yerel API istekleri, kaynak başlıkları, eklenen/silinen/değişen eşleştirmeler, kaynak yollarının kalıcı kaydı, yedek/geri koruma, rapor üretimi, açıklama taşıma ve kayıtların korunması sınandı. JavaScript testi belge taklidiyle renderer ve olay işleyicilerini sınar; gerçek tarayıcı yerleşimi testi değildir.

**Bu ortamın tarayıcısı yerel Python işlemine bağlanamadı. Windows başlangıcı, yerel dosya pencereleri, gerçek tarayıcı yerleşimi ve fareyle tooltip davranışı henüz Windows üzerinde canlı doğrulanmadı.** Hesaplar ve rapor üretimi doğrulanmış pilot sürüm olarak teslim edilmiştir. [Canlı kontrol listesi](../docs/CANLI_KONTROL.md) ile kendi bilgisayarında son kontrolü yapabilirsin. Bu paket makro sürümünün Windows testlerini tamamlanmış saymaz.
