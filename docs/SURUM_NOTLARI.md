# Sürüm notları

## V5.7.6 — Kasım 2026 (kurulum sonrası ilk açılış, anlaşılır hata metinleri)

- **Panel "Kurulum bekleniyor"da kalmıyor:** Kurulum.vbs, makrolar kapalı ya da VBA projesine erişim kapalıyken paneli hazırlayamıyordu; kullanıcı Alt+F8 › PanelHazirla'yı bilmezse ay listesi ve düğmeler hiç gelmiyordu. Artık `MakroPanel.Auto_Open` dosya her açıldığında ay listesinin varlığına bakar, yoksa paneli kendisi hazırlar; olmazsa nedenini durum satırına yazar. Windows Excel'de denendi: liste silinip dosya kapatıldı, normal açılışta panel kendiliğinden kuruldu, durum "Hazır" oldu.
- **Hata metinleri yerini söylüyor:** Giderler sayfası okunurken çıkan hatalar artık dosya, sayfa, hücre ve başlığı yazar (ör. "Aynı bölüm iki kez okunuyor: "Giriş" (… › Giderler, F5 "Bütçe - Gerçekleşen TL")"). Önceki metin yalnız çalışma kitabı adını ("Sheet9") veriyordu.
- **Not:** Bu bilgisayarda Excel'in "VBA proje nesne modeline güven" ayarı kapalı olduğu için Kurulum.vbs modülleri ekleyemedi ve elle kurulumu önerdi (beklenen, ayar değiştirilmedi). Elle kurulum (vba_elle'den içe aktarma) çalıştı.

## V5.7.5 — Kasım 2026 (makro sürümü gerçek Excel denemesi)

- **Makro raporu hiç oluşmuyordu (düzeltildi):** Müdürlük raporlarının Giderler sayfasında her bölüm için "… Gerçekleşen TL", "… Bütçe TL", "Bütçe - Gerçekleşen TL" başlıkları var. Açıklama okuyucusu (`MakroAciklama.ExpensePositions`) gerçekleşen sütununu başlıkta "gerceklesen" geçmesine bakarak buluyordu; fark sütunu da bu sözcüğü taşıdığı için her bölüm iki kez sayıldı ve "Rapor oluşturulamadı… Geçersiz bölüm başlığı: SheetN" hatası çıktı. Hata V5.7.1'deki başlık değişikliğiyle geldi; statik denetim ve Python testleri VBA'yı çalıştırmadığı için görünmedi. Koşul artık "bütçe" ile başlayan başlığı dışarıda bırakıyor. Test: `python/tests/test_makro_gider_basliklari.py`.
- **Gerçek Excel'de doğrulanan (Microsoft 365, Windows, 24 Kasım 2026):** örnek ayarlarla Ocak–Ekim 2026; direktörlük raporu + 4 müdürlük raporu üretildi. Gerçekleşen 3.449.383.390,31 TL, bütçe 3.440.938.980,13 TL (Power Query ve Python ile aynı, kuruşu kuruşuna). Ciro kontrolleri "Doğrulandı"; 96 uyarı örnek verideki geriye dönük düzeltmeler (Şubat rakamlarının Ekim dosyasında değişmesi), beklenen.

## V5.7.4 — Kasım 2026 (gerçek Excel denemesi)

- **Power Query şablonu gerçek Excel'de çalışmıyordu (düzeltildi):** Şablonun Veri sayfasında, tablonun sağında AL1 hücresinde bir açıklama notu vardı. Tablo 36 sütunken sorun yoktu; R1..R12 (ilk bütçe) sütunları eklenince tablo 48 sütuna çıktı ve nota çarptı. Excel `Veri!$A$1`'e yüklerken yer açmak için 11 sütun ekledi, şablondaki bütün `Veri!A:A` başvuruları `Veri!L:L`'ye kaydı; Kontrol Paneli "Veri yok" dedi, raporlar boş kaldı. CI şablonu LibreOffice'te veriyi doğrudan yazarak denediği için bunu görmüyordu. Not kaldırıldı (açıklaması Kontrol Paneli › Sayfalar'a taşındı); test boş şablonun Veri sayfasında tek bir dolu hücre olmamasını şart koşuyor.
- **Gerçek Excel'de doğrulanan (Microsoft 365, Windows, 23 Kasım 2026):** örnek veriyle 3.379 satır; Ekim 2026, Ocak–Ekim, Tümü için gerçekleşen 3.449.383.390,31 TL ve bütçe 3.440.938.980,13 TL (beklenen değerler); doğrulanamayan kontrol 0. Butce_Aciklama: aynı dakikada (14:11:13 ve 14:11:24) kaydedilmiş iki dosyadan gerçekten yeni olanı seçildi, çakışma notu yazıldı (V5.7.3 düzeltmesi); "Onaylandı" yenilemeden sonra korundu ve işlem geçmişine eklendi; şablondan hazırlanmamış dosyalar "Dosya okunamadı" satırı olarak listelendi.

## V5.7.3 — Kasım 2026 (denetim D19)

- **Açıklama kaynağı seçimi:** Aynı satır iki dönen dosyada varsa Power Query "en yeni" dosyayı dakikaya yuvarlanmış metinle seçiyordu; aynı dakikada gelen iki dosyada adı alfabetik sonra gelen eski dosya alınabiliyordu. Artık özgün değiştirilme zamanı karşılaştırılır, yalnız tam eşitlikte dosya adı karar verir. Farklı metinler için çakışma uyarısı aynen kalır. `powerquery/Butce_Aciklama.pq` ve Python benzetimi (`tools/pq_aciklama.py`) birlikte değişti; test `python/tests/test_pq_aciklama_secim.py`.

## V5.7.2 — Kasım 2026 (gerçek Excel testinden düzeltmeler)

Windows'ta gerçek Excel'de uçtan uca deneme sırasında bulunan sorunlar:

- **Derleme hatası:** `local` VBA'da değişken adı olamıyor; MakroAnaliz, MakroHesap ve MakroYaz derlenmiyordu. Değişken `yerelKume` oldu.
- **Atlanan çağrılar:** Satır başında `LoadSettings: LoadSnapshots` gibi yazılan satırlarda VBA ilk adı etiket sayıyor ve o prosedürü hiç çalıştırmıyordu (7 yer: ayar okuma, kontrol hesapları, panel okuma). Bu satırlar `Call` ile başlıyor.
- **Denetim aracı:** `tools/vba_araclari.py kontrol` bu iki hatayı artık yakalar (ayrılmış sözcükle değişken adı; prosedür adıyla başlayan etiket satırı). CI her gönderimde çalıştırır.
- **Konum bulma hata metni:** Maliyet konumu dosyadan bulunamazsa mesaj hangi dosya ve sayfada, hangi adımda takıldığını söyler (ör. "İşlem Adedi satırı yok"). Sentetik testte konum testi düşerse diğer testler yine çalışır ve sonuç listesinde görünür.
- **Görünüm:** Test Sonuçları'nda tarih seri sayı yerine `gg.aa.yyyy ss:dd`; Çalışma Merkezi açıklaması ve Kontrol Paneli durum metni kesilmiyor.
- **Belge:** KULLANIM.md kontrol listesinde modül sayısı on üç.

## V5.7.1 — Kasım 2026 (tek kalemin hesap izi)

- **Hesap izi:** Kalem ayrıntısı penceresinde yeni tablo: gerçekleşen (kaynak dosya ve kapanış dönemi, toplanan aylar), **seçili bütçe** (güncel mi ilk mi, hangi dosya ve dönem), fark ve fark %. Power Query şablonunun Kalem Ayrıntısı sayfası da karşılaştırma bütçesini yazar (B6).
- **Düzeltme:** Karşılaştırma bütçesi "İlk bütçe" seçiliyken kaynak hücre listesi ("Hesap nasıl oluştu?" ve kalem ayrıntısı) bütçe hücrelerini kapanış dosyasından okuyordu; tutar ilk bütçeyken hücreler güncel bütçeydi ve mutabakat "tutmuyor" diyordu. Artık bütçe hücreleri bölümün yılın ilk kapanış dosyasından okunur. Test: `tests/test_kalem.py › LineTraceBasisTests` (ilk dosyadaki bütçe değiştirilerek; düzeltme geri alınınca test düşer).
- **Python ve Excel yan yana:** [KALEM_IZI.md](KALEM_IZI.md) iki kalemi (Giriş · kalem 3, Sayım · kalem 1) adım adım ve ay ay iki sürümde gösterir; `tools/pqtest/kalem_izi.py` her gönderimde çalışır, fark ya da eskimiş belge varsa iş düşer.

## V5.7 — Kasım 2026 (Python web arayüzü ve tarayıcı demosu)

- **Sapma iş diliyle yazılır:** Sapma öncelikleri tablosunda eksi işaretli "Sapma" ve "Sapma %" sütunları yerine tek bir **Durum** cümlesi var: "Bütçeyi 17,61 milyon TL aştı (%20,54)". Panodaki hesap `bütçe − gerçekleşen` olduğu için eksi işaret aşım demekti; Excel raporlarındaki `gerçekleşen − bütçe` ve Lehte/Aleyhte kuralıyla karışıyordu. Genel durum kartında da "bütçe aşıldı / bütçenin altında" notu çıkar. Hesap değişmedi.
- **Genel durumda en büyük 3 aşım:** İlk ekranda etkin dönemle birlikte en büyük üç bölüm–alt gider aşımı görünür; aynı satırdan kaynak kalemler ve açıklama açılır, "Tüm aşımlar" düğmesi Sapma öncelikleri sayfasına gider.
- **Filtreler kapalı başlar:** Ortak seçimler tek satırlık özetle (etkin dönem, bölüm sayısı, daraltılan filtreler) görünür; "Filtreleri aç" tek tıkla altı filtreyi açar, tercih tarayıcıda hatırlanır.
- Makro ve Power Query raporları değişmedi: orada sapma zaten `gerçekleşen − bütçe` ve Lehte/Aleyhte sözcüğüyle yazılıyor.
- **Karşılaştırma bazı her yerde yazılı ve raporlarla aynı:** Cümle "Gerçekleşen gider, güncel bütçeyi 17,61 milyon TL aştı (%20,54)" biçiminde; panel alt başlığı ve filtre özeti bazı söyler (güncel bütçe: kapanış dosyası; ilk bütçe: yılın ilk kapanış dosyası). Düzeltme: Ayarlar › `karsilastirma_butcesi` "İlk bütçe" seçiliyken Excel raporları ilk bütçeyle, pano ise güncel bütçeyle hesaplıyordu; pano artık aynı ayarı uygular (`model.py`, test: `tests/test_surum.py` › `DashboardBasisTests`). Revizyon karşılaştırması kaynak dosyalardan yapılmaya devam eder.
- Terim notu: "tahmin" sözcüğü yalnız **yıl sonu tahmini** için kullanılır; sapma her zaman bütçeye (güncel ya da ilk) göredir.

- **Pano = rapor testi:** `tests/test_surum.py › DashboardEqualsReportTests` panodaki toplam gerçekleşen ve bütçeyi, Excel rapor motorunun (`butce.run` ile aynı adımlar) satır toplamlarıyla karşılaştırır; güncel ve ilk bütçe temelinde. Panonun temeli yok saydığı Kasım 2026 hatasını bu test yakalar (geri alınan düzeltmeyle düştüğü doğrulandı).
- **Temel açıklaması:** Karşılaştırma bazı metni artık ne demek olduğunu da söylüyor (ilk bütçe: yıl başında onaylanan; güncel bütçe: yıl içindeki son revizyon); öncelik sayfası değiştirme yerini gösterir (Ayarlar › Rapor › Karşılaştırma bütçesi).
## V5.6 — Kasım 2026

- **Kalem ayrıntısı** (Python web arayüzü ve tarayıcı demosu): Detay durum veya Sapma öncelikleri ekranında alt gider adına tıklayınca bütçe, gerçekleşen, sapma, hacim etkisi, birim maliyet etkisi, bölüm kırılımı, açıklama durumu ve rakamı oluşturan kaynak hücreler (dosya, sayfa, hücre) açılır. Hücre toplamı hesaplanan tutarla karşılaştırılır. [Ayrıntı](KALEM_AYRINTISI.md)
- **Kalem Ayrıntısı sayfası — Power Query ve makro:** PQ şablonunda gider kodu ve bölüm seçilen formüllü sayfa; makro raporunda her bölüm × kalem satırı, filtre ve görünen satır toplamı. Makro sürümü Windows Excel'de henüz denenmedi.
- VBA ve Power Query karşılığı öneri olarak yazıldı; Windows Excel'de denenince eklenecek.
- Arayüz: iki küçük metnin kontrastı WCAG AA'ya yükseltildi (sıra rozeti, kenar çubuğu alt notu).
- Erişilebilirlik: ortak seçim listeleri onay kutusu grubu olarak işaretlendi (iç içe etkileşimli öğe uyarısı giderildi); kaydırılabilir tablo ve geniş grafikler klavyeyle odaklanabiliyor. Bütün sayfalarda axe (WCAG 2 A/AA) hatası yok.

## V5.5 — Kasım 2026

- **Konumlar otomatik bulunur** (Python, makro, Power Query): Ayarlar › Dosya ve Konum Ayarları'ndaki Maliyet, Kontrol ve Endeks konum satırları artık isteğe bağlıdır. Satır yoksa konum dosyadan bulunur ve Uyarılar'a yazılır; Ayarlar'da değer varsa o kullanılır. Eski Ayarlar dosyaları değişmeden çalışır. [Kurallar](VERI_VE_HESAP.md#konumların-otomatik-bulunması-v55)
- Ayardaki konum bir maliyet sayfasında tutmazsa (satır/sütun eklenmişse) o sayfa için konum yeniden bulunur, sayfa okunur ve uyarılır. Power Query zaten "Bütçe" başlığına göre hizalanıyordu.
- Python Tablo düzeni ekranında girilen konum kaynakla uyuşmazsa kayıt reddedilir.
- Makroda yeni modül **MakroKonum** (toplam 13).

## V5.4 — Kasım 2026

Üç sürümde (Python, makro, Power Query) aynı kurallarla ([rehber](POWER_QUERY.md), [hesap](VERI_VE_HESAP.md#v54-bütçe-sürümleri-ve-trend)):
- **Hızlı mod (yalnız Power Query):** `HizliMod = true` yalnız kapanış, önceki ay ve yılın ilk dosyasını okur; revizyonlar yalnız son ay için.
- **Bütçe tabloları:** Bütçe - Bölüm / Müdürlük / Direktörlük sayfaları (güncel bütçe ay ay, yıl toplamı, ilk bütçeye göre fark).
- **Bütçe Değişimi:** önceki ay kapanış dosyasına göre bütçesi değişen satırlar ve kalemler.
- **İlk / güncel bütçe:** ilk bütçe = yılın ilk kapanış dosyasındaki bütçe; karşılaştırma bütçesi Ayarlar › Rapor › `karsilastirma_butcesi` ile (Power Query'de Kontrol Paneli'nden de, Python'da `--butce ilk` ile de) seçilir; **Bütçe Sürümleri** sayfası.
- **Trend:** aylık sapma tablosu ve grafiği; üst üste aleyhte giden kalemler listesi; Yönetici Özeti'nde sayısı.
- Şirkette ilk deneme için [kontrol listesi](PQ_DENEME.md).
- **Girdi Kaydı** sayfası: okunan dosyalar, dönemleri, boyut ve değiştirilme zamanı.
- Python ve makro bu özelliklere eşitlendi (makroda yeni modül **MakroSurum**, toplam 12 modül). Hızlı mod Power Query'nin yenileme süresini kısaltmak içindir; Python ve makro dosyaları zaten tek seferde okur.

## V5.3 — Kasım 2026

Makrosuz yolun eksikleri tamamlandı ([rehber](POWER_QUERY.md)):
- **Müdürlük dosyaları:** `Butce_Veri` sorgusuna `MudurlukDosyasi` satırı; dosyaya yalnız o müdürlüğün verisi yüklenir. Dağıtım sayfası müdürlükleri ve dosya adlarını listeler.
- **Açıklama toplama:** yeni `Butce_Aciklama` sorgusu dönen müdürlük dosyalarını Açıklama Onayı sayfasına toplar; kapanış, ay seçimi ve tutar değişikliği kontrolü.
- **Onay izi:** Durum (Onaylandı / Reddedildi / Bekliyor / Bu döneme aktar) ve onay notu yenilemede korunur; her işlem kişi ve zamanla satırın işlem geçmişine eklenir; kaynak dosyası kaldırılan açıklama silinmez.
- **Girdi kaydı:** okunan dosyaların boyutu ve değiştirilme zamanı; Kontrol Paneli'nde son yenileme.

## V5.2 — Kasım 2026

- **Makrosuz yol tam kapsamlı:** `powerquery/Butce_Rapor_Sablonu.xlsx` + yeni `Butce_Veri.pq`. Python/makro sürümündeki raporların tamamı (Yönetici Özeti, Pano, Genel Özet, Giderler, Birim Maliyet, Önemli Sapmalar, Yıl Sonu Tahmini, Hacim ve Birim Maliyet, Veri Kontrolü, Revizyonlar, Uyarılar) formüllerle üretilir. Kapsam, ay ve eşik seçimleri Kontrol Paneli'nden anında; yeni dosyalar için **Veri › Tümünü Yenile**. Excel 2016 ve sonrasıyla uyumlu. [Rehber](POWER_QUERY.md)
- **CI'da gerçek Power Query motoru:** sorgu Microsoft Power Query SDK (PQTest) ile çalıştırılıp Python benzetimiyle satır satır, şablon formülleri LibreOffice'te hesaplanıp Python raporlarıyla karşılaştırılır.
- Eski pivot tabanlı sorgunun sütunları değişti; şablona yeni sorgu yapıştırılmalıdır.

## V5.1 — Kasım 2026

- **Pano** sayfası (Python ve makro): açılır liste filtreleri, toplam kartları, aylık / gider grubu / bölüm tabloları ve iki grafik. Formül tabanlı; Power BI veya makro gerektirmez.
- **Makrosuz yol:** `powerquery/Butce_Veri.pq` tek Power Query sorgusu ve [kurulum rehberi](POWER_QUERY.md). Python ve makro kullanılamayan bilgisayarlarda dilimleyicili PivotTable panosu kurulur.

## V5 — Kasım 2026

Uzman incelemesindeki revizyonlar Python ve Excel VBA sürümlerine eşdeğer uygulandı.

**Yeni rapor sayfaları (iki sürümde de)**
- **Yönetici Özeti** ilk sayfadır ve tek sayfaya yazdırılır (PDF'e hazır): toplamlar, Lehte/Aleyhte yönü, yıl sonu tahmini, veri kontrolü, en büyük aleyhte/lehte önemli sapmalar.
- **Önemli Sapmalar**: tutar **ve** oran eşiği (Ayarlar.xlsx › `Rapor`). Müdürlük dosyalarında açıklanması gereken hücreler turuncu; açıklama durumu formülle izlenir.
- **Yıl Sonu Tahmini**: gerçekleşen + kalan bütçe ve aylık ortalama gidiş yöntemleri.
- **Hacim ve Birim Maliyet**: bütçe farkının hacim etkisi ve birim maliyet etkisi olarak ayrıştırılması.
- Açıklamalarda **onay izi**: son işlem, işlemi yapan kullanıcı ve zaman; silinmeyen işlem günlüğü.

**Hata düzeltmeleri**
- VBA: `Put` adlı prosedür (VBA'nın ayrılmış komutu) `SetCell` olarak değiştirildi; proje derlenmesini engelliyordu.
- VBA: koşullu biçimler dil bağımsız hücre değeri kuralıyla yazılıyor (İngilizce formül metni Türkçe Excel'de hata verebiliyordu).
- VBA: hesap izi ve tablo düzeni ekranları sayfa adını yükleme ile aynı kuralla (büyük/küçük ve Türkçe harf farkı gözetmeden) buluyor; müdürlük adları da aynı kuralla eşleşiyor.
- VBA: rapor hatasında Excel hesaplama/ekran ayarları her durumda geri yükleniyor; ayar dosyasında eksik `dip_toplam_satiri` veya sayı olmayan gider aralığı artık bağlamsız "Type mismatch" vermiyor.
- VBA: geçmiş çalışmayı yeniden üretmek paneldeki seçimleri değiştirmiyor.
- İki sürüm: Ocak kapanışında önceki yılın Aralık dosyasıyla yıl dönümü revizyonları gösteriliyor.
- Python: bütçe işlem adedi okuması sayfa adını Türkçe harf farkı gözetmeden buluyor.

**Sağlamlık ve kullanım**
- Metin olarak girilmiş tutarlar çevriliyor ve hücre adresiyle uyarılıyor; formül sonucu kayıtlı olmayan kaynak hücreleri uyarılıyor; kaynak düzeni kaydığında yeni başlangıç hücresi öneriliyor.
- Python rapor üretimi yaklaşık 12 kat hızlandı (Ekim örneğinde ~120 sn → ~10 sn).
- VBA: kaynaklar ve rapor satırları blok hâlinde okunup yazılıyor; dosya imzaları önbelleğe alınıyor; yalnız üretilen rapor hesaplanıyor.
- `Kurulum.vbs` Türkçe; mevcut xlsm'yi yedekleyip güncelleyebiliyor, V4'ten yükseltebiliyor ve kod sayfası Türkçe değilse uyarıyor.
- `Baslat.bat` Python bulunamazsa kurulum yollarını açıklıyor.
- Girdi arşivi son 20 çalışmayı tutuyor (`arsiv_saklama_adedi`).
- Konumlu düzende eski satır/sütun ayarları zorunlu değil; örnek Ayarlar.xlsx'te bu ayarlar "eski düzen yedeği" olarak işaretli.

**Bakım**
- Python ve web kodu biçimlendirildi (black, prettier); kullanılmayan eski dashboard kodu kaldırıldı.
- `vba_elle/` kopyaları `tools/vba_araclari.py` ile üretiliyor ve denetleniyor; GitHub Actions Python testlerini ve VBA denetimini çalıştırıyor.
- V2–V4 sürüm belgeleri `docs/arsiv/` klasörüne taşındı; ortak belgeler `docs/` altında tek kopya.

## V4

Önceki kapanış yorumlarını inceleme kuyruğuna alma, toplu düzenleme ve onay, kaynak kapanışın görünmesi ve mevcut dönem yorumlarının korunması eklendi. Ayrıntılar: [arşiv](arsiv/).

Gider verisinden faturaya inen ayrıntı çıkarımı, otomatik neden metni üretimi ve yerel dil modeli deneyi bu paketin çalışma akışında yer almaz ([açık talepler](ORTAK_KARARLAR.md)).
