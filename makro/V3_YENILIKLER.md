# V3 — Makro çalışma merkezi

Bu paket bağımsız VBA **rapor** sürümüdür; Python gerektirmez. Önceki bağımsız rapor paketinde bulunmayan yedi sayfalık eski Excel dashboard üreticisini içermez. Windows masaüstü Excel gereklidir. Hazır derlenmiş xlsm burada üretilememiştir; paket içindeki Kurulum.vbs, Kontrol_Sablonu.xlsx ve dokuz VBA modülünden bilgisayarınızda **Butce_Kontrol_Makro_V3.xlsm** oluşturur. Şirketin güvenlik ayarlarını otomatik değiştirmez.

## Kurulum ve başlangıç

Paketi tamamen bir klasöre açın. Kurulum.vbs'ye çift tıklayın. VBA proje erişimi kapalıysa KULLANIM.md'deki elle yüklemeyi kullanın; bu sürümde vba_elle klasöründeki **dokuz** modülün tamamı (MakroMerkez dahil) içe alınmalıdır. Ardından Alt+F8 > PanelHazirla çalıştırılır. Çalışma Merkezi açılır; Dönem ve çıktı seçimi düğmesi mevcut Kontrol Paneli'ne gider.

V2'deki kullanım adımları aynıdır; sürüm adı geçen yerlerde V3 dosyasını kullanın. Eski xlsm'nin üzerine yazmayın. Kişisel kaynak/çıktı seçimlerinizi yeni panelde tekrar seçin. Yeni çalışma geçmişi bu kontrol dosyasında saklanır. Kontrol dosyasını kaydedin.

## Yeni işlemler

- **Kontrolleri yenile:** bütün kaynakları okur; bölüm/kaynak uyarıları, dip toplam farkları ve kapsama alınmayan BM/OPEX bölümlerini gösterir. İlgili dosya ve yapılacak işlem yazılır.
- **Eşleştirme değişiklikleri:** yeni, kaldırılan, organizasyonu/endeks eşleştirmesi değişen bölümleri son onayla karşılaştırır. D sütununa açıkça girilen eski bölüm adı, onay sırasında yeni ada aynı kalıcı kimliği bağlar. Eski adlar geçmişte tutulur. Açıklama birleştirmede bu onaylı adlar tanınır; gider adı/kodu, dönem ve tutar kontrolleri sürer. Python'daki otomatik metin benzerliği önerisi yerine eski ad elle seçilir/yazılır.
- **Hesap nasıl oluştu?:** D4 bölüm, D5 rapor gider kodu (boş = tümü) ile kaynak dosya/sayfa/hücre/değer tablosu, BM hesabı, aylık endeks oranları ve müdürlük/direktörlük ağırlıkları gösterilir. Ham birleşen giderlerin her satırı görünür.
- **Tablo düzeni:** parametreleri D sütununda düzenleyin. Kaynak hücresinden seç düğmesiyle ilgili parametre satırı için gerçek Excel dosyasında hücre/aralık işaretlenebilir. Aday ayar kaydedilmeden bütün kaynaklar okunur; ayar yedeği ve geri dönüş koruması vardır.
- **Kayıtlı görünümler:** bağımsız rapor sürümünün rapor yılı/çoklu ay/kapanış/kaynak ve çıktı seçimleri adlandırılır. Bu rapor sürümünde organizasyon dashboard filtreleri yoktur; raporlar ayar Excel'indeki tüm eşleşen bölümler için üretilir.
- **Çalıştırma geçmişi:** başlangıç, durum, kaynak, tam ay kümesi, girdi arşivi ve çıktı/hata kaydedilir. Bir satır seçip Kaynak değişikliklerini göster ile önceki aynı kaynak ayarlı çalışma karşılaştırılabilir.
- **Güvenli durdur:** devam eden okuma/yazma adımı bitince durur. Yeni raporlar tamamlanmadan çıktı klasörü yayımlanmaz. Bu düğme uzun tek bir Excel kaydetme işlemini ortasında kesmez.

## Girdi arşivi ve yeniden üretim

Kontrol dosyasının yanında Calisma_Arsivi klasörü oluşur. Kaynak ayarın özgün kopyası, yeniden üretim için yolları göreli hale getirilmiş ayar, ilgili kapanış/revizyon maliyetleri, ciro/kontrol ve endeks kopyaları tutulur. Dosya boyu/Adler imzaları kazara değişimleri kontrol eder; kriptografik güvenlik iddiası yoktur. İmza dosyası ile arşivleri birlikte koruyun. Kaynak arşivleri disk alanı kullanır.

Geçmişte bir satır seçip yeniden üret düğmesine basıldığında saklanan girdilerle, o dönemin ayları kullanılarak, paneldeki güncel çıktı konumuna yeni raporlar üretilir. Önceki raporlar değiştirilmez. Makro raporunda açıklamalar kullanıcıların çıktı dosyalarına yazdığı metinlerdir; ham girdilerden yeniden üretim açıklamaları eski çıktı dosyasından kendiliğinden kopyalamaz. O açıklamalı müdürlük raporları korunmalı, normal birleştirme adımıyla kullanılmalıdır. Farklı hesap motoru sürümüyle birebir aynı çıktı garantisi yoktur; eski paketle birlikte arşivleyin.

## Doğrulama sınırı

VBA kaynak/paket yapısı incelenmiştir; burada Windows Excel bulunmadığı için VBA derlemesi, düğmeler, kaynak hücresi seçimi, kaydetme/geri dönüş, iptal ve rapor üretimi çalıştırılarak doğrulanmamıştır. CANLI_KONTROL.md ve V3_CANLI_KONTROL.md bilgisayarınızda uygulanmalıdır.

Yeni otomatik davranış açıklaması ve ortak sayfadan eşik üstü gider detayı aktarımı henüz etkin değildir. ORTAK_KARARLAR.md açık kuralları içerir.
