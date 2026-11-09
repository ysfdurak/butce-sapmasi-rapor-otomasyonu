> V3: çalışma merkezi ve teknik/test sınırları için V3_YENILIKLER.md ve Dogrulama_Durumu.json geçerlidir.

# Teknik yapı ve geliştirme noktaları

Bu uygulama önceki V11 rapor motorundan ayrı bir dağıtım olarak hazırlanmıştır. Mevcut Python/Excel ve bağımsız VBA rapor sürümleri değiştirilmez.

| Dosya / klasör | Sorumluluk |
|---|---|
| bootstrap.py | Paket içindeki saf Python wheel bağımlılıklarını kurulum yapmadan yükleme |
| app.py | Yalnız loopback üzerinde HTTP arayüzü, oturum, kaynak yenileme, iş kuyruğu, dosya indirme |
| model.py | Decimal ile ortak seçimler ve yedi sayfanın hesapları; Excel ekranından bağımsız |
| notes.py | SQLite açıklama kaydı, aynı ay/dönem/tutar eşleştirmesi, atomik dosya içe alma |
| configuration.py | Kaynak konumu önizleme, tüm eşleştirmeleri gösterme, doğrulanmış/yedekli/atomik ayar Excel'i güncelleme; eski Tümü seçimlerinin yeni kayıtlarla uzlaştırılması |
| jobs.py | Mevcut rapor motorunu çalıştırma, klasör olarak tamamlanmış çıktı yayımlama, açıklama taşıma |
| butce.py | V11 kaynak okuma, ayar okuma, kontrol/endeks, revizyon ve Excel rapor motoru |
| birim.py | Rapor motorunun birim maliyet hesapları ve bütçe işlem adedi okuması |
| personel.py | Personel bütçe tamamlama kuralı; rapor/dashbord eski fonksiyonları da korunur |
| web/index.html | Tek uygulama kabuğu |
| web/style.css | Ortak tasarım, duyarlı yerleşim, kaydırılabilir listeler ve tablolar |
| web/app.js | Gezinme, senkronize listbox'lar, SVG grafikler, görünüm modları, tooltip ve form etkileşimleri |
| dialog.py | Windows dosya/klasör seçimi; ayrı Python UI işlemi |
| tests/test_dashboard.py | Hesap, dönem, kapsam, kayıt, dosya üretimi ve HTTP testleri |
| tests/test_configuration.py | Kaynak başlıkları, yeni/silinen/değişen bölüm, konum kaydı, ayar yedeği, başarısız yazma koruması ve yeniden başlatma testleri |

Hesap doğruluğunun kaynağı Python `model.py` katmanıdır. JavaScript tutar hesabı yapmaz; Python'dan aldığı sonuçları biçimlendirir ve grafikleri çizer. Yalnız grafik eksen aralığı/yüzde dağılımının görsel yerleşimi tarayıcıda hesaplanır. Tutarlar sunucuda Decimal, JSON aktarımında sayıdır. Grafik, CSS ve JavaScript dosyaları yereldir; CDN yoktur.

Sayfa eklemek için modelin query sonucuna yeni sayfa verisi, JavaScript gezinmesine/render fonksiyonlarına sayfa eklenir. Kaynak konumları ve organizasyon eşleştirmeleri Excel ayarındadır; bölüm sayısı kodda sabit değildir. 2026 yalnız sentetik test verisidir. Grafikler DOM/SVG'dir; yeni bir grafik paketi kurmak gerekmez.

Bir isteğin kaynak yenilemesi/rapor işi sürerken önceki verinin yarım değiştirilmesi önlenir: yeni Dataset tamamen okununca değiştirilir. Dosya içe alınan açıklamalar da tüm doğrulamalar geçince tek SQLite işlemiyle yazılır. Aynı anda iki uzun iş başlatılmaz. Görüntü sorgularında eski yanıtın son seçimin üzerine yazılmasını engelleyen istek sıra numarası vardır.

Kayıt anahtarı kaynak ayar yolunun kimliği, kapanış ve tam seçili ay kümesidir. Açıklamalar ayrıca bölüm/gider kodu ve gerçekleşen/bütçe tutarlarıyla saklanır. Tutar değişince kayıt durur ama güncel açıklama olarak gösterilmez; kullanıcı yeniden değerlendirir.

Önceki rapor motorunun Excel içinde yer alan seçim paneli çıktısı korunmuştur. Bu Python dashboard'unun etkileşimli ekranı `.xlsx` değildir; ayrı yerel uygulamadır. Bu sürümde Power BI `.pbix` üretimi veya şirket içinde çok kullanıcılı web sunucusu dağıtımı yapılmaz.

## Geliştirici testi

Paket klasöründe `python run_tests.py` veya uygun Python sürümüyle `py -3.13 run_tests.py` çalıştırılır. Bağımlılıklar yine wheels klasöründen gelir. Testler sentetik veriyi okur; rapor çıktıları geçici klasörde oluşturulur ve temizlenir. Şirket kaynakları kullanılmaz. Ayrıntılı son test durumu `Dogrulama_Durumu.json` dosyasındadır.

JavaScript testleri teslimat sırasında Node V8 üzerinde belge taklidiyle ve gerçek yerel Python API'siyle çalıştırılmıştır. Gerçek tarayıcı görsel doğrulaması ve Windows dosya pencereleri için CANLI_KONTROL.md kullanılmalıdır. Bu ayrım test durumuna açıkça yazılmıştır.
