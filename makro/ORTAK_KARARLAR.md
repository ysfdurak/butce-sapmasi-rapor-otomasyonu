# V4 — Geçerli ortak kararlar

- Ortak işlev değişiklikleri ayrıca söylenmeden hem Python hem VBA sürümüne uygulanacak; teknik karşılıklar ve test sınırları açıkça belirtilecek.
- Rota güncellemesi durdurulmuştur.
- Bir önceki takvim rapor kapanışındaki açıklamalar kontrol bekleyen taslak olarak getirilecek. Kaynak ay kümesi farklı olabilir; fark görünür olacak. Otomatik onay olmayacak.
- Bu dönemin mevcut açıklaması ve düzenlenen taslaklar yeniden getirmeyle ezilmeyecek.
- İnceleme tek tabloda; seçili satırları topluca onaylama/geçersiz sayma, doğrudan metin düzenleme ve seçili satırlara ortak metin uygulama olacak.
- Onaysız yorumlar raporda kırmızı ve kaynak dönem/ay bilgili etiketle kalacak. Onaylılar normal renk, geçersizler çıktı dışında olacak. Eski metin ve karar kaydı korunacak.
- İki önceki müdürlük dosyasında 10 yorumluk test: 2 aynen, 8 düzenlenerek onaylanır. Bu senaryo iki pakette de bulunur.

## Açık ikinci aşama

- Fatura detayını rapora ayrı sayfaya taşıma henüz etkin değildir. Ortak Gider Datası sayfasındaki çift tıklamanın hangi veriyi, nereden getirdiği örnek dosyada incelenecek.
- Detay aktarımında bütçe aşımı TL eşiği VE yüzde eşiği birlikte aşılmalı. Eşiklerin değerleri ve uygulanacak kalem listesi açık.
- Detay kaynağı olan ortak sayfa ana toplamlara dahil edilmeyecek. Bölüm/gider/dönem eşleştirmesi, tutar mutabakatı ve mükerrer kontrolü tasarlanacak.
- Tahakkuk veya önceki yıl fatura eksikliği için açıklama kuralları ve metinleri henüz kabul edilmedi; kanıt olmadan kesin neden veya “gelecek ay görünmeyecek” yazılmayacak.
- Kullanıcıya özgü benzer davranışın otomatik tespiti ve yeni açıklama üretimi bu sürümde uygulanmadı.

## Önceki sürüm karar kaydı

# Ortak geliştirme kararı

Kullanıcı 8 Kasım 2026: Bundan sonraki ortak işlev değişiklikleri ayrıca söylenmeden hem Python hem VBA sürümüne uygulanacak. Teknik olarak aynı biçimde yapılamayan bir işlevin karşılığı ve test sınırı açıkça belirtilecek. Rota güncellemesi halen durdurulmuştur.

## Açık kurallar: otomatik açıklama ve gider detayları

- Benzer davranışın tanımını kullanıcı kendi örnekleriyle tarif edecek. Ardışık 3 ay gibi bir kural kabul edilmedi ve uygulanmadı.
- Otomatik açıklama mevcut kullanıcı açıklamasının üzerine yazmamalı. Önerilen tasarım: ayrı otomatik metin ve kullanıcı açıklaması; çıktı hücresinde açık etiketlerle birlikte gösterim. Henüz kullanıcı onaylı hesap kuralı değildir.
- Detay aktarımı: bütçe aşımı TL eşiği VE yüzde eşiği birlikte aşılmalı (kullanıcı seçimi).
- Eşiklerin bölüm/kalem/toplam hangi düzeyde uygulanacağı, sayısal değerleri, ay bazında mı seçili ay toplamında mı kontrol edileceği açık.
- Detay aktarımına izin verilen gider kalemleri ayarlardan seçilebilmeli.
- Ana okumaya dahil edilmeyen ortak sayfa, yalnız detay kaynağı olarak ayrıca tanımlanacak; ana toplamlara katılmayacak.
- Ortak sayfadaki tıklamanın PivotTable ayrıntı gösterimi, bağlantı, mevcut detay tablosu veya başka bir mekanizma olup olmadığı örnek dosyayla netleştirilmeli.
- Detay anahtarları (bölüm, alt gider, yıl, ay, gerçekleşen/bütçe), alan adları, tutar mutabakatı ve mükerrer önleme kuralları açık.
- Veri zaten kaynak kitapta bulunuyorsa Python doğrudan okuyabilir. Pivot/OLAP tıklaması dış bağlantıyla veri getiriyorsa dosya içeriği tek başına yeterli olmayabilir; yerel Excel ile üretim veya önceden dışa aktarım gerekir. Bu ayrım dosya görülmeden çözüldü sayılmaz.
- İlgili detaylar kapsamı uygun rapor dosyasında ayrı bir sayfaya taşınacak; müdürlük dosyası diğer müdürlük ayrıntılarını içermeyecek.
- Bu iki yeni işlev V3'te henüz etkin değildir. Kurallar netleştiğinde her iki uygulamaya aynı iş kurallarıyla eklenecek.
