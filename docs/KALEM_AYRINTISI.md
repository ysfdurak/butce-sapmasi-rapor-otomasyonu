# Kalem ayrıntısı (V5.6)

Bir gider satırına tıklayınca o rakamın nereden geldiği tek pencerede açılır: bütçe, gerçekleşen, sapma, **hacim etkisi**, **birim maliyet etkisi**, bölüm kırılımı, açıklama durumu ve rakamı oluşturan **kaynak hücreler**.

| Sürüm | Durum |
| --- | --- |
| Python web arayüzü (`python/web`) | Uygulandı, otomatik testli (`python/tests/test_kalem.py`) |
| Tarayıcı demosu (`demo/`) | Uygulandı; Python arayüzüyle aynı dosyaları kullandığı için ayrıca kod yok |
| Excel Power Query | **Uygulandı** — şablonda `Kalem Ayrıntısı` sayfası. LibreOffice hesabıyla Python'a karşı doğrulandı (fark 0); Windows Excel'de açılıp denenmedi |
| Excel VBA | **Uygulandı, denenmedi** — rapordaki `Kalem Ayrıntısı` sayfası ve `MakroTest` beklenen değerleri. Bu ortamda yalnız statik denetim (`tools/vba_araclari.py kontrol`) çalıştı; Windows Excel'de derlenip çalıştırılması gerekir |

Ortak karar gereği üç sürüm işlevde eşit tutulur. Üçü aynı kuralı ve aynı beklenen değerleri kullanır (aşağıdaki "Sentetik beklenen değerler"). Kalan farklar: kaynak **hücre adresi** yalnız Python'da; Excel sürümlerinde kaynak dosya adı ve aylık değerler görünür.

## Nereden açılır?

- **Detay durum:** Ana gideri ＋ ile açın, alt gider adına tıklayın. Seçimdeki bütün bölümlerin toplamı ve bölüm kırılımı gelir.
- **Sapma öncelikleri:** Satırdaki alt gider adına tıklayın. Yalnız o bölüm gelir.

Pencere ortak seçimleri (müdürlük, bölüm, yıl, ay) kullanır. Bütçe, sapma ve etkiler yalnız **bütçe yılı tek başına** seçildiğinde hesaplanır; birden fazla yıl seçiliyse yalnız gerçekleşen ve gerçekleşen kaynak hücreleri gösterilir.

## Hesap

Kurallar bölüm düzeyindeki **Hacim ve Birim Maliyet** sayfasıyla aynıdır (`analiz.volume_price_rows`); yalnız bölüm toplamı yerine tek gider satırına uygulanır.

| Değer | Formül |
| --- | --- |
| Sapma | kalem gerçekleşeni − kalem bütçesi (pozitif = **Aleyhte**) |
| Kalem bütçe birim maliyeti | kalem bütçesi / bölümün bütçe işlem adedi |
| Hacim etkisi | (gerçekleşen adet − bütçe adet) × kalem bütçe birim maliyeti |
| Birim maliyet etkisi | kalem gerçekleşeni − kalem bütçe birim maliyeti × gerçekleşen adet |

- İşlem adedi **bölümün** kaynaktaki işlem adedi satırıdır; kalemin kendi adedi yoktur. Bu yüzden aynı bölümdeki bütün kalemler aynı adet oranını kullanır.
- **Hacim + birim maliyet = sapma.** Bir bölümün bütün kalemlerinin etkileri toplanınca Hacim ve Birim Maliyet sayfasındaki bölüm satırı çıkar (testte kuruşuna kadar karşılaştırılır).
- Seçili aylardan birinde gerçekleşen veya bütçe adedi boş, sıfır ya da negatifse o bölümde etkiler boş kalır ve nedeni yazılır (bölüm sayfasıyla aynı kural).
- Detay durum tablosundaki "Bütçe − gerçekleşen" sütunu bu sapmanın ters işaretlisidir (geriye dönük uyum kararı, V5).
- Önemlilik etiketi Ayarlar › Rapor eşiklerini kullanır (tutar **ve** oran).

Sentetik örnek, Ekim 2026 kapanışı, Ocak–Ekim, 12 bölüm, **Ücret giderleri**:

| Bütçe | Gerçekleşen | Sapma | Hacim etkisi | Birim maliyet etkisi |
| ---: | ---: | ---: | ---: | ---: |
| 843.367.397 | 845.437.106 | 2.069.708 (Aleyhte) | 13.648.730 | −11.579.022 |

Okuma: işlem adedi bütçenin üstünde olduğu için ücret gideri ~13,6 milyon TL artmalıydı; adet başına maliyet bütçenin altında kaldığı için ~11,6 milyon TL geri kazanıldı. Net sapma 2,1 milyon TL aleyhte.

## Kaynak hücreler ve mutabakat

"Rakamı oluşturan kaynak satırlar" bölümü her bölüm × ay için dosya, sayfa ve hücre adresini gösterir: gerçekleşen, bütçe, gerçekleşen işlem adedi ve bütçe işlem adedi. Gider eşleştirmesiyle birden fazla kaynak satırı tek rapor kalemine birleşiyorsa aynı dönemde birden fazla hücre listelenir.

Hücre okuması "Hesap nasıl oluştu?" (hesap izi) ekranıyla aynı fonksiyondan gelir (`workspace.source_cells`). Pencere hücre toplamını hesaplanan tutarla karşılaştırır ve "tutuyor / tutmuyor" yazar. Kaynak dosya açıkken değişmişse pencere açılmaz; önce kaynakları yenilemek gerekir.

## Açıklama akışı

Her bölüm için mevcut açıklama ve durumu görünür: bu dönem açıklaması, önceki dönemden gelen (kontrol bekliyor / onaylandı), tutar değiştiği için geçersiz kalan veya önemli sapma olduğu hâlde açıklaması olmayan. "Açıklama ekle / düzenle" mevcut kayıt akışını kullanır (`/api/note`); açıklama yine tam seçili ay kümesine ve tutarlara bağlanır.

## Teknik

| Dosya | Görevi |
| --- | --- |
| `python/kalem.py` | Hesap: `detail(data, state, code, department, notes)` |
| `python/workspace.py` | `source_cells(...)` — hesap izi ve kalem ayrıntısının ortak hücre okuması |
| `python/app.py` | `POST /api/line-detail` `{state, code, department?}` |
| `python/web/app.js` | `showLineDetail`, `lineSources`, `lineNotes`; Detay durum ve Sapma öncelikleri satırlarındaki `data-line` düğmeleri |
| `python/tests/test_kalem.py` | Etkilerin sapmaya eşitliği, bölüm sayfasıyla uyum, mutabakat, çok yıl, eksik adet, geçersiz istek, açıklama akışı |

JavaScript hesap yapmaz; tutarlar Python'da Decimal ile hesaplanır.

## Sentetik beklenen değerler (örnek veri, Ocak–Ekim 2026)

| Durum | Gerçekleşen | Bütçe | Hacim etkisi | Birim maliyet etkisi |
| --- | ---: | ---: | ---: | ---: |
| Kalem 1, bütün bölümler | 845.437.105,52 | 843.367.397,13 | 13.648.730,22 | −11.579.021,83 |
| Kalem 3, Giriş | 17.033.694,15 | 16.915.214,06 | −1.760.939,53 | 1.879.419,62 |
| Kalem 7, Depolama müdürlüğü, Ocak + Mart + Mayıs | 26.858.415,48 | 32.195.176,16 | 2.610.900,32 | −7.947.661,00 |

Python (`kalem.detail`), Power Query şablonu (`sablon_dogrula.py`) ve VBA (`MakroTest.SentetikTest`, ilk iki satır) bu değerleri bekler. Etkilerin toplamı her durumda sapmaya eşittir (köprü farkı 0,00).

## Excel VBA (uygulandı, Windows'ta denenmedi)

- `MakroAnaliz.LineVolumeRow(bölüm, kod)` — `VolumeRow` ile aynı adımlar; bölüm toplamı yerine kalem anahtarları (`SumMonths(d, "A" & Sep, yıl, Sep & kod)`), adetler `AC|` / `BC|` (yalnız pozitif aylar). `Divide / Times / Minus` yardımcıları Null kuralını Python'la aynı tutar.
- `MakroAnaliz.WriteLineDetail` — raporda `Hacim ve Birim Maliyet`ten sonra **Kalem Ayrıntısı** sayfası: kapsamdaki her bölüm × gider kalemi bir satır (müdürlük, bölüm, kod, ad, kaynak dosya, gerçekleşen, bütçe, sapma, adetler, kalem bütçe BM, iki etki). Rapor dosyası makrosuz olduğundan seçim **filtreyle** yapılır; tablonun altındaki `SUBTOTAL(109, …)` satırı yalnız görünen satırları toplar ve köprü farkını gösterir.
- Bütçe yılı seçili değilse yalnız gerçekleşen yazılır, bütçe ve etkiler boş kalır (Python ve PQ ile aynı).
- `MakroTest.SentetikTest` yukarıdaki ilk iki satırı bekler.

**Windows Excel'de denenmesi gerekenler:** modüllerin derlenmesi, `SentetikTest`'in geçmesi, 12 bölüm × 26 kalem tablonun yazılma süresi, filtre + `SUBTOTAL` satırının beklendiği gibi toplaması. Çift tıklama ile açma yapılmadı: rapor dosyası makro taşımaz.

## Excel Power Query (uygulandı)

- Şablon (`tools/pq_sablon.py` → `build_kalem_ayrinti`) **Kalem Ayrıntısı** sayfası: `B3` gider kodu, `B4` bölüm (`Tümü` ya da kapsamdaki bölüm, açılır liste). Kapsam, yıl ve aylar Kontrol Paneli'nden gelir.
- Üstte özet (bütçe, gerçekleşen, sapma, %, yön, iki etki, köprü farkı); altında bölüm kırılımı (kaynak dosya, tutarlar, adetler, kalem bütçe BM, etkiler); tek bölüm seçiliyse aylık kaynak değerleri tablosu.
- Formüller `Veri` tablosundaki `Gider|bölüm|kod|yıl` satırından ve `_Bolum` yardımcı sayfasındaki adetlerden (`Hacim ve Birim Maliyet` ile aynı sütunlar) okur. Makro yoktur.
- Doğrulama: `python tools/pqtest/sablon_dogrula.py` dört durum çalıştırır (yukarıdaki üç satır + geçersiz kod). CI'da LibreOffice işi bunu çalıştırır; `powerquery-motor` işi sorguyu gerçek M motorunda dener (sorgu değişmedi).
- Kaynak **hücre adresi** sorgu çıktısında yok. Eklemek için `Butce_Veri.pq` kalem satırına `Hücre` sütunu gerekir; bu, PQTest beklenen çıktısını değiştirir. Şimdilik yapılmadı.

**Excel'de denenmesi gerekenler:** sayfanın Windows Excel'de açılış ve seçim değiştirme hızı (formüller 60 bölüm satırı için hazır).
