# Kalem ayrıntısı (V5.6)

Bir gider satırına tıklayınca o rakamın nereden geldiği tek pencerede açılır: bütçe, gerçekleşen, sapma, **hacim etkisi**, **birim maliyet etkisi**, bölüm kırılımı, açıklama durumu ve rakamı oluşturan **kaynak hücreler**.

| Sürüm | Durum |
| --- | --- |
| Python web arayüzü (`python/web`) | Uygulandı, otomatik testli (`python/tests/test_kalem.py`) |
| Tarayıcı demosu (`demo/`) | Uygulandı; Python arayüzüyle aynı dosyaları kullandığı için ayrıca kod yok |
| Excel VBA | **Öneri** — aşağıda. Windows Excel'de derlenip denenmedi |
| Excel Power Query | **Öneri** — aşağıda. Gerçek M motorunda ve Excel'de denenmedi |

Ortak karar gereği üç sürüm işlevde eşit tutulur. VBA ve Power Query karşılığı onaylanıp uygulanana kadar bu özellik **yalnız Python'da** vardır; bu fark [Ortak kararlar](ORTAK_KARARLAR.md) sayfasında açık madde olarak yazılıdır.

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

## VBA önerisi (denenmedi)

Makroda hesaplar `MakroAnaliz.VolumeRow` ile bölüm düzeyinde yapılıyor. Önerilen karşılık:

1. **MakroAnaliz'e `LineVolumeRow(department, code)`** — `VolumeRow` ile aynı adımlar; bölüm toplamı anahtarları (`"AT|"`, `"BT|"`) yerine `MakroHesap`ta zaten kullanılan kalem anahtarları (`SumMonths(d, "A" & Sep, y, Sep & code)`). Adet anahtarları (`AC|`, `BC|`) ve `Divide / Times / Minus` yardımcıları aynen kullanılır, böylece Null kuralları Python'la aynı kalır. Taslak:

   ```vba
   ' Taslak — Windows Excel'de derlenmedi.
   Public Function LineVolumeRow(ByVal department As String, ByVal code As Long) As Variant
       Dim d As Object, a As Variant, b As Variant, q As Variant, qb As Variant, ucb As Variant
       Dim volume As Variant, price As Variant
       Set d = CurrentSnap(department)
       If d Is Nothing Then LineVolumeRow = Null: Exit Function
       a = SumMonths(d, "A" & Sep, ReportYear, Sep & code)
       b = SumMonths(d, "B" & Sep, ReportYear, Sep & code)
       q = SumMonths(d, "AC|", ReportYear, "", True): qb = SumMonths(d, "BC|", ReportYear, "", True)
       ucb = Divide(b, qb): volume = Null: price = Null
       If Not IsNull(a) And Not IsNull(ucb) And Not IsNull(q) Then
           volume = Times(Minus(q, qb), ucb): price = Minus(a, Times(ucb, q))
       End If
       LineVolumeRow = Array(department, code, a, b, Minus(a, b), q, qb, Divide(a, q), ucb, volume, price)
   End Function
   ```
2. **Yeni `Kalem Ayrıntısı` sayfası** (genel ve müdürlük raporlarında): üstte bölüm ve gider seçimi (veri doğrulama listesi), altında aynı beş kart, köprü satırı, bölüm kırılımı ve kaynak hücre tablosu. Makro raporu üretirken tablo değerlerini yazar; kaynak hücre sütununa `Dosya!Sayfa!Hücre` adresi metin olarak gider.
3. **Tıklama:** Giderler sayfasında `Worksheet_BeforeDoubleClick` ile seçilen satırın bölüm/kodunu `Kalem Ayrıntısı` seçim hücrelerine yazmak. Bu olay kodu rapor dosyasına modül olarak eklenmez (rapor .xlsx makrosuz); bu yüzden asıl yol seçim listesi + formüldür, çift tıklama yalnız kontrol panelinde açıkken çalışır.
4. `MakroTest` sentetik testine Ücret giderleri satırı (yukarıdaki tablo) eklenir; Python testiyle aynı beklenen değer.

**Windows Excel'de denenmesi gerekenler:** VBA derlemesi, `Decimal` (`CDec`) yuvarlamasının kuruş farkı vermediği, veri doğrulama listesinin 27 kalem × 12 bölümle yavaşlamadığı, çift tıklama olayının kontrol panelinde tetiklendiği. Bu ortamda (Linux) yalnız `tools/vba_araclari.py` statik denetimi çalışabilir.

## Power Query önerisi (denenmedi)

1. `Butce_Veri` çıktısındaki kalem satırları bölüm × kod × ay gerçekleşen/bütçe tutarlarını ve işlem adetlerini zaten taşıyor. Şablona (`tools/pq_sablon.py`) **`Kalem Ayrıntısı`** sayfası eklenir: bölüm ve gider açılır listesi; bütçe, gerçekleşen, adetler `SUMIFS` ile; hacim ve birim maliyet etkisi yukarıdaki formüllerle hücrede. Makro gerekmez.
2. **Kaynak hücre adresi** şu an sorgu çıktısında yok. `Butce_Veri.pq` kalem satırına `Hücre` sütunu (ör. `DZ65`) eklenmesi gerekir; bu, `tools/pqtest/karsilastir.py` karşılaştırmasını ve PQTest beklenen çıktısını da değiştirir.
3. Doğrulama: `tools/pqtest/sablon_dogrula.py` LibreOffice hesabıyla Python `kalem.detail` sonucunu karşılaştıracak bir durum eklenir (Linux'ta çalışır). Gerçek M motoru testi Windows runner'da (`powerquery-motor` işi) çalışır.

**Excel'de denenmesi gerekenler:** açılır listeye bağlı `SUMIFS` formüllerinin büyük veri setinde yenileme süresi, `Hücre` sütunu eklendikten sonra eski şablonların yenilemede bozulmadığı.
