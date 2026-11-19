# Mülakatta sorulabilecek sorular · Bütçe Sapması Rapor Otomasyonu

Bu kart, projeyi kendi ağzınla anlatabilmen için. Cevaplar kısa tutuldu; ayrıntısı belirtilen belgede. Kod yapay zekâ (Claude) ile yazıldı; kapsam, öncelik ve aşağıdaki kararlar senin. Bunu saklama, açıkça söyle: "Ürünü ve kararları ben belirledim, kodu yapay zekâyla yazdım, her kararın nedenini anlatabilirim."

**1. Bu proje ne yapıyor?**
Aylık kapanış dosyalarından bütçe sapma raporlarını üretiyor: önemli sapmalar, yıl sonu tahmini, hacim ve birim maliyet ayrıştırması, yönetici özeti. Python web arayüzü, Excel VBA ve Power Query sürümleri var.

**2. Hesabın doğruluğunu nasıl sağlıyorsun?**
Hesap tek yerde (Python model.py, Decimal ile); arayüz yalnız biçimlendiriyor. Panodaki toplamların Excel raporuyla aynı olduğunu her sürümde bir test denetliyor.

**3. Bu test neden var?**
'İlk bütçe' seçiliyken pano güncel bütçeyle, Excel raporu ilk bütçeyle hesaplıyordu. Hatayı bulup düzelttim ve tekrar etmesin diye 'pano = rapor' testini ekledim; eski kodda test düşüyor.

**4. 'Tahmin' ile 'bütçe' farkı?**
'Tahmin' yalnız yıl sonu tahmini için kullanılıyor; sapma her zaman bütçeye (güncel ya da ilk) göre. Karşılaştırma bazı her cümlede yazıyor.

**5. Sapmayı nasıl anlatıyorsun?**
İş diliyle: 'Gerçekleşen gider, güncel bütçeyi 17,61 milyon TL aştı (%20,54)'. Eksi işaretini kullanıcı çözmek zorunda kalmıyor.

**6. Hacim ve birim maliyet ayrıştırması?**
Sapmanın ne kadarı iş hacminden, ne kadarı birim maliyetten geliyor ayrılıyor; kalem ayrıntısında bölüm toplamına eşitliği test ediliyor.

**7. Teknik yapı?**
Python (yalnız loopback HTTP arayüzü), openpyxl, SQLite açıklama kaydı; tarayıcı demosu Pyodide ile aynı Python kodunu çalıştırıyor.

**8. Açıklamalar nasıl yönetiliyor?**
Sapma açıklamaları aynı ay, dönem ve tutar eşleşmesiyle taşınıyor, onay izi tutuluyor.

**9. Kimin için?**
Kontrol ve finans ekipleri; aylık kapanışta elle yapılan raporu dakikalara indiriyor.

**10. Neyi farklı yapardın?**
Pano ile rapor arasındaki tutarlılık testini ilk günden koyardım.


## Bu projede verdiğim kararlar (sorun → karar → bedel)

- **Tek hesap katmanı:** Arayüz ve Excel ayrışabilirdi → hesap yalnız Python'da → JavaScript tarafı hesap yapamıyor.
- **Pano = rapor testi:** Temel farkı gözden kaçmıştı → toplamlar her sürümde karşılaştırılıyor → örnek veri dışındaki durumları sınamıyor.

---
Son güncelleme: 19 Kasım 2026.
