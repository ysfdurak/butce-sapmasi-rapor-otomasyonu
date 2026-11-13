# Şirket bilgisayarında deneme (Power Query yolu)

Amaç: şablonu ve iki sorguyu gerçek verinizle bir kapanış boyunca baştan sona çalıştırmak, sorunları erkenden görmek. Şirket verisini kimseyle paylaşmanız gerekmez: aşağıdaki **geri bildirim** satırları yalnız sayı, süre ve hata metnidir.

Süre: yaklaşık 1 saat (müdürlük dönüşleri hariç).

## 1. Hazırlık

- [ ] Excel sürümü: **Dosya › Hesap › Excel Hakkında** (ör. 2021, 64 bit).
- [ ] Ayarlar.xlsx'te kendi düzeninize göre konumlar, bölüm eşleştirmesi, gider grupları doğru.
- [ ] Maliyet klasöründe yalnız bu yılın kapanış dosyaları var (eski yıllar alt klasörde).
- [ ] Şablonun bir kopyası: `Butce_Rapor_2026-10.xlsx` gibi.

## 2. Ana sorgu (Butce_Veri)

- [ ] Rehberdeki kurulum adımları ([POWER_QUERY.md](POWER_QUERY.md)). Gizlilik sorulursa **Kuruluş**.
- [ ] İlk yükleme süresi (saatle): ……… sn. 60 sn'yi aşarsa `HizliMod = true` ile tekrar deneyin: ……… sn.
- [ ] Kontrol Paneli: Veri satırı ……… · Uyarı ……… · Doğrulanamayan kontrol ……… · Son yenileme dolu.
- [ ] **Girdi Kaydı**: okunan dosya sayısı ……… ; "Okundu" olmayan satır var mı?
- [ ] **Uyarılar**: ilk 5 uyarının metni (bölüm adlarını gizleyebilirsiniz).
- [ ] **Genel Özet** direktörlük toplamı, Python/makro raporunuzun veya muhasebenin toplamıyla aynı mı?
- [ ] **Veri Kontrolü**: tüm satırlar "Doğrulandı" mı? Değilse ilk farklı satırın durumu.
- [ ] Kontrol Paneli'nde Kapsam / ay seçimi değişince yeniden hesap süresi: ……… sn.

## 3. Müdürlük dosyası

- [ ] Dağıtım sayfasındaki adımlarla bir müdürlük dosyası hazırlayın. Kontrol Paneli "Müdürlük dosyası: …" gösteriyor mu?
- [ ] Dosyada başka müdürlüğün bölümü görünmüyor (Genel Özet, Giderler, Veri sayfası).
- [ ] Müdürlük Açıklamalar sayfasına 2–3 açıklama yazıp kaydetsin (veya siz deneyin).

## 4. Açıklama sorgusu (Butce_Aciklama)

- [ ] Dönen dosya `pq_aciklama` klasöründe; ana raporda **Tümünü Yenile**.
- [ ] Açıklama Onayı'nda satırlar geldi; Kontrol sütunu "Güncel".
- [ ] Bir satırı **Onaylandı** yapıp tekrar yenileyin: Durum korundu mu, İşlem geçmişine adınızla yazıldı mı?
- [ ] Giderler'de onaylı açıklama görünüyor; Önemli Sapmalar'da durum "Açıklandı".
- [ ] Excel'i kapatıp açın, tekrar yenileyin: durumlar hâlâ yerinde mi?

## Geri bildirim (bana yazın)

```
Excel sürümü:
Bölüm sayısı / maliyet dosyası sayısı / dosya boyutu (MB, yaklaşık):
İlk yükleme süresi (normal / hızlı mod):
Seçim değişince hesap süresi:
Veri satırı / uyarı sayısı / doğrulanamayan kontrol:
Toplamlar diğer raporla aynı mı:
Açıklama onayı yenilemede korunuyor mu:
Hata mesajları (aynen):
```
