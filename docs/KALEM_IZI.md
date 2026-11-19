# Kalem hesap izi: Python ve Excel yan yana

Bu sayfa `tools/pqtest/kalem_izi.py --yaz` ile üretilir; elle düzenlemeyin. Aynı sentetik örnekte tek gider kalemi
için iki sürümün verdiği tutarları satır satır koyar. Excel tarafı Power Query şablonudur
(`powerquery/Butce_Rapor_Sablonu.xlsx`, Kalem Ayrıntısı sayfası); LibreOffice ile yeniden hesaplanır.
GitHub Actions her gönderimde aynı karşılaştırmayı çalıştırır, fark çıkarsa iş düşer.

**Nasıl okunur?** Bir yönetici "bu kalemde hangi bütçeye göre, ne kadar aştık?" diye sorduğunda cevap şu sırayla
izlenir: gerçekleşen (hangi dosya, hangi kapanış), seçili bütçe (güncel mi ilk mi, hangi dosya), fark ve yüzde.
Ay ay tablo, özet rakamın kaynak dosyadaki aylık değerlerin toplamı olduğunu gösterir.

**Kapsam dışı:** Makro (VBA) sürümü aynı beklenen değerleri `MakroTest.SentetikTest` ile kontrol eder; bu ortamda
Windows Excel olmadığı için makro çıktısı bu tabloya konmadı. Windows'ta makro testini çalıştırıp sonucu
`makro/Dogrulama_Durumu.json` dosyasına yazmak kullanıcıya kalan iştir.

**İlk bütçe bazı:** Karşılaştırma bütçesi "İlk bütçe" seçildiğinde Python hesap izi bütçe hücrelerini yılın ilk
kapanış dosyasından okur (Kasım 2026 düzeltmesi; önceden kapanış dosyasından okunuyor ve mutabakat "tutmuyor"
diyordu). Test: `python/tests/test_kalem.py › LineTraceBasisTests`.

### Kalem 3 · Fazla Mesai · Giriş · Ocak–Ekim 2026 (Ekim 2026 kapanışı)

| Adım | Python | Excel (Power Query) | Aynı mı? |
| --- | ---: | ---: | :---: |
| Gider kalemi | Fazla Mesai | Fazla Mesai | ✓ |
| Karşılaştırma bütçesi | Güncel bütçe | Güncel bütçe | ✓ |
| Gerçekleşen (TL) | 17.033.694,15 | 17.033.694,15 | ✓ |
| Seçili bütçe (TL) | 16.915.214,06 | 16.915.214,06 | ✓ |
| Fark = gerçekleşen − bütçe (TL) | 118.480,09 | 118.480,09 | ✓ |
| Fark % | %0,70 | %0,70 | ✓ |
| Yön | Aleyhte | Aleyhte | ✓ |
| Hacim etkisi (TL) | -1.760.939,53 | -1.760.939,53 | ✓ |
| Birim maliyet etkisi (TL) | 1.879.419,62 | 1.879.419,62 | ✓ |
| Gerçekleşen kaynak dosyası | Ekim 2026 - Dosya 1.xlsx | Ekim 2026 - Dosya 1.xlsx | ✓ |

Ay ay kaynak değerleri (Python: kaynak dosyadaki hücreler; Excel: şablonun aylık tablosu):

| Ay | Gerçekleşen · Python | Gerçekleşen · Excel | Bütçe · Python | Bütçe · Excel | Aynı mı? |
| --- | ---: | ---: | ---: | ---: | :---: |
| Ocak | 1.393.783,37 | 1.393.783,37 | 1.708.999,34 | 1.708.999,34 | ✓ |
| Şubat | 1.806.124,57 | 1.806.124,57 | 1.514.317,80 | 1.514.317,80 | ✓ |
| Mart | 1.682.359,31 | 1.682.359,31 | 1.792.749,64 | 1.792.749,64 | ✓ |
| Nisan | 1.947.659,06 | 1.947.659,06 | 1.515.803,13 | 1.515.803,13 | ✓ |
| Mayıs | 1.413.024,69 | 1.413.024,69 | 1.853.578,41 | 1.853.578,41 | ✓ |
| Haziran | 1.657.933,36 | 1.657.933,36 | 1.630.132,79 | 1.630.132,79 | ✓ |
| Temmuz | 1.715.159,30 | 1.715.159,30 | 1.455.902,77 | 1.455.902,77 | ✓ |
| Ağustos | 1.672.721,41 | 1.672.721,41 | 1.913.093,91 | 1.913.093,91 | ✓ |
| Eylül | 2.011.862,53 | 2.011.862,53 | 1.609.733,54 | 1.609.733,54 | ✓ |
| Ekim | 1.733.066,55 | 1.733.066,55 | 1.920.902,73 | 1.920.902,73 | ✓ |

Ayların toplamı gerçekleşen 17.033.694,15 · bütçe 16.915.214,06 TL; özet satırıyla aynı ✓.

Python hesap izi (uygulamadaki Kalem ayrıntısı › Hesap izi tablosu):

| Adım | Tutar | Kaynak dosya · dönem | Nasıl |
| --- | ---: | --- | --- |
| Gerçekleşen | 17.033.694,15 TL | Ekim 2026 - Dosya 1.xlsx (kapanış 2026-10) | 1 bölüm × 1 yıl × ay 1, 2, 3, 4, 5, 6, 7, 8, 9, 10: kaynak hücrelerin toplamı |
| Seçili bütçe · Güncel bütçe | 16.915.214,06 TL | Ekim 2026 - Dosya 1.xlsx (kapanış 2026-10) | Kapanış dosyasındaki bütçe (yıl içindeki son revizyon); aynı aylar (1, 2, 3, 4, 5, 6, 7, 8, 9, 10) toplanır |
| Fark | 118.480,09 TL | — | gerçekleşen − seçili bütçe (pozitif: Aleyhte, negatif: Lehte) |
| Fark % | %0,70 | — | fark / seçili bütçe |

### Kalem 1 · Ücret giderleri · Sayım · Ocak–Ekim 2026 (Ekim 2026 kapanışı)

| Adım | Python | Excel (Power Query) | Aynı mı? |
| --- | ---: | ---: | :---: |
| Gider kalemi | Ücret giderleri | Ücret giderleri | ✓ |
| Karşılaştırma bütçesi | Güncel bütçe | Güncel bütçe | ✓ |
| Gerçekleşen (TL) | 79.071.745,67 | 79.071.745,67 | ✓ |
| Seçili bütçe (TL) | 65.436.487,99 | 65.436.487,99 | ✓ |
| Fark = gerçekleşen − bütçe (TL) | 13.635.257,68 | 13.635.257,68 | ✓ |
| Fark % | %20,84 | %20,84 | ✓ |
| Yön | Aleyhte | Aleyhte | ✓ |
| Hacim etkisi (TL) | 515.826,82 | 515.826,82 | ✓ |
| Birim maliyet etkisi (TL) | 13.119.430,86 | 13.119.430,86 | ✓ |
| Gerçekleşen kaynak dosyası | Ekim 2026 - Dosya 3.xlsx | Ekim 2026 - Dosya 3.xlsx | ✓ |

Ay ay kaynak değerleri (Python: kaynak dosyadaki hücreler; Excel: şablonun aylık tablosu):

| Ay | Gerçekleşen · Python | Gerçekleşen · Excel | Bütçe · Python | Bütçe · Excel | Aynı mı? |
| --- | ---: | ---: | ---: | ---: | :---: |
| Ocak | 6.322.704,46 | 6.322.704,46 | 6.652.974,76 | 6.652.974,76 | ✓ |
| Şubat | 8.215.606,71 | 8.215.606,71 | 5.774.811,72 | 5.774.811,72 | ✓ |
| Mart | 7.836.343,28 | 7.836.343,28 | 6.811.580,29 | 6.811.580,29 | ✓ |
| Nisan | 9.286.483,21 | 9.286.483,21 | 5.860.754,43 | 5.860.754,43 | ✓ |
| Mayıs | 6.749.704,15 | 6.749.704,15 | 7.329.226,68 | 7.329.226,68 | ✓ |
| Haziran | 7.746.490,88 | 7.746.490,88 | 6.484.789,37 | 6.484.789,37 | ✓ |
| Temmuz | 7.812.751,72 | 7.812.751,72 | 5.704.553,28 | 5.704.553,28 | ✓ |
| Ağustos | 7.582.157,90 | 7.582.157,90 | 7.330.441,34 | 7.330.441,34 | ✓ |
| Eylül | 9.298.541,76 | 9.298.541,76 | 6.108.133,68 | 6.108.133,68 | ✓ |
| Ekim | 8.220.961,60 | 8.220.961,60 | 7.379.222,44 | 7.379.222,44 | ✓ |

Ayların toplamı gerçekleşen 79.071.745,67 · bütçe 65.436.487,99 TL; özet satırıyla aynı ✓.

Python hesap izi (uygulamadaki Kalem ayrıntısı › Hesap izi tablosu):

| Adım | Tutar | Kaynak dosya · dönem | Nasıl |
| --- | ---: | --- | --- |
| Gerçekleşen | 79.071.745,67 TL | Ekim 2026 - Dosya 3.xlsx (kapanış 2026-10) | 1 bölüm × 1 yıl × ay 1, 2, 3, 4, 5, 6, 7, 8, 9, 10: kaynak hücrelerin toplamı |
| Seçili bütçe · Güncel bütçe | 65.436.487,99 TL | Ekim 2026 - Dosya 3.xlsx (kapanış 2026-10) | Kapanış dosyasındaki bütçe (yıl içindeki son revizyon); aynı aylar (1, 2, 3, 4, 5, 6, 7, 8, 9, 10) toplanır |
| Fark | 13.635.257,68 TL | — | gerçekleşen − seçili bütçe (pozitif: Aleyhte, negatif: Lehte) |
| Fark % | %20,84 | — | fark / seçili bütçe |
