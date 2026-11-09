# İlk çalıştırma

Önce sentetik veriyle deneyin. GitHub'dan ZIP olarak aldıysanız tüm arşivi bir klasöre çıkarın. Alt klasörlerin düzenini koruyun.

## Excel VBA

Windows masaüstü Excel gerekir. `makro/` altında `Kontrol_Sablonu.xlsx`, `Kurulum.vbs`, `vba_elle/` ve `ornek/` yan yana durur.

1. Kurum ayarları izin veriyorsa `Kurulum.vbs` çalıştırın. Bu adım aynı klasöre `Butce_Kontrol_Makro_V4.xlsm` oluşturur.
2. Elle kurulum için `Kontrol_Sablonu.xlsx` dosyasını Excel'de açın ve aynı klasöre **Excel Makro İçerebilen Çalışma Kitabı (*.xlsm)** olarak kaydedin. **Alt+F11** tuşlarına basın, ilgili çalışma kitabının projesini seçin ve **File → Import File** ile `vba_elle/` içindeki **10** `.bas` dosyasını ekleyin. Bu kopyalar Türkçe Windows kodlamasıyla hazırlanmıştır.
3. **Debug → Compile VBAProject** ile derleyin ve kaydedin. Excel'de **Alt+F8 → PanelHazirla → Çalıştır** yoluyla paneli açın.
4. **Sentetik hesap testi** düğmesine basın ve **Test Sonuçları** sayfasına bakın.
5. Ardından `makro/ornek/Ayarlar.xlsx` dosyasını seçin. Kapanış **2026 / 10**, rapor yılı **2026** ve aylar **Ocak–Ekim** olsun. Bir çıktı klasörü seçip **Kaynakları önizle → Rapor oluştur** adımlarını uygulayın.

Kurulum betiği için Excel'in VBA proje erişimine, elle kurulum için VBA düzenleyicisine ve çalışma kitabı makrolarını çalıştırma iznine ihtiyaç vardır. Kurumsal kısıtlarla karşılaştığınızda şirketinizin BT prosedürünü izleyin.

[Excel VBA ayrıntılı kullanım](../makro/KULLANIM.md) · [Makro yorum senaryosu](../makro/V4_YORUM_KULLANIMI.md)

## Python

1. Python 3.9 veya üzeri bulunan Windows bilgisayarda `python/Baslat.bat` dosyasını çalıştırın. Dosya ilişkilendirmesi varsa `python/Baslat.pyw` da kullanılabilir. Komut satırından başlatmak için `python/bootstrap.py` dosyasını Python ile çalıştırın.
2. Tarayıcıda `127.0.0.1` adresinde sentetik dashboard açılır. Kapanış **Ekim 2026**, yıl **2026**, aylar **Ocak–Ekim** ve tüm bölümler seçilidir.
3. Sayfalar arasında geçip ay ve bölüm listelerini değiştirin. **Rapor oluştur** ile bir çıktı klasörü seçin ve raporları üretin.
4. İşiniz bitince sağ üstteki güç düğmesiyle yerel işlemi durdurun.

Paketin `wheels/` klasöründe openpyxl ve et_xmlfile bulunur. Başlatıcı bunları doğrudan kullanır. Masaüstü Excel, pip ve PowerShell başlangıç için gerekmez. Formüllü gerçek Excel kaynaklarıyla çalışırken kaynak formüllerinin sonuçlarını önceden Excel'de hesaplayıp kaydedin.

[Python ayrıntılı kullanım](../python/KULLANIM.md) · [Python yorum senaryosu](../python/V4_YORUM_KULLANIMI.md)

## Beklenen demo değerleri

| Ölçüm | Ocak–Ekim 2026 | Ocak, Mart, Mayıs 2026 |
| --- | ---: | ---: |
| Gerçekleşen gider | 3.449.383.390,31 TL | 913.186.126,67 TL |
| Bütçe gideri | 3.440.938.980,13 TL | 1.094.635.989,61 TL |
| Bütçe − gerçekleşen | −8.444.410,18 TL | 181.449.862,94 TL |

Eylül ve Ekim dosyalarıyla 26 geçmiş gider/bütçe revizyonu beklenir. Önceki kapanış geçişleri için örnek dosyalar bulunmadığından ilgili eksik kaynak uyarıları beklenir. [Referans çalışma kitabı](../makro/Sentetik_Beklenen_Sonuclar.xlsx)
