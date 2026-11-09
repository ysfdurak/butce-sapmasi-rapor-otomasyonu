# Bütçe Sapması — Python V4

Excel kaynaklarından bütçe sapması ve operasyonel maliyet analizi. Şirket verileri yerel bilgisayarda kalır. Python 3.9+ gerekir; saf Python bağımlılıkları paket içindedir. PowerShell, Excel COM ve AI servisi kullanılmaz.

**Başlangıç:** Paketi açın, `Baslat.bat` veya `Baslat.pyw` dosyasını çalıştırın. Kaynaklarınızı seçin, kapanış yıl/ayını belirleyin. Mevcut yedi dashboard sayfasını ortak listbox seçimleriyle kullanın; **Çalışma merkezi** üzerinden kaynak sorunlarını, hesap izini, eşleştirme değişikliklerini, kayıtlı görünümleri ve çalışma geçmişini yönetin.

- [V4 kullanım ve yenilikler](V4_YENILIKLER.md)
- [Windows canlı kabul kontrolü](V4_CANLI_KONTROL.md)
- [Ortak kararlar ve açık yeni talepler](ORTAK_KARARLAR.md)
- [Test durumu](Dogrulama_Durumu.json)

Tutarlar gider satırlarından hesaplanır; dört personel ve bir işlem adedi satırı parasal toplama alınmaz. Tam seçili ay kümesi önceki yılda da aynen uygulanır. Bütçe yılı kaynaktan belirlenir. Sentetik örnek 12 bölüm, dört müdürlük ve tek direktörlük içerir.

Yeni: **Önceki ayın yorumlarını getir** ve **Açıklamaları toplu düzenle**. [Kullanım ve 10 yorumluk örnek](V4_YORUM_KULLANIMI.md).
