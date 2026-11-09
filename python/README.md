# Bütçe Sapması — Python V5

Excel kaynaklarından bütçe sapması ve operasyonel maliyet analizi. Şirket verileri yerel bilgisayarda kalır. Python 3.9+ gerekir; saf Python bağımlılıkları paket içindedir. PowerShell, Excel COM ve AI servisi kullanılmaz.

**Başlangıç:** Paketi açın, `Baslat.bat` veya `Baslat.pyw` dosyasını çalıştırın. Kaynaklarınızı seçin, kapanış yıl/ayını belirleyin. Mevcut yedi dashboard sayfasını ortak listbox seçimleriyle kullanın; **Çalışma merkezi** üzerinden kaynak sorunlarını, hesap izini, eşleştirme değişikliklerini, kayıtlı görünümleri ve çalışma geçmişini yönetin.

- [Kullanım](KULLANIM.md) · [Veri ve hesap](../docs/VERI_VE_HESAP.md) · [Teknik yapı](TEKNIK_YAPI.md)
- [Windows canlı kabul kontrolü](../docs/CANLI_KONTROL.md)
- [Ortak kararlar ve açık talepler](../docs/ORTAK_KARARLAR.md)
- [Test durumu](Dogrulama_Durumu.json)

Tutarlar gider satırlarından hesaplanır; dört personel ve bir işlem adedi satırı parasal toplama alınmaz. Tam seçili ay kümesi önceki yılda da aynen uygulanır. Bütçe yılı kaynaktan belirlenir. Sentetik örnek 12 bölüm, dört müdürlük ve tek direktörlük içerir.

V5 raporlarında: **Yönetici Özeti** (ilk sayfa), **Önemli Sapmalar** (eşik + açıklama zorunluluğu), **Yıl Sonu Tahmini**, **Hacim ve Birim Maliyet**; açıklamalarda onay izi. [Açıklamalar ve 10 yorumluk örnek](../docs/YORUM_KULLANIMI.md).
