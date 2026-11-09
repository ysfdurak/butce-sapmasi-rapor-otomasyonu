# Sürüm notları

## V5 — Kasım 2026

Uzman incelemesindeki revizyonlar Python ve Excel VBA sürümlerine eşdeğer uygulandı.

**Yeni rapor sayfaları (iki sürümde de)**
- **Yönetici Özeti** ilk sayfadır ve tek sayfaya yazdırılır (PDF'e hazır): toplamlar, Lehte/Aleyhte yönü, yıl sonu tahmini, veri kontrolü, en büyük aleyhte/lehte önemli sapmalar.
- **Önemli Sapmalar**: tutar **ve** oran eşiği (Ayarlar.xlsx › `Rapor`). Müdürlük dosyalarında açıklanması gereken hücreler turuncu; açıklama durumu formülle izlenir.
- **Yıl Sonu Tahmini**: gerçekleşen + kalan bütçe ve aylık ortalama gidiş yöntemleri.
- **Hacim ve Birim Maliyet**: bütçe farkının hacim etkisi ve birim maliyet etkisi olarak ayrıştırılması.
- Açıklamalarda **onay izi**: son işlem, işlemi yapan kullanıcı ve zaman; silinmeyen işlem günlüğü.

**Hata düzeltmeleri**
- VBA: `Put` adlı prosedür (VBA'nın ayrılmış komutu) `SetCell` olarak değiştirildi; proje derlenmesini engelliyordu.
- VBA: koşullu biçimler dil bağımsız hücre değeri kuralıyla yazılıyor (İngilizce formül metni Türkçe Excel'de hata verebiliyordu).
- VBA: hesap izi ve tablo düzeni ekranları sayfa adını yükleme ile aynı kuralla (büyük/küçük ve Türkçe harf farkı gözetmeden) buluyor; müdürlük adları da aynı kuralla eşleşiyor.
- VBA: rapor hatasında Excel hesaplama/ekran ayarları her durumda geri yükleniyor; ayar dosyasında eksik `dip_toplam_satiri` veya sayı olmayan gider aralığı artık bağlamsız "Type mismatch" vermiyor.
- VBA: geçmiş çalışmayı yeniden üretmek paneldeki seçimleri değiştirmiyor.
- İki sürüm: Ocak kapanışında önceki yılın Aralık dosyasıyla yıl dönümü revizyonları gösteriliyor.
- Python: bütçe işlem adedi okuması sayfa adını Türkçe harf farkı gözetmeden buluyor.

**Sağlamlık ve kullanım**
- Metin olarak girilmiş tutarlar çevriliyor ve hücre adresiyle uyarılıyor; formül sonucu kayıtlı olmayan kaynak hücreleri uyarılıyor; kaynak düzeni kaydığında yeni başlangıç hücresi öneriliyor.
- Python rapor üretimi yaklaşık 12 kat hızlandı (Ekim örneğinde ~120 sn → ~10 sn).
- VBA: kaynaklar ve rapor satırları blok hâlinde okunup yazılıyor; dosya imzaları önbelleğe alınıyor; yalnız üretilen rapor hesaplanıyor.
- `Kurulum.vbs` Türkçe; mevcut xlsm'yi yedekleyip güncelleyebiliyor, V4'ten yükseltebiliyor ve kod sayfası Türkçe değilse uyarıyor.
- `Baslat.bat` Python bulunamazsa kurulum yollarını açıklıyor.
- Girdi arşivi son 20 çalışmayı tutuyor (`arsiv_saklama_adedi`).
- Konumlu düzende eski satır/sütun ayarları zorunlu değil; örnek Ayarlar.xlsx'te bu ayarlar "eski düzen yedeği" olarak işaretli.

**Bakım**
- Python ve web kodu biçimlendirildi (black, prettier); kullanılmayan eski dashboard kodu kaldırıldı.
- `vba_elle/` kopyaları `tools/vba_araclari.py` ile üretiliyor ve denetleniyor; GitHub Actions Python testlerini ve VBA denetimini çalıştırıyor.
- V2–V4 sürüm belgeleri `docs/arsiv/` klasörüne taşındı; ortak belgeler `docs/` altında tek kopya.

## V4

Önceki kapanış yorumlarını inceleme kuyruğuna alma, toplu düzenleme ve onay, kaynak kapanışın görünmesi ve mevcut dönem yorumlarının korunması eklendi. Ayrıntılar: [arşiv](arsiv/).

Gider verisinden faturaya inen ayrıntı çıkarımı, otomatik neden metni üretimi ve yerel dil modeli deneyi bu paketin çalışma akışında yer almaz ([açık talepler](ORTAK_KARARLAR.md)).
