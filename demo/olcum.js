// Anonim kullanım ölçümü ve geri bildirim düğmesi (ortak servis: olcum.ysfdurak.com; kaynak: ysfdurak.com → sunucu/olcum).
// Yalnız GitHub Pages'teki resmî demoda yüklenir; masaüstü sürümü ve kendi kopyalar dışarıya istek atmaz.
// Gönderilenler: açılan sayfa adı (ör. "priority"), "rapor üretildi" işlemi, aktif süre, cihaz türü, dil.
// Dosya, tutar, bölüm adı gönderilmez; çerez yok; DNT/GPC'ye uyar.
(function () {
  if (location.hostname !== "ysfdurak.github.io") return;
  window.olcumButce = true;
  const kuyruk = (window.olcumKuyruk ||= []);
  kuyruk.push(["ekran", "general"]);
  document.addEventListener(
    "click",
    (e) => {
      const b = e.target.closest && e.target.closest("[data-page]");
      if (b && b.dataset.page) (window.olcumKuyruk ||= []).push(["ekran", b.dataset.page]);
    },
    true,
  );
  const s = document.createElement("script");
  s.src = "https://olcum.ysfdurak.com/olcum.js";
  s.defer = true;
  s.dataset.proje = "butce";
  document.head.appendChild(s);
})();
