# Tarayıcı demosu

Adres: https://ysfdurak.github.io/butce-sapmasi-rapor-otomasyonu/demo/

Demo, Python sürümünü **kurulum ve sunucu olmadan** ziyaretçinin tarayıcısında çalıştırır. Hesap kodu (`model.py`, `butce.py`, `analiz.py` …) ve arayüz (`python/web/`) masaüstü sürümüyle **aynı dosyalardır**; demo için kopyalanmaz. Masaüstünde yapılan her düzeltme demoya da yansır.

## Nasıl çalışıyor?

```
Tarayıcı sekmesi                          Web Worker (arka plan)
───────────────                           ──────────────────────
python/web/app.js  ── fetch("/api/…") ─▶  demo/bridge.js isteği yakalar
                                           │ postMessage
                                           ▼
                                          demo/worker.js
                                           • Pyodide (WebAssembly Python 3.14) yüklenir
                                           • manifest.json'daki .py, wheel ve örnek .xlsx
                                             dosyaları sanal diske (/app) yazılır
                                           • demo/adapter.py → app.Handler yönlendirmesi
```

| Dosya | Görevi |
|---|---|
| `demo/index.html` | `python/web/index.html` kabuğunun demo sürümü: yükleme ekranı ve alt bilgi bandı |
| `demo/bridge.js` | `app.js`'in `/api/…` ve `/download/…` isteklerini worker'a yönlendirir; dosya seçimini tarayıcı seçicisine, indirmeleri Blob'a çevirir |
| `demo/worker.js` | Pyodide'i yükler, dosyaları sanal diske yazar, istekleri Python'a iletir |
| `demo/adapter.py` | Masaüstündeki `app.Handler` sınıfını HTTP soketi olmadan çağırır; uzun işleri sırayla çalıştırır |
| `demo/manifest.json` | Yüklenecek dosya listesi; `python tools/demo_manifest.py` ile üretilir |

## Masaüstünden farkları

- **İş parçacığı yok:** Pyodide thread açamaz. Rapor üretimi istek içinde çalışır; ilerleme adımları yerine "Hesaplanıyor" ekranı görünür.
- **Dosya seçimi:** Windows penceresi yerine tarayıcı dosya seçicisi. Seçilen dosyalar sanal diske kopyalanır, hiçbir sunucuya gönderilmez.
- **Çıktı klasörü sanaldır:** Raporlar indirme bağlantısıyla verilir.
- **Kalıcılık yok:** Açıklamalar ve kayıtlı görünümler sekme kapanınca silinir.
- **Kapat düğmesi gizlidir.**

## Bakım

- Örnek veri veya yeni bir `.py` modülü eklenince `python tools/demo_manifest.py` çalıştırın. GitHub Actions listenin güncel olduğunu denetler.
- Pyodide sürümü `demo/worker.js` içindeki `PYODIDE_URL` satırındadır. Yerel test için: `python -m http.server` ile repo kökünü açıp `demo/?pyodide=<yerel-pyodide-adresi>/` kullanılabilir.
- Yayın: GitHub Pages, `main` dalının kök klasöründen.
