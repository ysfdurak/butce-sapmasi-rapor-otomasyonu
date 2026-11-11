/* Modül worker'ı. Tarayıcı demosu: Python uygulamasını Pyodide ile bu sekmede çalıştırır.
   Dosyalar hiçbir sunucuya gönderilmez. */
const PYODIDE_URL =
  new URL(self.location.href).searchParams.get("pyodide") ||
  "https://cdn.jsdelivr.net/npm/pyodide@314.0.7/";
const APP_BASE = new URL("../python/", self.location.href).href;

let handle = null;

const ready = (async () => {
  const progress = (text) => postMessage({ type: "progress", text });
  progress("Python motoru yükleniyor…");
  const { loadPyodide } = await import(PYODIDE_URL + "pyodide.mjs");
  const pyodide = await loadPyodide({ indexURL: PYODIDE_URL });

  progress("Uygulama ve örnek veriler hazırlanıyor…");
  const manifest = await (await fetch("manifest.json", { cache: "no-cache" })).json();
  await Promise.all(
    manifest.files.map(async (rel) => {
      const r = await fetch(APP_BASE + rel.split("/").map(encodeURIComponent).join("/"));
      if (!r.ok) throw new Error(rel + " yüklenemedi (" + r.status + ")");
      const target = "/app/" + rel;
      pyodide.FS.mkdirTree(target.slice(0, target.lastIndexOf("/")));
      pyodide.FS.writeFile(target, new Uint8Array(await r.arrayBuffer()));
    }),
  );
  pyodide.FS.mkdirTree("/state");
  pyodide.FS.mkdirTree("/yuklenen");

  progress("Örnek kapanış okunuyor…");
  const adapter = await (await fetch("adapter.py", { cache: "no-cache" })).text();
  await pyodide.runPythonAsync(adapter);
  handle = pyodide.globals.get("handle");
  return pyodide;
})();

ready.then(
  () => postMessage({ type: "ready" }),
  (e) => postMessage({ type: "fatal", text: String(e && e.message ? e.message : e) }),
);

self.onmessage = async ({ data }) => {
  const reply = (msg, transfer) => postMessage({ id: data.id, ...msg }, transfer || []);
  try {
    const pyodide = await ready;
    if (data.type === "upload") {
      const folder = "/yuklenen/" + Date.now();
      const paths = [];
      for (const f of data.files) {
        const target = folder + "/" + f.name;
        pyodide.FS.mkdirTree(target.slice(0, target.lastIndexOf("/")));
        pyodide.FS.writeFile(target, new Uint8Array(f.data));
        paths.push(target);
      }
      reply({ ok: true, folder, paths });
      return;
    }
    const result = handle(data.method, data.path, data.body || new Uint8Array());
    const [status, kind, body, headers] = result.toJs({ dict_converter: Object.fromEntries });
    result.destroy();
    const bytes = new Uint8Array(body);
    reply({ ok: true, status, kind, headers, body: bytes }, [bytes.buffer]);
  } catch (e) {
    reply({ ok: false, error: String(e && e.message ? e.message : e) });
  }
};
