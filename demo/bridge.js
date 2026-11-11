/* Masaüstü arayüzünün (python/web/app.js) /api isteklerini tarayıcıdaki Python'a yönlendirir. */
(function () {
  const params = new URL(location.href).searchParams;
  const worker = new Worker("worker.js" + (params.get("pyodide") ? "?pyodide=" + encodeURIComponent(params.get("pyodide")) : ""), {
    type: "module",
  });
  const pending = new Map();
  let seq = 0;
  let lastPick = null;

  const loader = () => document.getElementById("demoLoader");
  const busy = (on, text) => {
    const el = document.getElementById("demoBusy");
    if (!el) return;
    el.hidden = !on;
    if (text) el.querySelector("span").textContent = text;
  };

  worker.onerror = (e) => {
    const el = loader();
    if (el) {
      el.querySelector("p").textContent = "Demo başlatılamadı: " + (e.message || "tarayıcı desteklemiyor");
      el.classList.add("error");
    }
  };

  worker.onmessage = ({ data }) => {
    if (data.type === "progress") {
      const el = loader();
      if (el) el.querySelector("p").textContent = data.text;
      return;
    }
    if (data.type === "ready") {
      const el = loader();
      if (el) el.hidden = true;
      return;
    }
    if (data.type === "fatal") {
      const el = loader();
      if (el) {
        el.querySelector("p").textContent = "Demo başlatılamadı: " + data.text;
        el.classList.add("error");
      }
      return;
    }
    const p = pending.get(data.id);
    if (!p) return;
    pending.delete(data.id);
    data.ok ? p.resolve(data) : p.reject(new Error(data.error));
  };

  const call = (msg, transfer) =>
    new Promise((resolve, reject) => {
      const id = ++seq;
      pending.set(id, { resolve, reject });
      worker.postMessage({ id, ...msg }, transfer || []);
    });

  const json = (status, value) =>
    new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });

  // Dosya seçimi: masaüstündeki Windows penceresi yerine tarayıcı dosya seçicisi.
  function chooseFiles(mode) {
    return new Promise((resolve) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".xlsx";
      if (mode === "files") input.multiple = true;
      if (mode === "folder") input.webkitdirectory = true;
      input.style.display = "none";
      document.body.appendChild(input);
      input.addEventListener("change", async () => {
        const files = await Promise.all(
          [...input.files].map(async (f) => ({
            name: f.webkitRelativePath || f.name,
            data: await f.arrayBuffer(),
          })),
        );
        input.remove();
        resolve(files);
      });
      input.addEventListener("cancel", () => {
        input.remove();
        resolve([]);
      });
      input.click();
    });
  }

  async function pick(mode) {
    // Çıktı klasörleri sanaldır; raporlar indirme bağlantısıyla verilir.
    if (mode === "folder" && lastPick && /output/i.test(lastPick)) return { value: "/app/ciktilar" };
    const files = await chooseFiles(mode);
    if (!files.length) return { value: mode === "files" ? [] : "" };
    const r = await call({ type: "upload", files }, files.map((f) => f.data));
    if (mode === "folder") {
      const top = files[0].name.includes("/") ? "/" + files[0].name.split("/")[0] : "";
      return { value: r.folder + top };
    }
    return { value: mode === "files" ? r.paths : r.paths[0] };
  }

  const nativeFetch = window.fetch.bind(window);
  window.fetch = async (input, init = {}) => {
    const url = new URL(typeof input === "string" ? input : input.url, location.href);
    const path = url.pathname.replace(/^.*?(\/api\/|\/download\/)/, "$1");
    if (!/^\/(api|download)\//.test(path)) return nativeFetch(input, init);
    const method = (init.method || "GET").toUpperCase();
    let args = {};
    try {
      args = init.body ? JSON.parse(init.body) : {};
    } catch {}

    if (path === "/api/pick") return json(200, await pick(args.mode));
    if (path === "/api/shutdown")
      return json(400, { error: "Tarayıcı demosunda kapatma yok; sekmeyi kapatmanız yeterli." });

    const long = path === "/api/job" && method === "POST";
    if (long) busy(true, "Tarayıcıda hesaplanıyor… Rapor üretimi 15–40 saniye sürebilir.");
    try {
      const body = init.body ? new TextEncoder().encode(init.body) : new Uint8Array();
      const r = await call({ type: "request", method, path: path + url.search, body }, [body.buffer]);
      return new Response(r.body, {
        status: r.status === 204 ? 200 : r.status,
        headers: { "Content-Type": r.kind, ...(r.headers || {}) },
      });
    } finally {
      if (long) busy(false);
    }
  };

  document.addEventListener(
    "click",
    async (e) => {
      const pickBtn = e.target.closest("[data-pick]");
      if (pickBtn) lastPick = pickBtn.dataset.pick;
      const link = e.target.closest('a[href^="/download/"]');
      if (!link) return;
      e.preventDefault();
      const r = await fetch(link.getAttribute("href"));
      if (!r.ok) return;
      const blob = await r.blob();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = link.textContent.replace(/\s*↓\s*$/, "").trim();
      document.body.appendChild(a);
      a.click();
      setTimeout(() => (URL.revokeObjectURL(a.href), a.remove()), 1000);
    },
    true,
  );
})();
