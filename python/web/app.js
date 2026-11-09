"use strict";
const $ = (s) => document.querySelector(s),
  $$ = (s) => [...document.querySelectorAll(s)];
const months = [
  "Ocak",
  "Şubat",
  "Mart",
  "Nisan",
  "Mayıs",
  "Haziran",
  "Temmuz",
  "Ağustos",
  "Eylül",
  "Ekim",
  "Kasım",
  "Aralık",
];
const pages = [
  ["general", "Genel durum"],
  ["staff", "Personel"],
  ["transactions", "İşlem adedi"],
  ["detail", "Detay durum"],
  ["departments", "Bölüm analizi"],
  ["unit", "Birim maliyet"],
  ["priority", "Sapma öncelikleri"],
];
const filterNames = {
  manager: "Müdürlük",
  submanager: "Müdürlük altı",
  location: "Lokasyon",
  department: "Bölüm",
  year: "Yıl",
  month: "Ay",
};
const palette = ["#4b7daf", "#8d6cb0", "#398c82", "#cb8640", "#b95872", "#6b8d45", "#586db5", "#468eaa"];
let meta,
  data,
  page = "general",
  state = { filters: {}, bmLocal: null, opexLocal: null },
  staffMode = "monthly",
  departmentMode = "expense";
let groupsOpen = new Set(),
  yearsClosed = new Set(),
  querySeq = 0,
  timer,
  queryTimer,
  trendBounds = null,
  prioritySearch = "";
let sourcePreview = null,
  workspaceData = null,
  layoutData = null;
let commentReview = null,
  commentSelected = new Set(),
  commentDirty = new Set(),
  commentFilter = "pending",
  commentSearch = "",
  commentOpen = false;
const fmt = (v, d = 2) =>
  v === null || v === undefined || !Number.isFinite(Number(v))
    ? "—"
    : Number(v).toLocaleString("tr-TR", { minimumFractionDigits: d, maximumFractionDigits: d });
const pct = (v) => (v === null || v === undefined ? "—" : fmt(v * 100) + "%");
const esc = (s) =>
  String(s ?? "").replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c],
  );
const signed = (v, positiveGood = true) =>
  v === null || v === undefined || v === 0 ? "" : v > 0 === positiveGood ? "good" : "bad";
function toast(text) {
  $("#toast").textContent = text;
  $("#toast").hidden = false;
  clearTimeout(timer);
  timer = setTimeout(() => ($("#toast").hidden = true), 6500);
}
async function api(path, args) {
  const r = await fetch(
    "/api/" + path,
    args === undefined
      ? {}
      : { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(args) },
  );
  let value;
  try {
    value = await r.json();
  } catch {
    throw Error("Uygulama yanıt vermedi. Baslat.bat ile yeniden açın.");
  }
  if (!r.ok) throw Error(value.error || "İşlem tamamlanamadı.");
  return value;
}
function defaults() {
  state = {
    filters: Object.fromEntries(
      Object.keys(filterNames).map((k) => [
        k,
        k === "year" ? [meta.budgetYear] : k === "month" ? meta.defaults.month.slice() : null,
      ]),
    ),
    bmLocal: null,
    opexLocal: null,
  };
}
function selected(key) {
  return state.filters[key] ?? meta.options[key];
}
function label(key, v) {
  return key === "month" ? `${String(v).padStart(2, "0")} ${months[v - 1]}` : v;
}
function buildFilters() {
  const scrolls = $$("#filters .listbox").map((el) => el.scrollTop);
  $("#filters").innerHTML = Object.keys(filterNames)
    .map(
      (key) =>
        `<div class="filter"><div class="filter-heading"><span>${filterNames[key]}</span><button data-filter-all="${key}">${selected(key).length === meta.options[key].length ? "Temizle" : "Tümü"}</button></div><div class="listbox" role="listbox" aria-label="${filterNames[key]}" aria-multiselectable="true">${meta.options[key].map((v, i) => `<label role="option" aria-selected="${selected(key).includes(v)}"><input type="checkbox" data-filter="${key}" data-index="${i}" ${selected(key).includes(v) ? "checked" : ""}>${esc(label(key, v))}</label>`).join("")}</div></div>`,
    )
    .join("");
  $$("#filters .listbox").forEach((el, i) => (el.scrollTop = scrolls[i] || 0));
}
function renderNav() {
  $("#nav").innerHTML = pages
    .map(
      ([id, label], i) =>
        `<button data-page="${id}" class="${page === id ? "active" : ""}" aria-current="${page === id ? "page" : "false"}"><span class="nav-num">0${i + 1}</span>${label}</button>`,
    )
    .join("");
}
function schedule() {
  clearTimeout(queryTimer);
  queryTimer = setTimeout(refresh, 100);
}
async function refresh() {
  const seq = ++querySeq;
  document.body.classList.add("loading");
  $("#status").textContent = "Seçim hesaplanıyor…";
  try {
    const result = await api("query", state);
    if (seq !== querySeq) return;
    data = result;
    trendBounds = null;
    render();
    $("#status").textContent = "Güncel";
  } catch (e) {
    if (seq === querySeq) {
      toast(e.message);
      $("#status").textContent = "Hesaplanamadı";
    }
  } finally {
    if (seq === querySeq) document.body.classList.remove("loading");
  }
}
function render() {
  renderNav();
  $("#pageTitle").textContent = pages.find((p) => p[0] === page)[1];
  $("#subtitle").textContent =
    `${months[meta.closing[1] - 1]} ${meta.closing[0]} kapanışı · ${meta.options.department.length} bölüm · ${meta.options.manager.length} müdürlük`;
  const sel = data.selection;
  $("#selectionCaption").textContent =
    `${sel.departments.length} bölüm · ${sel.years.join(", ") || "Yıl seçilmedi"} · ${sel.months.map((m) => months[m - 1]).join(", ") || "Ay seçilmedi"}`;
  $("#unitCaption").textContent =
    page === "staff" || page === "transactions"
      ? "Birim: adet · binlik ayıraç"
      : page === "unit"
        ? "Birim: TL / işlem"
        : "Tüm tutarlar: Milyon TL";
  $("#messages").innerHTML = data.messages.map((m) => `<div class="notice">${esc(m)}</div>`).join("");
  $("#sourceStamp").textContent =
    `Kaynak yenileme: ${meta.loadedAt.replace("T", " ")} · Bütçe yılı ${meta.budgetYear}`;
  $("#content").innerHTML = {
    general: renderGeneral,
    staff: renderStaff,
    transactions: renderTransactions,
    detail: renderDetail,
    departments: renderDepartments,
    unit: renderUnit,
    priority: renderPriority,
  }[page]();
  $("#tooltip").hidden = true;
}
function card(label, value, unit, color = "") {
  return `<div class="card"><div class="card-label">${esc(label)}</div><div class="card-value ${color}">${value}</div><div class="card-unit">${unit}</div></div>`;
}
function cards(c, counts = false, title = "İşlem adedi") {
  const d = counts ? 0 : 2;
  return `<div class="cards">${card(counts ? "Bütçe · " + title : "Bütçe tutarı", fmt(c.budget, d), c.unit)}${card(counts ? "Gerçekleşen · " + title : "Gerçekleşen tutar", fmt(c.actual, d), c.unit)}${card("Bütçe − gerçekleşen", fmt(c.gap, d), c.unit, signed(c.gap, !counts))}${card("Bütçe sapma oranı", pct(c.rate), "Sapma / bütçe", signed(c.rate, !counts))}</div>`;
}
function panel(title, content, sub = "", extra = "") {
  return `<section class="panel"><div class="panel-head"><div><h3>${esc(title)}</h3>${sub ? `<div class="panel-subtitle">${esc(sub)}</div>` : ""}</div>${extra}</div>${content}</section>`;
}
function empty(msg = "Seçime uygun veri yok.") {
  return `<div class="empty">${esc(msg)}</div>`;
}
function legend(items) {
  return `<div class="legend">${items.map(([name, color]) => `<span><i style="background:${color}"></i>${esc(name)}</span>`).join("")}</div>`;
}
function bars(
  rows,
  { field = "actual", money = true, signedColor = false, width = 510, labelWidth = 133 } = {},
) {
  if (!rows.some((r) => r[field] !== null && r[field] !== undefined)) return empty();
  const h = Math.max(150, rows.length * 46 + 20),
    left = labelWidth,
    right = 92,
    range = width - left - right;
  const vals = rows.map((r) => r[field]).filter((v) => v !== null && v !== undefined),
    lo = Math.min(0, ...vals),
    hi = Math.max(0, ...vals);
  const span = hi - lo || 1;
  const x = (v) => left + ((v - lo) / span) * range;
  const zero = x(0);
  return `<div class="chart"><svg viewBox="0 0 ${width} ${h}" role="img" aria-label="Yatay çubuk grafik"><line x1="${zero}" x2="${zero}" y1="3" y2="${h - 12}" class="axis"/>${rows
    .map((r, i) => {
      const v = r[field],
        y = 20 + i * 46;
      return `<text x="${left - 10}" y="${y + 12}" text-anchor="end">${esc(r.name)}</text>${v === null || v === undefined ? `<text x="${left + 6}" y="${y + 12}">—</text>` : `<rect x="${Math.min(x(v), zero)}" y="${y}" width="${Math.max(1, Math.abs(x(v) - zero))}" height="19" rx="3" fill="${signedColor ? (v < 0 ? "#c56b75" : "#5a9d85") : "#5c869e"}"/><text class="value-label" x="${Math.max(x(v), zero) + 7}" y="${y + 14}" text-anchor="start">${fmt(v, money ? 2 : 0)}</text>`}`;
    })
    .join("")}</svg></div>`;
}
function lineChart(
  series,
  categories,
  { width = 720, height = 330, counts = false, bounds = null, labels = true } = {},
) {
  const values = series.flatMap((s) => s.values.filter((v) => v !== null));
  if (!values.length) return empty();
  const min = Math.min(...values),
    max = Math.max(...values),
    pad = (max - min) * 0.17 || Math.max(Math.abs(max) * 0.1, 1);
  const lo = bounds ? bounds[0] : min < 0 ? min - pad : Math.max(0, min - pad),
    hi = bounds ? bounds[1] : max + pad;
  const left = 65,
    right = 35,
    top = 33,
    bottom = 46,
    iw = width - left - right,
    ih = height - top - bottom;
  const xx = (i) => left + (categories.length > 1 ? (i / (categories.length - 1)) * iw : iw / 2),
    yy = (v) => top + ((hi - v) / (hi - lo || 1)) * ih;
  let body = "";
  for (let i = 0; i <= 4; i++) {
    const v = lo + ((hi - lo) * i) / 4,
      y = yy(v);
    body += `<line x1="${left}" x2="${width - right}" y1="${y}" y2="${y}" class="grid"/><text x="${left - 9}" y="${y + 4}" text-anchor="end">${fmt(v, counts ? 0 : 2)}</text>`;
  }
  body += categories
    .map((c, i) => `<text x="${xx(i)}" y="${height - 17}" text-anchor="middle">${esc(c)}</text>`)
    .join("");
  for (let si = 0; si < series.length; si++) {
    const s = series[si];
    let segment = [];
    const flush = () => {
      if (segment.length)
        body += `<polyline points="${segment.join(" ")}" fill="none" stroke="${s.color}" stroke-width="2.5" stroke-linejoin="round"/>`;
      segment = [];
    };
    s.values.forEach((v, i) => {
      if (v === null) {
        flush();
        return;
      }
      segment.push(`${xx(i)},${yy(v)}`);
    });
    flush();
    s.values.forEach((v, i) => {
      if (v === null || v < lo || v > hi) return;
      body += `<circle cx="${xx(i)}" cy="${yy(v)}" r="3.8" fill="${s.color}"><title>${esc(s.name)} · ${esc(categories[i])}: ${fmt(v, counts ? 0 : 2)}</title></circle>`;
      if (labels)
        body += `<text x="${xx(i)}" y="${yy(v) + (si % 2 === 0 ? -11 : 17)}" text-anchor="middle" class="value-label">${fmt(v, counts ? 0 : 2)}</text>`;
    });
  }
  return `<div class="chart wide"><svg width="${width}" viewBox="0 0 ${width} ${height}" role="img" aria-label="Dönem çizgi grafiği">${body}</svg></div>`;
}
function renderGeneral() {
  const years = data.selection.years,
    selected = data.selection.months,
    rows = data.general.trend;
  const series = years.map((y, i) => ({
    name: `${y} Gerçekleşen`,
    color: palette[i % palette.length],
    values: selected.map((m) => rows.find((r) => r.year === y && r.month === m)?.actual ?? null),
  }));
  if (years.length === 1 && years[0] === meta.budgetYear)
    series.push({
      name: `${meta.budgetYear} Bütçe`,
      color: "#c28a50",
      values: selected.map((m) => rows.find((r) => r.month === m)?.budget ?? null),
    });
  const chart = lineChart(
    series,
    selected.map((m) => months[m - 1].slice(0, 3)),
    { width: Math.max(600, selected.length * 64), height: 350, bounds: trendBounds },
  );
  return (
    cards(data.general.cards) +
    `<div class="general-grid">${panel("Müdürlük bazında gerçekleşen gider", bars(data.general.managers), "Seçili yıl ve ayların toplamı")}${panel("Müdürlük bazında bütçe sapması", bars(data.general.managers, { field: "gap", signedColor: true }), "Bütçe − gerçekleşen · Milyon TL")}<section class="panel trend-panel"><div class="panel-head"><div><h3>Aylık bütçe ve gerçekleşen</h3><div class="panel-subtitle">Her ayın kendi değeri · Milyon TL</div></div></div>${legend(series.map((s) => [s.name, s.color]))}${chart}<div class="zoom-controls"><span>Dikey aralık</span><input id="ymin" type="number" aria-label="Grafik alt sınırı" placeholder="Alt" value="${trendBounds?.[0] ?? ""}"><input id="ymax" type="number" aria-label="Grafik üst sınırı" placeholder="Üst" value="${trendBounds?.[1] ?? ""}"><button data-zoom="apply">Uygula</button><button data-zoom="in" aria-label="Yakınlaştır">+</button><button data-zoom="out" aria-label="Uzaklaştır">−</button><button data-zoom="up" aria-label="Aralığı yukarı kaydır">↑</button><button data-zoom="down" aria-label="Aralığı aşağı kaydır">↓</button><button data-zoom="reset">Sıfırla</button></div></section></div>`
  );
}
function combo(rows, totals = false) {
  const years = data.selection.years,
    ms = totals ? [null] : data.selection.months;
  const categories = totals ? ["Seçili aylar"] : ms.map((m) => months[m - 1].slice(0, 3));
  const values = rows.flatMap((r) => [r.actual, r.budget]).filter((v) => v !== null);
  if (!values.length) return empty();
  const width = Math.max(470, ms.length * Math.max(years.length, 1) * 73 + 85),
    height = 290,
    left = 56,
    top = 32,
    bottom = 43;
  const hi = Math.max(...values) * 1.2 || 1,
    iw = width - left - 25,
    ih = height - top - bottom,
    slot = iw / Math.max(ms.length, 1),
    barW = Math.min(38, (slot * 0.65) / Math.max(years.length, 1));
  const x = (i) => left + slot * (i + 0.5),
    yy = (v) => top + ih - (v / hi) * ih;
  let out = "";
  for (let i = 0; i <= 4; i++) {
    const v = (hi * i) / 4;
    out += `<line class="grid" x1="${left}" x2="${width - 20}" y1="${yy(v)}" y2="${yy(v)}"/><text x="${left - 7}" y="${yy(v) + 4}" text-anchor="end">${fmt(v, 0)}</text>`;
  }
  categories.forEach(
    (cat, i) => (out += `<text x="${x(i)}" y="${height - 14}" text-anchor="middle">${esc(cat)}</text>`),
  );
  for (let yi = 0; yi < years.length; yi++)
    for (let mi = 0; mi < ms.length; mi++) {
      const r = rows.find((r) => r.year === years[yi] && (totals || r.month === ms[mi]));
      if (r?.actual === null || !r) continue;
      const bx = x(mi) + (yi - (years.length - 1) / 2) * barW - barW * 0.42;
      out += `<rect x="${bx}" y="${yy(r.actual)}" width="${barW * 0.84}" height="${top + ih - yy(r.actual)}" rx="2" fill="${palette[yi % palette.length]}"><title>${years[yi]} ${esc(categories[mi])}: ${fmt(r.actual, 0)}</title></rect><text class="value-label" text-anchor="middle" x="${bx + barW * 0.42}" y="${yy(r.actual) - 7}">${fmt(r.actual, 0)}</text>`;
    }
  const b = ms.map(
    (m) => rows.find((r) => r.year === meta.budgetYear && (totals || r.month === m))?.budget ?? null,
  );
  if (b.some((v) => v !== null)) {
    out += `<polyline fill="none" stroke="#bd743b" stroke-width="2.5" points="${b.map((v, i) => (v === null ? "" : `${x(i)},${yy(v)}`)).join(" ")}"/>`;
    b.forEach((v, i) => {
      if (v !== null)
        out += `<circle cx="${x(i)}" cy="${yy(v)}" r="4" fill="#bd743b"/><text x="${x(i)}" y="${yy(v) + 18}" text-anchor="middle" class="value-label">${fmt(v, 0)}</text>`;
    });
  }
  return (
    legend([
      ...years.map((y, i) => [`${y} Gerçekleşen`, palette[i % palette.length]]),
      ...(b.some((v) => v !== null) ? [[`${meta.budgetYear} Bütçe`, "#bd743b"]] : []),
    ]) +
    `<div class="chart wide"><svg width="${width}" viewBox="0 0 ${width} ${height}" role="img" aria-label="Personel sütun ve bütçe çizgi grafiği">${out}</svg></div>`
  );
}
function segments(id, entries, current) {
  return `<div class="segments">${entries.map(([value, text]) => `<button data-mode="${id}" data-value="${value}" class="${current === value ? "active" : ""}">${text}</button>`).join("")}</div>`;
}
function renderStaff() {
  return `<div class="section-tools">${segments(
    "staff",
    [
      ["monthly", "Yıl / ay bazında"],
      ["totals", "Seçili aylar toplamı"],
    ],
    staffMode,
  )}<span class="control-note">Başlıklar kaynak Excel’den okunur. Çok yıl seçildiğinde grafik yatay genişler.</span></div><div class="chart-grid">${data.staff.map((s) => panel(s.name, combo(s[staffMode], staffMode === "totals"), s.titleVariants?.length > 1 ? s.titleVariants.map((v) => `${v.name}: ${v.departments.join(", ")}`).join(" · ") : "")).join("")}</div>`;
}
function pair(c) {
  return `${fmt(c.value, 0)} (${pct(c.rate)})`;
}
function renderTransactions() {
  const columns = [...data.selection.months.map((m) => months[m - 1]), "Seçili aylar toplamı"];
  let rows = "";
  for (const block of data.transactions.rows) {
    const closed = yearsClosed.has(block.year);
    rows += `<tr class="year-row"><td><button class="collapse" data-year="${block.year}" aria-expanded="${!closed}">${closed ? "＋" : "−"}</button>${block.year}</td>${block.values.map((v) => `<td>${fmt(v.actual, 0)}</td>`).join("")}</tr>`;
    if (!closed) {
      const metrics = [
        ["Gerçekleşen", (v) => fmt(v.actual, 0), ""],
        ...(block.year === meta.budgetYear ? [["Bütçe", (v) => fmt(v.budget, 0), ""]] : []),
        ["Önceki yıla göre değişim", (v) => pair(v.annual), ""],
        ...(block.year === meta.budgetYear
          ? [["Bütçeye göre sapma", (v) => `${fmt(v.gap.value, 0)} (${pct(v.gapRate)})`, "gap"]]
          : []),
      ];
      for (const [name, formatter, kind] of metrics)
        rows += `<tr class="detail"><td>${name}</td>${block.values.map((v) => `<td class="${kind === "gap" ? signed(v.gap.value, false) : ""}">${formatter(v)}</td>`).join("")}</tr>`;
    }
  }
  return (
    cards(data.transactions.cards, true, data.transactions.name) +
    panel(
      "Yıl ve ay bazında · " + data.transactions.name,
      `<div class="scroll-table"><table><thead><tr><th>Yıl / metrik</th>${columns.map((c) => `<th>${esc(c)}</th>`).join("")}</tr></thead><tbody>${rows || `<tr><td colspan="14">Seçime uygun kayıt yok.</td></tr>`}</tbody></table></div>`,
      "Yılların yanındaki + / − ile metrikleri açın. Sapma = bütçe − gerçekleşen; bütçeyi aşan adet yeşil.",
    )
  );
}
function hoverAttr(history, name) {
  return `tabindex="0" data-hover="${esc(JSON.stringify({ history, name }))}"`;
}
function detailRow(row, kind) {
  return `<tr class="${kind}"><td>${kind === "group" ? `<button class="collapse" data-group="${esc(row.name)}" aria-expanded="${groupsOpen.has(row.name)}">${groupsOpen.has(row.name) ? "−" : "＋"}</button>` : ""}${esc(row.name)}</td><td ${hoverAttr(row.history || [], row.name)}>${fmt(row.actual)}</td><td>${fmt(row.budget)}</td><td class="${signed(row.gap)}">${fmt(row.gap)}</td><td class="${signed(row.rate)}">${pct(row.rate)}</td></tr>`;
}
function donut(groups) {
  const vals = groups.map((g) => g.actual);
  if (vals.some((v) => v === null || v < 0) || !vals.some((v) => v > 0))
    return empty("Seçim için gider dağılımı hesaplanamıyor.");
  const sum = vals.reduce((a, b) => a + b, 0);
  let angle = -Math.PI / 2,
    out = "";
  const cx = 165,
    cy = 145,
    r = 90;
  const points = [];
  groups.forEach((g, i) => {
    const share = g.actual / sum;
    if (!share) return;
    const end = angle + share * Math.PI * 2,
      mid = (angle + end) / 2;
    const x1 = cx + Math.cos(angle) * r,
      y1 = cy + Math.sin(angle) * r,
      x2 = cx + Math.cos(end) * r,
      y2 = cy + Math.sin(end) * r;
    if (share > 0.999999)
      out += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${palette[i % palette.length]}" stroke-width="34"/>`;
    else
      out += `<path d="M${x1},${y1} A${r},${r} 0 ${share > 0.5 ? 1 : 0} 1 ${x2},${y2}" fill="none" stroke="${palette[i % palette.length]}" stroke-width="34"><title>${esc(g.name)}: ${pct(share)}</title></path>`;
    points.push({ x: cx + Math.cos(mid) * 125, y: cy + Math.sin(mid) * 125, share });
    angle = end;
  });
  for (let i = 0; i < points.length; i++)
    for (let j = 0; j < i; j++)
      if (Math.abs(points[i].x - points[j].x) < 60 && Math.abs(points[i].y - points[j].y) < 18)
        points[i].y += 20;
  out += points
    .map((p) => `<text x="${p.x}" y="${p.y}" text-anchor="middle" class="value-label">${pct(p.share)}</text>`)
    .join("");
  out += `<text x="${cx}" y="${cy - 3}" text-anchor="middle">GERÇEKLEŞEN</text><text x="${cx}" y="${cy + 17}" text-anchor="middle">GİDER DAĞILIMI</text>`;
  return `<div class="chart"><svg viewBox="0 0 330 290" role="img" aria-label="Gerçekleşen giderlerin ana gider dağılımı">${out}</svg></div><div class="donut-legend">${groups.map((g, i) => `<div><i style="background:${palette[i % palette.length]}"></i>${esc(g.name)}<b>${pct(g.actual / sum)}</b></div>`).join("")}</div>`;
}
function renderDetail() {
  let rows = "";
  for (const g of data.detail.groups) {
    rows += detailRow(g, "group");
    if (groupsOpen.has(g.name)) rows += g.items.map((r) => detailRow(r, "detail")).join("");
  }
  rows += detailRow({ name: "TOPLAM", ...data.detail.total, history: [] }, "total");
  return `<div class="detail-grid">${panel("Ana ve alt giderler", `<div class="scroll-table"><table><thead><tr><th>Gider kalemi</th><th>Gerçekleşen<br>Milyon TL</th><th>Bütçe<br>Milyon TL</th><th>Bütçe − gerçekleşen<br>Milyon TL</th><th>Sapma %</th></tr></thead><tbody>${rows}</tbody></table></div>`, "Gerçekleşen hücresinin üzerine gelerek seçili yılların trendini görün.")}<section class="panel donut-panel"><div class="panel-head"><h3>Ana giderlerin payı</h3></div>${donut(data.detail.groups)}</section></div>`;
}
function localList(key) {
  const current = state[key] ?? meta.options.department;
  return `<div class="local-filter"><details><summary>Tablo bölümleri · ${current.length} seçili</summary><div class="listbox"><div class="filter-heading"><button data-local-all="${key}">Tümü</button><button data-local-none="${key}">Temizle</button></div>${meta.options.department.map((n, i) => `<label><input type="checkbox" data-local="${key}" data-index="${i}" ${current.includes(n) ? "checked" : ""}>${esc(n)}</label>`).join("")}</div></details></div>`;
}
function matrix(rows, fields) {
  return `<div class="scroll-table"><table><thead><tr><th>Bölüm / düzey</th>${fields.map((f) => `<th>${esc(f.label)}</th>`).join("")}</tr></thead><tbody>${
    rows
      .map(
        (r) =>
          `<tr class="${r.kind}"><td>${esc(r.name)}</td>${fields
            .map((f) => {
              const v = r.values.find((v) => v.year === f.year)?.[f.key];
              return `<td class="${f.color ? signed(v) : ""}">${f.percent ? pct(v) : fmt(v)}</td>`;
            })
            .join("")}</tr>`,
      )
      .join("") || `<tr><td colspan="20">Karşılaştırılabilir kayıt yok.</td></tr>`
  }</tbody></table></div>`;
}
function renderDepartments() {
  const ratio = departmentMode === "opex",
    years = data.selection.years,
    fields = [];
  for (const year of years) {
    fields.push({
      year,
      key: "actual",
      label: `${year} Gerçekleşen${ratio ? "" : " · Milyon TL"}`,
      percent: ratio,
    });
    if (year === meta.budgetYear) {
      fields.push({
        year,
        key: "budget",
        label: `${year} Bütçe${ratio ? "" : " · Milyon TL"}`,
        percent: ratio,
      });
      if (ratio)
        fields.push(
          {
            year,
            key: "annualImprovement",
            label: `${year - 1} yılına göre iyileşme %`,
            percent: true,
            color: true,
          },
          { year, key: "budgetImprovement", label: "Bütçeye göre iyileşme %", percent: true, color: true },
        );
      else
        fields.push(
          { year, key: "gap", label: "Bütçe − gerçekleşen · Milyon TL", color: true },
          { year, key: "rate", label: "Sapma %", percent: true, color: true },
        );
    }
  }
  const tools = `<div class="section-tools">${segments(
    "department",
    [
      ["expense", "Bölüm bazlı gider tutarı"],
      ["opex", "OPEX / ciro oranı"],
    ],
    departmentMode,
  )}${ratio ? localList("opexLocal") : ""}</div>`;
  return (
    tools +
    panel(
      ratio ? "Bölüm ve müdürlük bazında OPEX / şirket cirosu" : "Bölüm ve müdürlük bazında gider",
      matrix(data.departments[departmentMode], fields),
      ratio
        ? "Önceki yıl verisi olmayan bölümler kıyastan çıkarılır. Şirket cirosu bir kez payda olarak kullanılır."
        : "Tutarlar milyon TL. Müdürlük alt toplamları ilgili bölümlerden sonra gelir.",
    )
  );
}
function renderUnit() {
  const years = data.selection.years,
    active = years.includes(meta.budgetYear),
    fields = [];
  for (const year of years)
    if (!active || ![meta.budgetYear - 1, meta.budgetYear].includes(year))
      fields.push({ year, key: "actual", label: `${year} Gerçekleşen BM` });
  if (active) {
    const year = meta.budgetYear;
    fields.push(
      ...[
        ["prior", `${year - 1} Gerçekleşen BM`],
        ["actual", `${year} Gerçekleşen BM`],
        ["budget", `${year} Bütçe BM`],
        ["inflation", `${year - 1}–${year} Enflasyon %`],
        ["adjusted", `Endekslenmiş ${year - 1} BM`],
        ["rate", "Reel iyileşme %"],
        ["weight", `${year} Direktörlük gider ağırlığı %`],
        ["managerWeight", "Müdürlük gider ağırlığı %"],
      ].map(([key, label]) => ({
        year,
        key,
        label,
        percent: ["inflation", "rate", "weight", "managerWeight"].includes(key),
        color: key === "rate",
      })),
    );
  }
  return (
    `<div class="section-tools"><span class="control-note">Endeks eşleşmeli ${data.unit.eligible.length} bölüm · Ağırlıklar karşılaştırılabilir kapsamda yeniden hesaplanır.</span>${localList("bmLocal")}</div>` +
    panel(
      "Endekslenmiş birim maliyet karşılaştırması",
      matrix(data.unit.rows, fields),
      "TL / işlem · Endeks, seçili aylardaki yıllık oranların aritmetik ortalamasıdır.",
    ) +
    `<div class="summary-wide"><div><h3>Genel reel iyileşme</h3><p>Seçili karşılaştırılabilir bölümler · Güncel gider ağırlıklarıyla</p></div><div class="number ${signed(data.unit.rate)}">${pct(data.unit.rate)}</div></div>`
  );
}
function renderPriority() {
  const p = data.priority,
    filtered = p.rows.filter((r) =>
      `${r.department} ${r.item} ${r.group}`
        .toLocaleLowerCase("tr")
        .includes(prioritySearch.toLocaleLowerCase("tr")),
    );
  const ready = data.selection.years.length === 1 && data.selection.years[0] === meta.budgetYear;
  const table = `<div class="scroll-table"><table><thead><tr><th>Bölüm / alt gider</th><th>Gerçekleşen<br>Milyon TL</th><th>Bütçe<br>Milyon TL</th><th>Sapma<br>Milyon TL</th><th>Sapma %</th><th>Açıklama</th></tr></thead><tbody>${filtered.map((r, i) => `<tr><td><span class="rank">${p.rows.indexOf(r) + 1}</span><strong>${esc(r.department)}</strong><br><span class="metric-label">${esc(r.item)}</span></td><td>${fmt(r.actual)}</td><td>${fmt(r.budget)}</td><td class="bad">${fmt(r.gap)}</td><td class="bad">${pct(r.rate)}</td><td><button class="note-button ${r.status === "pending" ? "pending-note" : ""}" data-note="${p.rows.indexOf(r)}">${esc(r.note || (r.stale ? "Tutar değişti · yeniden açıklama gir" : "＋ Açıklama ekle"))}</button></td></tr>`).join("") || `<tr><td colspan="6">${!ready ? "Bütçe yılını tek başına seçin." : p.incomplete ? "Eksik gider/bütçe verisi nedeniyle sıralama oluşturulmadı." : "Seçimde bütçe aşımı yok."}</td></tr>`}</tbody></table></div>`;
  return (
    `<div class="cards">${card("Brüt bütçe aşımı", fmt(p.gross), "Milyon TL", p.gross > 0 ? "bad" : "")}${card("Net bütçeyi aşan bölüm", fmt(p.netDepartments, 0), "bölüm")}${card("En büyük 5 aşımın payı", pct(p.top5), "Brüt bütçe aşımı içindeki pay")}${card("Bölüm–gider aşım kaydı", ready ? fmt(p.rows.length, 0) : "—", "kayıt")}</div><div class="section-tools"><input class="mini-search" id="prioritySearch" value="${esc(prioritySearch)}" placeholder="Bölüm veya gider ara" aria-label="Sapma tablosunda ara"><div><button data-comments="review">Açıklamaları toplu düzenle</button> <button id="importNotesBtn">Açıklama yükle</button> <button id="mergeNotesBtn">Rapor açıklamalarını birleştir</button></div></div>` +
    panel(
      "En büyük 10 bütçe aşımı",
      bars(
        p.rows.slice(0, 10).map((r) => ({ name: `${r.department} · ${r.item}`, actual: r.overrun })),
        { width: 1000, labelWidth: 340 },
      ),
      "Bölüm–alt gider bazında brüt aşım · Milyon TL",
    ) +
    `<div style="height:18px"></div>` +
    panel(
      "Bütçe aşımı ve açıklamalar",
      table,
      `Açıklamalar tam seçili ay kümesine bağlıdır. ${data.noteCount} kayıtlı açıklama.`,
    )
  );
}
function showModal(title, body) {
  $("#modalTitle").textContent = title;
  $("#modalBody").innerHTML = body;
  if (!$("#modal").open) $("#modal").showModal();
}
function field(label, id, value = "", type = "text", full = false, mode = "") {
  return `<label class="form-field ${full ? "full" : ""}"><span>${label}</span><div class="input-row"><input id="${id}" type="${type}" value="${esc(value)}">${mode ? `<button type="button" data-pick="${id}" data-pick-mode="${mode}">Gözat</button>` : ""}</div></label>`;
}
function sources() {
  sourcePreview = meta.configuration;
  showModal(
    "Kaynaklar ve eşleştirmeler",
    `<div class="form-grid">${field("Ayar Excel dosyası", "settingPath", meta.settings, "text", true, "file")}<div class="form-field full"><button id="previewSettingsBtn">Seçilen ayarları oku</button></div>${field("Maliyet dosyaları klasörü", "costFolder", sourcePreview.locations.costFolder, "text", true, "folder")}${field("Ciro ve tutar kontrol dosyası", "controlFile", sourcePreview.locations.controlFile, "text", true, "file")}${field("Endeks dosyası", "indexFile", sourcePreview.locations.indexFile, "text", true, "file")}${field("Kaynak kapanış yılı", "closingYear", meta.closing[0], "number")}${field("Kaynak kapanış ayı", "closingMonth", meta.closing[1], "number")}</div><p class="control-note">Konumları bir kez kaydedin; sonraki açılışlarda hatırlanır. Değişiklikte yeniden seçebilirsiniz. Konum kaydı öncesinde ayar Excel’inin yedeği alınır; eşleştirme sayfaları korunur.</p><div id="mappingPreview">${mappingPreview(sourcePreview)}</div><details><summary>Dosya adında yıl / ay kontrolü</summary>${rawTable(["Dosya", "Dönem", "Seçim sonucu"], meta.files || [])}</details><details><summary>Okunan bölüm kaynakları · ${meta.sources.length}</summary>${rawTable(
      ["Bölüm", "Dosya"],
      meta.sources.map((s) => [s.department, s.file]),
    )}</details><div class="modal-actions"><button id="saveLocationsBtn" class="primary">Konumları kaydet ve yenile</button><button id="loadSourcesBtn">Yalnız kaynakları yenile</button></div>`,
  );
}
function mappingPreview(config) {
  return `${config.error ? `<div class="notice">${esc(config.error)}</div>` : ""}<details open><summary>Tüm bölüm eşleştirmeleri · ${config.mappings.length} kayıt</summary>${rawTable(
    ["Bölüm", "Müdürlük", "Direktörlük", "Müdürlük altı", "Lokasyon", "Kontrol bölümü", "Endeks adı"],
    config.mappings.map((m) => [
      m.department,
      m.mudurluk,
      m.direktorluk,
      m.mudurluk_alti,
      m.lokasyon,
      m.kontrol_bolumu,
      m.endeks,
    ]),
  )}</details><details><summary>Ana gider grupları</summary>${rawTable(["İlk indeks", "Son indeks", "Ana gider adı"], config.expenseGroups || [])}</details><details><summary>Gider birleştirme eşleştirmeleri</summary>${rawTable(["Kaynak indeks", "Rapor indeksi", "Rapor gider adı"], config.expenseMappings || [])}</details><details><summary>Okunmayacak ortak sayfalar</summary>${rawTable(
    ["Sayfa"],
    (config.excludedSheets || []).map((n) => [n]),
  )}</details><p class="control-note">Bölüm, müdürlük, lokasyon, kontrol/endeks adları ve gider grupları ayar Excel’indeki tüm kayıtlardan okunur. Yeni bölümün kaynak sayfası ve eşleştirmesi eklendikten sonra yenilemek yeterlidir.</p>`;
}
async function readSettingsPreview() {
  sourcePreview = await api("settings-preview", { settings: $("#settingPath").value });
  $("#settingPath").value = sourcePreview.settings;
  for (const key of ["costFolder", "controlFile", "indexFile"])
    $("#" + key).value = sourcePreview.locations[key];
  $("#mappingPreview").innerHTML = mappingPreview(sourcePreview);
}
function reportDialog() {
  showModal(
    "Excel raporlarını oluştur",
    `<div class="form-grid">${field("Çıktı klasörü", "outputPath", meta.output, "text", true, "folder")}<label class="form-field"><span>Rapor yılı</span><select id="reportYear">${meta.options.year.map((y) => `<option ${y === (data.selection.years.length === 1 ? data.selection.years[0] : meta.budgetYear) ? "selected" : ""}>${y}</option>`).join("")}</select></label></div><p class="control-note">Aylar: ${esc(data.selection.months.map((m) => months[m - 1]).join(", ") || "Ay seçilmedi")}. Rapor kapsamı: ayar dosyasındaki tüm bölümler. Her müdürlük için ayrı dosya ve bir genel dosya oluşturulur.</p><p class="control-note">Kaynaklar yeniden okunur. Çıktılar seçilen klasör altında tarih/saat içeren yeni bir klasöre kaydedilir.</p><div class="modal-actions"><button id="createReportsBtn" class="primary">Raporları oluştur</button></div>`,
  );
}
async function runJob(kind, args) {
  await api("job", { kind, args });
  $("#modalBody").innerHTML =
    '<div class="progress" id="jobProgress" role="status">İşlem hazırlanıyor…</div><button id="cancelJobBtn">Güvenli durdur</button>';
  while (true) {
    await new Promise((r) => setTimeout(r, 400));
    const job = await api("job");
    if (job.state === "running") {
      $("#status").textContent = job.stage;
      if ($("#jobProgress"))
        $("#jobProgress").textContent = job.stage + (job.cancelRequested ? " · Durdurma bekleniyor" : "");
      continue;
    }
    if (job.state === "error" || job.state === "cancelled") throw Error(job.error);
    if (job.state !== "done") throw Error("İşlem durumu alınamadı.");
    const result = job.result;
    if (result.meta) {
      meta = result.meta;
      if (kind === "load" || kind === "configure" || kind === "layout") {
        defaults();
        buildFilters();
      }
      await refresh();
    }
    $("#modalBody").innerHTML =
      `<div class="path-result"><strong>${esc(result.message || "İşlem tamamlandı.")}</strong>${result.folder ? `<p>Kaydedilen klasör: ${esc(result.folder)}</p>` : ""}${result.backup ? `<p>Ayar yedeği: ${esc(result.backup)}</p>` : ""}${(result.downloads || []).map((d) => `<a href="${esc(d.url)}" download>${esc(d.name)} ↓</a>`).join("")}</div>`;
    if (kind === "import") await refresh();
    return result;
  }
}
function showAudit() {
  const a = data.audit;
  showModal(
    "Doğrulama ve revizyonlar",
    `<p class="control-note">${a.revisions.length} geriye dönük alt gider/bütçe değişikliği · ${a.warnings.length} kaynak uyarısı. Kontroller yalnız dip toplamdadır.</p><details open><summary>Dip toplam kontrolleri</summary>${rawTable(["Kontrol bölümü", "Müdürlük", "Direktörlük", "Dönem", "Kaynak TL", "Kontrol TL", "Fark TL", "Sonuç", "Kaynak", "Bölümler"], a.checks)}</details><details><summary>Geriye dönük alt gider değişiklikleri</summary>${rawTable(["Yeni kapanış", "Etkilenen ay", "Müdürlük", "Bölüm", "Gider kodu", "Alt gider", "Tür", "Önce TL", "Sonra TL", "Değişiklik TL", "Eski kaynak", "Yeni kaynak"], a.revisions)}</details><details><summary>Kaynak uyarıları</summary>${rawTable(["Dönem", "Bölüm", "Dosya", "Uyarı"], a.warnings)}</details>`,
  );
}
function rawTable(headers, rows) {
  return `<div class="scroll-table"><table><thead><tr>${headers.map((h) => `<th>${esc(h)}</th>`).join("")}</tr></thead><tbody>${rows.map((row) => `<tr>${row.map((v) => `<td>${typeof v === "number" ? fmt(v) : esc(v ?? "—")}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`;
}
function hover(cell, event) {
  const info = JSON.parse(cell.dataset.hover);
  if (!info.history.length) return;
  const chart = lineChart(
    [{ name: info.name, color: palette[0], values: info.history.map((h) => h.actual) }],
    info.history.map((h) => String(h.year)),
    { width: 360, height: 180, labels: true },
  );
  const tip = $("#tooltip");
  tip.innerHTML = `<h3>${esc(info.name)}</h3><p>Aynı seçili aylar · Milyon TL</p>${chart}`;
  tip.hidden = false;
  const box = cell.getBoundingClientRect();
  const x = event?.clientX ?? box.left + box.width / 2,
    y = event?.clientY ?? box.bottom;
  tip.style.left = Math.max(8, Math.min(window.innerWidth - 342, x + 15)) + "px";
  tip.style.top = Math.max(8, Math.min(window.innerHeight - 230, y + 13)) + "px";
}
document.addEventListener("change", (e) => {
  const el = e.target;
  if (el.dataset.commentSelect !== undefined) {
    const i = Number(el.dataset.commentSelect);
    el.checked ? commentSelected.add(i) : commentSelected.delete(i);
    $("#commentSelectedCount").textContent = commentSelected.size + " satır seçili";
  }
  if (el.id === "commentFilter") {
    commentFilter = el.value;
    commentSelected.clear();
    renderComments();
  }
  if (el.dataset.filter) {
    const key = el.dataset.filter,
      value = meta.options[key][Number(el.dataset.index)],
      values = new Set(selected(key));
    el.checked ? values.add(value) : values.delete(value);
    state.filters[key] =
      !["year", "month"].includes(key) && values.size === meta.options[key].length ? null : [...values];
    buildFilters();
    schedule();
  }
  if (el.dataset.local) {
    const key = el.dataset.local,
      value = meta.options.department[Number(el.dataset.index)],
      values = new Set(state[key] ?? meta.options.department);
    el.checked ? values.add(value) : values.delete(value);
    state[key] = values.size === meta.options.department.length ? null : [...values];
    schedule();
  }
});
document.addEventListener("click", async (e) => {
  const b = e.target.closest("button");
  if (!b) return;
  if (
    commentOpen &&
    commentDirty.size &&
    (b.id === "modalClose" ||
      b.dataset.page ||
      ["sourceBtn", "workspaceBtn", "reportBtn", "carryPreviousBtn"].includes(b.id))
  ) {
    if (!window.confirm("Kaydedilmemiş taslaklar var. Kaydetmeden çıkılsın mı?")) return;
    commentDirty.clear();
    commentOpen = false;
  }
  try {
    if (b.dataset.comments) {
      if (b.dataset.comments === "review") await showComments();
      return;
    }

    if (b.dataset.layoutCell) {
      const i = Number($("#layoutTarget").value),
        param = layoutData.parameters[i],
        coord = b.dataset.layoutCell,
        match = coord.match(/^([A-Z]+)([0-9]+)$/);
      let value = coord;
      if (param[1].includes("satiri")) value = Number(match[2]);
      else if (param[1].includes("sutunu"))
        value = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
      else if (param[1] === "sayfa") value = layoutData.sheet;
      else if (param[1] === "gider_kalem_sayisi" || param[1] === "tolerans")
        throw Error("Bu alanı parametre tablosunda elle girin.");
      const input = $$("[data-layout-index]").find((el) => Number(el.dataset.layoutIndex) === i);
      if (input) input.value = value;
      toast(param[1] + " için " + value + " seçildi; kaydetmeden önce kontrol edin.");
      return;
    }
    if (b.dataset.replay) {
      const output = $("#historyOutput").value;
      await runJob("replay", { id: b.dataset.replay, output });
      return;
    }
    if (b.dataset.page) {
      page = b.dataset.page;
      render();
      return;
    }
    if (b.dataset.filterAll) {
      const k = b.dataset.filterAll;
      state.filters[k] = selected(k).length === meta.options[k].length ? [] : null;
      buildFilters();
      schedule();
      return;
    }
    if (b.dataset.mode) {
      if (b.dataset.mode === "staff") staffMode = b.dataset.value;
      else departmentMode = b.dataset.value;
      render();
      return;
    }
    if (b.dataset.group) {
      const g = b.dataset.group;
      groupsOpen.has(g) ? groupsOpen.delete(g) : groupsOpen.add(g);
      render();
      return;
    }
    if (b.dataset.year) {
      const y = Number(b.dataset.year);
      yearsClosed.has(y) ? yearsClosed.delete(y) : yearsClosed.add(y);
      render();
      return;
    }
    if (b.dataset.localAll || b.dataset.localNone) {
      state[b.dataset.localAll || b.dataset.localNone] = b.dataset.localAll ? null : [];
      await refresh();
      return;
    }
    if (b.dataset.pick) {
      b.disabled = true;
      const result = await api("pick", { mode: b.dataset.pickMode });
      if (result.value) {
        $("#" + b.dataset.pick).value = result.value;
        if (b.dataset.pick === "settingPath") await readSettingsPreview();
      }
      return;
    }
    if (b.dataset.note) {
      const r = data.priority.rows[Number(b.dataset.note)];
      showModal(
        `${r.department} · ${r.item}`,
        `<label class="form-field"><span>Seçili aylar için açıklama</span><textarea id="noteText" maxlength="32767">${esc(r.note)}</textarea></label><p class="control-note">${esc(data.selection.months.map((m) => months[m - 1]).join(", "))} ${meta.budgetYear} · Tutarlar değişirse açıklama yeniden doğrulama ister.</p><div class="modal-actions"><button id="saveNoteBtn" data-department="${esc(r.department)}" data-code="${r.code}" class="primary">Açıklamayı kaydet</button></div>`,
      );
      return;
    }
    if (b.dataset.zoom) {
      const command = b.dataset.zoom,
        vals = data.general.trend.flatMap((r) => [r.actual, r.budget]).filter((v) => v !== null);
      if (!vals.length) return;
      const lo = Math.min(...vals),
        hi = Math.max(...vals),
        pad = (hi - lo) * 0.17 || 1;
      let range = trendBounds ?? [Math.max(0, lo - pad), hi + pad];
      let [a, z] = range;
      const mid = (a + z) / 2,
        span = z - a;
      if (command === "apply") {
        a = Number($("#ymin").value);
        z = Number($("#ymax").value);
        if (!(z > a)) throw Error("Üst sınır alt sınırdan büyük olmalı.");
        trendBounds = [a, z];
      } else if (command === "reset") trendBounds = null;
      else if (command === "in" || command === "out") {
        const factor = command === "in" ? 0.65 : 1.5;
        trendBounds = [mid - (span * factor) / 2, mid + (span * factor) / 2];
      } else
        trendBounds = [
          a + span * 0.2 * (command === "up" ? 1 : -1),
          z + span * 0.2 * (command === "up" ? 1 : -1),
        ];
      render();
      return;
    }
    switch (b.id) {
      case "resetBtn":
        defaults();
        buildFilters();
        await refresh();
        break;
      case "carryPreviousBtn": {
        const r = await api("comment-carry", { state });
        toast(
          `${r.created} geçmiş yorum getirildi; ${r.skipped} mevcut kayıt korundu. ${r.issues.join(" ")}`,
        );
        await showComments();
        break;
      }
      case "commentFilesBtn": {
        if (commentDirty.size) throw Error("Önce taslakları kaydedin veya onaylayın.");
        const picked = await api("pick", { mode: "files" });
        if (picked.value.length) {
          const r = await api("comment-carry", { state, paths: picked.value });
          toast(
            `${r.created} geçmiş yorum getirildi; ${r.skipped} mevcut kayıt korundu. ${r.issues.join(" ")}`,
          );
          await showComments();
        }
        break;
      }
      case "commentSelectBtn":
        for (const i of commentVisible()) if (!commentReview.rows[i].protected) commentSelected.add(i);
        renderComments();
        break;
      case "commentClearBtn":
        commentSelected.clear();
        renderComments();
        break;
      case "commentApproveBtn":
        await saveCommentBatch("approve");
        break;
      case "commentRejectBtn":
        await saveCommentBatch("reject");
        break;
      case "commentSaveBtn":
        await saveCommentBatch("draft", true);
        break;
      case "commentSameBtn": {
        const text = $("#commentCommonText").value;
        if (!text.trim()) throw Error("Seçili satırlara uygulanacak metni yazın.");
        for (const i of commentVisible())
          if (commentSelected.has(i) && !commentReview.rows[i].protected) {
            commentReview.rows[i].draft = text;
            commentDirty.add(i);
          }
        renderComments();
        break;
      }
      case "workspaceBtn":
        await showWorkspace();
        break;
      case "cancelJobBtn":
        await api("cancel", {});
        toast("Mevcut dosya adımı sonunda duracak.");
        break;
      case "saveViewBtn":
        await api("views", { name: $("#viewName").value, state });
        await showWorkspace();
        break;
      case "loadViewBtn": {
        const v = workspaceData.views[$("#savedView").value];
        if (!v) throw Error("Bir görünüm seçin.");
        state = JSON.parse(JSON.stringify(v.state));
        for (const key of Object.keys(filterNames)) {
          const values = state.filters[key];
          if (Array.isArray(values)) {
            if (
              !["year", "month"].includes(key) &&
              JSON.stringify([...values].sort()) === JSON.stringify([...(v.options[key] || [])].sort())
            )
              state.filters[key] = null;
            else state.filters[key] = values.filter((x) => meta.options[key].includes(x));
          }
        }
        for (const key of ["bmLocal", "opexLocal"])
          if (Array.isArray(state[key]))
            state[key] = state[key].filter((n) => meta.options.department.includes(n));
        buildFilters();
        await refresh();
        $("#modal").close();
        break;
      }
      case "deleteViewBtn":
        await api("views", { name: $("#savedView").value, delete: true });
        await showWorkspace();
        break;
      case "acceptMappingsBtn": {
        const renames = $$("[data-rename-new]")
          .filter((el) => el.value)
          .map((el) => ({ old: el.value, new: el.dataset.renameNew }));
        const result = await api("mapping-accept", {
          fingerprint: workspaceData.review.fingerprint,
          renames,
        });
        toast(result.message);
        await showWorkspace();
        break;
      }
      case "traceBtn":
        await showTrace();
        break;
      case "traceRunBtn":
        await showTrace($("#traceDepartment").value, Number($("#traceCode").value) || null);
        break;
      case "layoutBtn":
        await showLayout();
        break;
      case "saveLayoutBtn": {
        const parameters = $$("[data-layout-index]").map((el) => {
          const old = layoutData.parameters[Number(el.dataset.layoutIndex)];
          return [old[0], old[1], typeof old[2] === "number" ? Number(el.value) : el.value];
        });
        await runJob("layout", { parameters, fingerprint: layoutData.fingerprint });
        break;
      }
      case "workspaceSourcesBtn":
      case "sourceBtn":
        meta = await api("meta");
        sources();
        break;
      case "workspaceReportBtn":
      case "reportBtn":
        reportDialog();
        break;
      case "modalClose":
        $("#modal").close();
        break;
      case "auditBtn":
        showAudit();
        break;
      case "saveFiltersBtn":
        await api("save-selection", state);
        toast("Ortak seçimler ve yerel tablo kapsamları kaydedildi.");
        break;
      case "loadSourcesBtn":
        await runJob("load", {
          settings: $("#settingPath").value,
          closing: [Number($("#closingYear").value), Number($("#closingMonth").value)],
        });
        break;
      case "previewSettingsBtn":
        await readSettingsPreview();
        break;
      case "saveLocationsBtn":
        await runJob("configure", {
          settings: $("#settingPath").value,
          expectedSettings: sourcePreview.settings,
          fingerprint: sourcePreview.fingerprint,
          locations: Object.fromEntries(
            ["costFolder", "controlFile", "indexFile"].map((key) => [key, $("#" + key).value]),
          ),
          closing: [Number($("#closingYear").value), Number($("#closingMonth").value)],
        });
        break;
      case "createReportsBtn":
        await runJob("report", {
          year: Number($("#reportYear").value),
          months: data.selection.months,
          output: $("#outputPath").value,
        });
        break;
      case "saveNoteBtn":
        await api("note", {
          state,
          department: b.dataset.department,
          code: Number(b.dataset.code),
          note: $("#noteText").value,
        });
        $("#modal").close();
        await refresh();
        toast("Açıklama kaydedildi.");
        break;
      case "importNotesBtn": {
        const result = await api("pick", { mode: "files" });
        if (!result.value.length) break;
        showModal("Açıklamalar yükleniyor", "");
        await runJob("import", { paths: result.value, state });
        break;
      }
      case "mergeNotesBtn":
        showModal(
          "Müdürlük açıklamalarını genel raporda birleştir",
          `<div class="form-grid">${field("Genel rapor", "mergeGeneral", "", "text", true, "file")}<label class="form-field full"><span>Müdürlük raporları (her satıra bir tam dosya yolu)</span><textarea id="mergeManagers"></textarea><button data-pick-managers="true">Müdürlük dosyalarını seç</button></label>${field("Çıktı klasörü", "mergeOutput", meta.output, "text", true, "folder")}</div><div class="modal-actions"><button id="mergeRunBtn" class="primary">Açıklamaları birleştir</button></div>`,
        );
        break;
      case "mergeRunBtn":
        await runJob("merge", {
          general: $("#mergeGeneral").value,
          managers: $("#mergeManagers")
            .value.split("\n")
            .map((s) => s.trim())
            .filter(Boolean),
          output: $("#mergeOutput").value,
        });
        break;
      case "closeBtn":
        await api("shutdown", {});
        showModal("Uygulama kapatıldı", "<p>Yerel işlem durduruldu. Bu sekmeyi kapatabilirsiniz.</p>");
        break;
    }
    if (b.dataset.pickManagers) {
      const result = await api("pick", { mode: "files" });
      $("#mergeManagers").value = result.value.join("\n");
    }
  } catch (err) {
    toast(err.message);
    if ($("#modal").open && $("#modalBody .progress"))
      $("#modalBody").innerHTML =
        `<div class="notice">${esc(err.message)}</div><p class="control-note">Kaynaklar veya Rapor oluştur düğmesiyle yeniden deneyebilirsiniz.</p>`;
  } finally {
    b.disabled = false;
  }
});
document.addEventListener("input", (e) => {
  if (e.target.dataset.commentDraft !== undefined) {
    const i = Number(e.target.dataset.commentDraft);
    commentReview.rows[i].draft = e.target.value;
    commentDirty.add(i);
    commentSelected.add(i);
    const box = $(`[data-comment-select="${i}"]`);
    if (box) box.checked = true;
    $("#commentSelectedCount").textContent =
      commentSelected.size + " satır seçili · Kaydedilmemiş düzenleme var";
  }
  if (e.target.id === "commentSearch") {
    commentSearch = e.target.value;
    commentSelected.clear();
    renderCommentRows();
  }
  if (e.target.id === "prioritySearch") {
    const start = e.target.selectionStart;
    prioritySearch = e.target.value;
    $("#content").innerHTML = renderPriority();
    $("#prioritySearch").focus();
    $("#prioritySearch").setSelectionRange(start, start);
  }
});
document.addEventListener("mouseover", (e) => {
  const cell = e.target.closest("[data-hover]");
  if (cell) hover(cell, e);
});
document.addEventListener("mouseout", (e) => {
  if (e.target.closest("[data-hover]")) $("#tooltip").hidden = true;
});
document.addEventListener("focusin", (e) => {
  if (e.target.dataset.hover) hover(e.target);
});
document.addEventListener("focusout", (e) => {
  if (e.target.dataset.hover) $("#tooltip").hidden = true;
});

async function showWorkspace() {
  workspaceData = await api("workspace", { state });
  const q = workspaceData.quality,
    r = workspaceData.review;
  showModal(
    "Çalışma merkezi",
    `${workspaceData.currentJob?.state === "running" ? `<div class="progress">${esc(workspaceData.currentJob.stage)}</div><button id="cancelJobBtn">Güvenli durdur</button>` : ""}<div class="cards">${card("Beklenen bölüm", fmt(q.expectedDepartments, 0), "Ayar Excel’i")}${card("Okunan bölüm", fmt(q.receivedDepartments, 0), "Seçili kapanış")}${card("İncelenecek konu", fmt(q.issues.length, 0), "Kaynak ve dip toplam", q.issues.length ? "bad" : "good")}${card("Revizyon", fmt(q.revisions.length, 0), "Seçili bölüm kapsamı")}</div><p class="control-note">Kapanış: ${meta.closing.join("-")} · Raporlar tüm eşleşmiş bölümler için üretilir. Dashboard kapsamı: ${q.scope.departments.length} bölüm; ${q.scope.years.join(", ")}; aylar ${q.scope.months.join(", ")}.</p><div class="modal-actions"><button id="workspaceSourcesBtn">1 · Kaynakları seç / yenile</button><button id="traceBtn">Hesap nasıl oluştu?</button><button id="layoutBtn">Tablo düzeni</button><button id="workspaceReportBtn" class="primary">2 · Rapor oluştur</button></div><details open><summary>Kontrol merkezi</summary>${rawTable(
      ["Tür", "Bölüm", "Dönem", "Dosya", "Sayfa", "Konum", "Sorun", "Yapılacak işlem"],
      q.issues.map((i) => [i.kind, i.department, i.period, i.file, i.sheet, i.cell, i.message, i.action]),
    )}${q.issues.length ? "" : "<p>Seçili kapsamda kaynak uyarısı veya dip toplam farkı bulunmadı.</p>"}</details><details><summary>Hesaba alınmayan bölümler</summary>${rawTable(["Bölüm", "Hesap", "Gerekçe"], q.exclusions)}</details><details><summary>Eşleştirme değişiklikleri</summary><p>Son onay: ${esc(r.baseline.acceptedAt || "İlk inceleme")}. Öneriler otomatik uygulanmaz. İsim değişikliği varsa eski bölümü seçin; açıklamalar yalnız açık onayla taşınır.</p>${rawTable(["Durum", "Bölüm", "Açıklama"], [...r.added.map((n) => ["Eklendi", n, "Kaynak sayfası ve eşleştirmesi kontrol edilmeli"]), ...r.removed.map((n) => ["Kaldırıldı", n, "Eski kayıtlar geçmişte korunur"]), ...r.changed.map((n) => ["Eşleştirme değişti", n.department, JSON.stringify(n.before) + " → " + JSON.stringify(n.after)])])}${r.added.map((n) => `<label class="form-field"><span>${esc(n)} için eski ad (varsa)</span><select data-rename-new="${esc(n)}"><option value="">Yeni bölüm / taşıma yok</option>${r.removed.map((old) => `<option value="${esc(old)}">${esc(old)}</option>`).join("")}</select></label>`).join("")}${r.suggestions.length ? `<p>Olası isim değişiklikleri: ${r.suggestions.map((x) => esc(x.old) + " → " + esc(x.new)).join("; ")}. Kesin eşleştirme değildir.</p>` : ""}<button id="acceptMappingsBtn">İnceledim; eşleştirme durumunu kaydet</button></details><details><summary>Kayıtlı görünümler</summary><div class="form-grid">${field("Yeni görünüm adı", "viewName")}<label class="form-field"><span>Kayıtlı görünüm</span><select id="savedView">${Object.keys(
      workspaceData.views,
    )
      .map((n) => `<option>${esc(n)}</option>`)
      .join(
        "",
      )}</select></label></div><div class="modal-actions"><button id="saveViewBtn">Mevcut seçimi adlandır ve kaydet</button><button id="loadViewBtn">Seçili görünümü aç</button><button id="deleteViewBtn">Görünümü sil</button></div></details><details><summary>Çalıştırma geçmişi ve girdi kopyaları</summary>${field("Yeniden üretim çıktı konumu", "historyOutput", meta.output, "text", true, "folder")}<p>Her raporda kullanılan girdiler ayrı kopyalanır. Yeniden üretim o kopyaları kullanır; güncel dosyalara dönmez. Kopyalar disk alanı kullanır. Aynı ayar yolunun bir önceki çalışmasıyla değişen kaynaklar karşılaştırılır.</p><div class="scroll-table"><table><thead><tr><th>Çalışma</th><th>Durum</th><th>Dönem</th><th>Değişen / kaldırılan kaynak</th><th>Sonuç</th><th>İşlem</th></tr></thead><tbody>${workspaceData.history.map((h) => `<tr><td>${esc(h.started)}</td><td>${esc(h.state)}</td><td>${esc(h.args.year)} · ${esc(h.args.months.join(", "))}</td><td>${esc(h.changedFiles.map((x) => x.split(/[\\/]/).pop()).join(", "))}${h.removedFiles.length ? " / Kaldırılan: " + esc(h.removedFiles.join(", ")) : ""}</td><td>${esc(h.result?.folder || h.result?.error || "")}</td><td><button data-replay="${esc(h.id)}">Bu girdilerle yeniden üret</button></td></tr>`).join("")}</tbody></table></div></details>`,
  );
}
async function showTrace(department = "", code = null) {
  const r = await api("trace", { state, department: department || null, code });
  showModal(
    "Hesap nasıl oluştu?",
    `<div class="form-grid"><label class="form-field"><span>Bölüm</span><select id="traceDepartment"><option value="">Seçili bölümlerin tümü</option>${data.selection.departments.map((n) => `<option value="${esc(n)}" ${n === department ? "selected" : ""}>${esc(n)}</option>`).join("")}</select></label><label class="form-field"><span>Gider</span><select id="traceCode"><option value="">Tüm giderler</option>${data.detail.groups
      .flatMap((g) => g.items || [])
      .map((g) => `<option value="${g.code}" ${g.code === code ? "selected" : ""}>${esc(g.name)}</option>`)
      .join(
        "",
      )}</select></label></div><button id="traceRunBtn">Hesabı incele</button><p>${esc(r.formula)}</p><p>Gerçekleşen: ${fmt(r.actualTL)} TL · Bütçe: ${fmt(r.budgetTL)} TL · Sapma: ${fmt(r.gapTL)} TL · ${pct(r.rate)}</p>${r.missingDepartments.length ? `<div class="notice">Eksik kaynak: ${esc(r.missingDepartments.join(", "))}</div>` : ""}<details open><summary>Kaynak hücreleri · ${r.rows.length}</summary>${rawTable(["Bölüm", "Müdürlük", "Gider kodu", "Gider", "Tür", "Dönem", "Dosya", "Sayfa", "Hücre", "Değer TL"], r.rows)}</details><details><summary>Birim maliyet hesabı</summary>${rawTable(
      [
        "Bölüm",
        "Yıl",
        "Gider TL",
        "İşlem adedi",
        "BM",
        "Önceki BM",
        "Endeks faktörü",
        "Endekslenmiş BM",
        "İyileşme",
        "Müdürlük ağırlığı",
        "Direktörlük ağırlığı",
        "Karşılaştırılabilir",
      ],
      r.unit.map((i) => [
        i.department,
        i.year,
        i.cost,
        i.count,
        i.actual,
        i.prior,
        i.factor,
        i.adjusted,
        pct(i.rate),
        pct(i.managerWeight),
        pct(i.directorWeight),
        i.eligible ? "Evet" : "Hayır",
      ]),
    )}${r.unit.map((i) => `<details><summary>${esc(i.department)} · ${i.year} endeks oranları</summary>${rawTable(["Ay", "Bu yıl endeks", "Önceki yıl endeks", "Oran"], i.indices)}</details>`).join("")}</details>`,
  );
}
async function showLayout() {
  layoutData = await api("layout-preview", {});
  showModal(
    "Tablo düzeni ve örnek alan",
    `<p>Kaynak: ${esc(layoutData.file)} · Sayfa: ${esc(layoutData.sheet)}</p><label class="form-field"><span>Hücreye tıklayarak doldurulacak parametre</span><select id="layoutTarget">${layoutData.parameters.map((r, i) => `<option value="${i}">${esc(r[0] + " / " + r[1])}</option>`).join("")}</select></label><div class="scroll-table"><table><tbody>${layoutData.sample.map((row) => `<tr>${row.map((c) => `<td><button data-layout-cell="${esc(c[0])}">${esc(c[0] + ": " + c[1])}</button></td>`).join("")}</tr>`).join("")}</tbody></table></div><p>Başlangıç hücreleri ve satırlar mevcut ayar Excel’ine yazılır. Bütçe ve endeks bloklarını bulma kuralları korunur. Kaydetmeden önce kaynakların tamamı okunarak doğrulanır; önceki ayarın yedeği alınır.</p><div class="scroll-table"><table><thead><tr><th>Alan</th><th>Parametre</th><th>Değer</th></tr></thead><tbody>${layoutData.parameters.map((r, i) => `<tr><td>${esc(r[0])}</td><td>${esc(r[1])}</td><td><input data-layout-index="${i}" value="${esc(r[2])}" type="${typeof r[2] === "number" ? "number" : "text"}" step="any"></td></tr>`).join("")}</tbody></table></div><button id="saveLayoutBtn" class="primary">Düzeni doğrula ve kaydet</button>`,
  );
}
function commentVisible() {
  return commentReview.rows
    .map((r, i) => i)
    .filter((i) => {
      const r = commentReview.rows[i],
        text = (
          r.manager +
          " " +
          r.department +
          " " +
          r.item +
          " " +
          r.oldNote +
          " " +
          r.draft
        ).toLocaleLowerCase("tr-TR");
      return (
        (!commentSearch || text.includes(commentSearch.toLocaleLowerCase("tr-TR"))) &&
        (commentFilter === "all" ||
          (commentFilter === "changed" && r.dataChanged) ||
          (commentFilter === "same" && !r.dataChanged) ||
          r.status === commentFilter)
      );
    });
}
function commentState(r) {
  return r.protected
    ? "Bu dönem açıklaması korunuyor"
    : {
        pending: "Kontrol bekliyor",
        approved: "Onaylandı",
        rejected: "Bu dönem geçersiz",
        current: "Bu dönem açıklaması",
      }[r.status] || r.status;
}
async function showComments() {
  commentReview = await api("comment-review", { state });
  commentSelected.clear();
  commentDirty.clear();
  commentOpen = true;
  commentFilter = "pending";
  commentSearch = "";
  showModal("Açıklama inceleme", "");
  renderComments();
}
function renderCommentRows() {
  const visible = commentVisible();
  $("#commentRows").innerHTML =
    visible
      .map((i) => {
        const r = commentReview.rows[i];
        return `<tr><td><input type="checkbox" aria-label="${esc(r.department + " " + r.item)} seç" data-comment-select="${i}" ${commentSelected.has(i) ? "checked" : ""} ${r.protected ? "disabled" : ""}></td><td><strong>${esc(r.department)}</strong><br>${esc(r.manager)}<br>${esc(r.item)} (${r.code})<br><span class="metric-label">${esc(commentState(r))}</span></td><td><div class="old-comment">${esc(r.oldNote || "—")}</div><small>${esc(r.previousClosing)} · ${esc(r.previousMonths)}</small></td><td><textarea data-comment-draft="${i}" maxlength="32000" ${r.protected ? "disabled" : ""} class="${r.status === "pending" ? "pending-note" : ""}" aria-label="${esc(r.department + " " + r.item)} bu dönem açıklaması">${esc(r.draft)}</textarea></td><td>${r.sameMonths ? "Aynı ay kümesi" : "<strong>Ay seçimi farklı</strong>"}<br>${r.dataChanged ? "Tutar / kapsam değişmiş" : "Tutarlar aynı"}<br><small>Gerçekleşen: ${fmt(r.actual === null ? null : r.actual / 1000000)} M TL<br>Bütçe: ${fmt(r.budget === null ? null : r.budget / 1000000)} M TL</small></td></tr>`;
      })
      .join("") ||
    '<tr><td colspan="5">Bu görünümde kayıt yok. Geçmiş dosyalarını seçebilir veya durum filtresini değiştirebilirsiniz.</td></tr>';
  $("#commentSelectedCount").textContent =
    commentSelected.size + " satır seçili · " + visible.length + " görünür";
}
function renderComments() {
  const counts = Object.fromEntries(
    ["pending", "approved", "rejected", "current"].map((s) => [
      s,
      commentReview.rows.filter((r) => r.status === s).length,
    ]),
  );
  $("#modalBody").innerHTML =
    `<p>${commentReview.closing.join("-")} kapanışı · Seçili aylar: ${commentReview.months.join(", ")} · <strong>${counts.pending} bekleyen, ${counts.approved} onaylı, ${counts.rejected} geçersiz, ${counts.current} mevcut açıklama</strong></p><p class="control-note">Önceki yorumlar öneridir. Tutarların aynı olması açıklamayı otomatik doğrulamaz. Bu dönemki mevcut açıklamaların üzerine yazılmaz. İşlem yalnız görünür ve seçili satırlara uygulanır.</p><div class="comment-tools"><button id="commentFilesBtn">Önceki ayın açıklamalı dosyalarını seç</button><button id="commentSelectBtn">Görünürleri seç</button><button id="commentClearBtn">Seçimi temizle</button><button id="commentApproveBtn" class="primary">Seçileni onayla</button><button id="commentRejectBtn">Seçileni bu dönem geçersiz say</button><button id="commentSaveBtn">Düzenlenen taslakları kaydet</button></div><div class="comment-tools"><label>Durum <select id="commentFilter">${[
      ["pending", "Kontrol bekleyenler"],
      ["approved", "Onaylananlar"],
      ["rejected", "Bu dönem geçersiz"],
      ["changed", "Verisi / kapsamı değişenler"],
      ["same", "Verisi aynı kalanlar"],
      ["all", "Tümü"],
    ]
      .map(([v, t]) => `<option value="${v}" ${v === commentFilter ? "selected" : ""}>${t}</option>`)
      .join(
        "",
      )}</select></label><input id="commentSearch" value="${esc(commentSearch)}" placeholder="Bölüm, gider veya metin ara"><span id="commentSelectedCount"></span></div><details><summary>Seçili satırlara ortak açıklama uygula</summary><textarea id="commentCommonText" maxlength="32000" placeholder="Aynı açıklamanın geçerli olduğu satırları seçin. Metin taslağa uygulanır; onay ayrı yapılır."></textarea><button id="commentSameBtn">Metni seçili taslaklara uygula</button></details><div class="scroll-table comment-table"><table><thead><tr><th>Seç</th><th>Bölüm / alt gider</th><th>Önceki yorum · kaynak dönem</th><th>Bu dönem taslağı · doğrudan düzenle</th><th>Veri kontrolü</th></tr></thead><tbody id="commentRows"></tbody></table></div><p class="control-note">Kırmızı taslaklar rapora “Kontrol bekliyor — önceki dönemden” etiketiyle gider. Onaylananlar normal renge döner; geçersiz sayılanlar rapora alınmaz. Mevcut bu dönem açıklamasını Sapma öncelikleri ekranından düzenleyebilirsiniz.</p>`;
  renderCommentRows();
}
async function saveCommentBatch(action, dirtyOnly = false) {
  const indices = dirtyOnly
    ? [...commentDirty].filter(
        (i) => commentReview.rows[i].status === "pending" && !commentReview.rows[i].protected,
      )
    : commentVisible().filter((i) => commentSelected.has(i) && !commentReview.rows[i].protected);
  if (!indices.length) throw Error("İşlem yapılacak bir satır seçin veya taslak düzenleyin.");
  // A changed draft elsewhere must not be lost when a selected subset is committed.
  const remaining = [...commentDirty].filter((i) => !indices.includes(i));
  const entries = indices.map((i) => {
    const r = commentReview.rows[i];
    return { department: r.department, code: r.code, token: r.token, draft: r.draft, action };
  });
  await api("comment-apply", { state, entries });
  const old = commentReview;
  commentReview = await api("comment-review", { state });
  commentDirty.clear();
  for (const i of remaining) {
    const r = old.rows[i],
      j = commentReview.rows.findIndex((x) => x.department === r.department && x.code === r.code);
    if (j >= 0 && !commentReview.rows[j].protected) {
      commentReview.rows[j].draft = r.draft;
      commentDirty.add(j);
    }
  }
  commentSelected.clear();
  renderComments();
  await refresh();
  toast(entries.length + " açıklama işlendi.");
}
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && commentOpen && commentDirty.size) {
    if (!window.confirm("Kaydedilmemiş taslaklar var. Kaydetmeden çıkılsın mı?")) e.preventDefault();
    else {
      commentDirty.clear();
      commentOpen = false;
    }
  }
});
async function init() {
  try {
    meta = await api("meta");
    defaults();
    if (meta.savedFilters) {
      const saved = meta.savedFilters;
      try {
        await api("query", saved);
        state = saved;
      } catch {
        defaults();
      }
    }
    buildFilters();
    await refresh();
  } catch (e) {
    $("#content").innerHTML = empty(e.message);
    toast(e.message);
  }
}
init();
