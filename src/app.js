import Chart from "chart.js/auto";
import L from "leaflet";
import * as XLSX from "xlsx";

    // ─── API endpoints ────────────────────────────────────────
    const API_AIMAG =
      "https://data.1212.mn/api/v1/mn/NSO/Economy, environment/Consumer Price Index/DT_NSO_0300_010V5.px";
    const API_UB =
      "https://data.1212.mn/api/v1/mn/NSO/Economy, environment/Consumer Price Index/DT_NSO_0600_001V4.px";

    // 21 aimags + capital coords (skip region aggregates 1,2,3,4)
    const AIMAGS = {
      "183": { name: "Баян-Өлгий", lat: 48.9683, lng: 89.9687 },
      "182": { name: "Говь-Алтай", lat: 46.3722, lng: 96.2572 },
      "181": { name: "Завхан", lat: 47.7417, lng: 96.8444 },
      "185": { name: "Увс", lat: 49.9811, lng: 92.0667 },
      "184": { name: "Ховд", lat: 48.0056, lng: 91.6419 },
      "265": { name: "Архангай", lat: 47.4769, lng: 101.4503 },
      "264": { name: "Баянхонгор", lat: 46.1944, lng: 100.7181 },
      "263": { name: "Булган", lat: 48.8125, lng: 103.5347 },
      "261": { name: "Орхон", lat: 49.0272, lng: 104.0444 },
      "262": { name: "Өвөрхангай", lat: 46.2667, lng: 102.7778 },
      "267": { name: "Хөвсгөл", lat: 49.6342, lng: 100.1625 },
      "342": { name: "Говьсүмбэр", lat: 46.5269, lng: 108.3411 },
      "345": { name: "Дархан-Уул", lat: 49.4867, lng: 105.9228 },
      "344": { name: "Дорноговь", lat: 44.8917, lng: 110.1367 },
      "348": { name: "Дундговь", lat: 45.7625, lng: 106.2708 },
      "346": { name: "Өмнөговь", lat: 43.5708, lng: 104.4250 },
      "343": { name: "Сэлэнгэ", lat: 50.2364, lng: 106.2056 },
      "341": { name: "Төв", lat: 47.7069, lng: 106.9528 },
      "421": { name: "Дорнод", lat: 48.0725, lng: 114.5328 },
      "422": { name: "Сүхбаатар", lat: 46.6806, lng: 113.2792 },
      "423": { name: "Хэнтий", lat: 47.3194, lng: 110.6556 },
      "511": { name: "Улаанбаатар", lat: 47.9186, lng: 106.9176 },
    };
    const AIMAG_CODES = Object.keys(AIMAGS).filter((c) => c !== "511");
    const aimagName = (code) => AIMAGS[code]?.name || code;

    /** Shared product mapping: aimag code <-> ub code (approximate name match) */
    const SHARED = [
      { key: "sheep", label: "Хонины мах, ястай, кг", aimag: "7", ub: "9" },
      { key: "beef", label: "Үхрийн мах, ястай, кг", aimag: "6", ub: "10" },
      { key: "goat", label: "Ямааны мах, ястай, кг", aimag: "10", ub: "13" },
      { key: "flour1", label: "Гурил, I зэрэг, савласан, кг", aimag: "0", ub: "3" },
      { key: "sugar", label: "Элсэн чихэр, кг", aimag: "9", ub: "20" },
      { key: "rice", label: "Цагаан будаа, кг", aimag: "8", ub: "8" },
      { key: "milk", label: "Сүү (задгай)", aimag: "5", ub: "14" },
      { key: "ai80", label: "Аи-80 автобензин, л", aimag: "1", ub: "29" },
      { key: "ai92", label: "Аи-92 автобензин, л", aimag: "2", ub: "30" },
      { key: "diesel", label: "Дизелийн түлш, л", aimag: "4", ub: "31" },
    ];

    const CHART_COLORS = [
      "#38bdf8","#818cf8","#f472b6","#4ade80","#fbbf24","#fb7185",
      "#2dd4bf","#a78bfa","#f97316","#34d399","#60a5fa","#e879f9",
    ];

    // ─── State ────────────────────────────────────────────────
    let metaAimag = null;
    let metaUb = null;
    /** aimagData[productCode][aimagCode][weekCode] = number|null */
    let aimagData = {};
    /** ubData[productCode][weekCode] = number|null */
    let ubData = {};
    let aimagProducts = []; // {code,name}
    let ubProducts = [];
    let aimagWeeks = []; // {code,label} newest first
    let ubWeeks = [];
    let charts = {};
    let mapMode = "pct"; // pct | price
    const mapInstances = {}; // key -> { map, layer }
    let mapMarkersByCode = {}; // for highlight

    // ─── DOM ──────────────────────────────────────────────────
    const $ = (id) => document.getElementById(id);
    const elOverlay = $("overlay");
    const elLoadMsg = $("loadMsg");
    const elError = $("errorBox");
    const elUpdated = $("lastUpdated");

    // ─── Utils ────────────────────────────────────────────────
    function showLoading(msg) {
      elLoadMsg.textContent = msg || "Өгөгдөл ачаалж байна…";
      elOverlay.classList.remove("hidden");
    }
    function hideLoading() { elOverlay.classList.add("hidden"); }
    function showError(msg) {
      elError.textContent = msg;
      elError.classList.add("show");
    }
    function clearError() {
      elError.classList.remove("show");
      elError.textContent = "";
    }
    function fmtMNT(n) {
      if (n == null || Number.isNaN(n)) return "—";
      return new Intl.NumberFormat("mn-MN").format(Math.round(n)) + " ₮";
    }
    function fmtNum(n, d = 0) {
      if (n == null || Number.isNaN(n)) return "—";
      return new Intl.NumberFormat("mn-MN", { maximumFractionDigits: d, minimumFractionDigits: d }).format(n);
    }
    function fmtPct(n) {
      if (n == null || Number.isNaN(n)) return "—";
      const s = n > 0 ? "+" : "";
      return s + n.toFixed(1) + "%";
    }
    function chgClass(pct) {
      if (pct == null || Number.isNaN(pct)) return "flat";
      if (pct > 0.05) return "up";
      if (pct < -0.05) return "down";
      return "flat";
    }
    function pctChange(a, b) {
      if (a == null || b == null || b === 0 || Number.isNaN(a) || Number.isNaN(b)) return null;
      return ((a - b) / b) * 100;
    }
    function mean(arr) {
      const v = arr.filter((x) => x != null && !Number.isNaN(x));
      if (!v.length) return null;
      return v.reduce((s, x) => s + x, 0) / v.length;
    }
    function destroyChart(key) {
      if (charts[key]) { charts[key].destroy(); delete charts[key]; }
    }
    function chartDefaults() {
      Chart.defaults.color = "#9aa3b2";
      Chart.defaults.borderColor = "rgba(255,255,255,.08)";
      Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
    }

    // ─── Map (Leaflet) ────────────────────────────────────────
    function isMobileView() {
      return window.matchMedia && window.matchMedia("(max-width: 640px)").matches;
    }
    function isTabletView() {
      return window.matchMedia && window.matchMedia("(max-width: 1024px)").matches;
    }
    function defaultMapZoom() {
      if (isMobileView()) return 4.4;
      if (isTabletView()) return 4.85;
      return 5.15;
    }

    function ensureMap(key) {
      if (mapInstances[key]) return mapInstances[key];
      const elId = key === "main" ? "priceMap" : "aimagMap";
      const el = $(elId);
      if (!el || typeof L === "undefined") return null;

      const map = L.map(elId, {
        center: [46.8, 103.8],
        zoom: defaultMapZoom(),
        zoomSnap: 0.25,
        minZoom: 3.5,
        maxZoom: 10,
        tapTolerance: 15,
      });
      L.tileLayer("https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png", {
        attribution:
          '&copy; <a href="https://www.openstreetmap.org/copyright">OSM</a> &copy; <a href="https://carto.com/">CARTO</a> · <a href="https://www.1212.mn">ҮСХ</a>',
        subdomains: "abcd",
        maxZoom: 19,
      }).addTo(map);
      const layer = L.layerGroup().addTo(map);
      mapInstances[key] = { map, layer };
      // Fix tiles when container was hidden
      setTimeout(() => map.invalidateSize(), 80);
      return mapInstances[key];
    }

    function invalidateMaps() {
      Object.values(mapInstances).forEach((m) => {
        try { m.map.invalidateSize({ animate: false }); } catch (_) { /* ignore */ }
      });
    }

    let resizeTimer = null;
    window.addEventListener("resize", () => {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(() => {
        invalidateMaps();
        // Re-fit charts when container width changes
        Object.values(charts).forEach((c) => {
          try { c.resize(); } catch (_) { /* ignore */ }
        });
      }, 150);
    });

    function colorForPct(pct) {
      if (pct == null || Number.isNaN(pct)) return "#475569";
      if (pct > 0.05) return "#ef4444";
      if (pct < -0.05) return "#22c55e";
      return "#94a3b8";
    }

    function colorForPriceLevel(price, minP, maxP) {
      if (price == null || Number.isNaN(price) || maxP === minP) return "#38bdf8";
      const t = (price - minP) / (maxP - minP);
      // blue (cheap) -> amber -> red (expensive)
      const r = Math.round(56 + t * (239 - 56));
      const g = Math.round(189 - t * (189 - 68));
      const b = Math.round(248 - t * (248 - 68));
      return `rgb(${r},${g},${b})`;
    }

    function radiusForPrice(price, minP, maxP) {
      if (price == null || Number.isNaN(price)) return 8;
      if (maxP === minP) return 14;
      const t = (price - minP) / (maxP - minP);
      return 8 + t * 20; // 8–28
    }

    function setMapLegend(legendEl, mode) {
      if (!legendEl) return;
      if (mode === "price") {
        legendEl.innerHTML = `
          <strong>Үнийн түвшин</strong>
          <div class="row"><span class="dot" style="background:#38bdf8"></span> Хямд</div>
          <div class="row"><span class="dot" style="background:#fbbf24"></span> Дунд</div>
          <div class="row"><span class="dot" style="background:#ef4444"></span> Үнэтэй</div>
          <div style="margin-top:0.35rem">Тойргийн хэмжээ = үнэ</div>`;
      } else {
        legendEl.innerHTML = `
          <strong>7 хоногийн өөрчлөлт</strong>
          <div class="row"><span class="dot" style="background:#ef4444"></span> Өссөн</div>
          <div class="row"><span class="dot" style="background:#22c55e"></span> Буурсан</div>
          <div class="row"><span class="dot" style="background:#94a3b8"></span> Өөрчлөлтгүй</div>
          <div class="row"><span class="dot" style="background:#475569"></span> Мэдээлэлгүй</div>
          <div style="margin-top:0.35rem">Тойргийн хэмжээ = үнэ</div>`;
      }
    }

    /**
     * rows: { code, name, priceA, priceB, pct }[]
     * opts: productName, labelA, labelB, mode, legendEl, onSelect
     */
    function paintMap(key, rows, opts) {
      const inst = ensureMap(key);
      if (!inst) return;
      const { map, layer } = inst;
      const mode = opts.mode || "pct";
      layer.clearLayers();
      if (key === "main") mapMarkersByCode = {};

      const prices = rows.map((r) => r.priceA).filter((v) => v != null && !Number.isNaN(v));
      const minP = prices.length ? Math.min(...prices) : 0;
      const maxP = prices.length ? Math.max(...prices) : 1;

      setMapLegend(opts.legendEl, mode);

      rows.forEach((row) => {
        const geo = AIMAGS[row.code];
        if (!geo) return;
        const color =
          mode === "price"
            ? colorForPriceLevel(row.priceA, minP, maxP)
            : colorForPct(row.pct);
        const r = radiusForPrice(row.priceA, minP, maxP);
        const isUb = row.code === "511";
        // Slightly smaller markers on phones so map doesn't feel crowded
        const scale = isMobileView() ? 0.78 : isTabletView() ? 0.9 : 1;
        const circle = L.circleMarker([geo.lat, geo.lng], {
          radius: Math.max(5, r * scale),
          color: isUb ? "#fbbf24" : "#fff",
          weight: isUb ? 2.5 : 1.5,
          fillColor: color,
          fillOpacity: 0.88,
        });
        const cls = chgClass(row.pct);
        const html = `
          <div class="price-popup">
            <strong>${escapeHtml(row.name)}${isUb ? " ★" : ""}</strong>
            <div class="p-name">${escapeHtml(opts.productName || "")}</div>
            <div class="p-big">${fmtMNT(row.priceA)}</div>
            <div class="p-row"><span>${escapeHtml(opts.labelA || "A")}</span><span>${fmtMNT(row.priceA)}</span></div>
            <div class="p-row"><span>${escapeHtml(opts.labelB || "B")}</span><span>${fmtMNT(row.priceB)}</span></div>
            <div class="p-row"><span>Өөрчлөлт</span><span class="chg ${cls}">${fmtPct(row.pct)}</span></div>
          </div>`;
        circle.bindPopup(html);
        circle.on("click", () => {
          if (typeof opts.onSelect === "function") opts.onSelect(row.code);
        });
        circle.addTo(layer);
        if (key === "main") mapMarkersByCode[row.code] = circle;
      });

      setTimeout(() => map.invalidateSize(), 60);
    }

    function buildMapRows(productCode, weekA, weekB, includeUb) {
      const rows = AIMAG_CODES.map((c) => {
        const priceA = aimagData[productCode]?.[c]?.[weekA] ?? null;
        const priceB = aimagData[productCode]?.[c]?.[weekB] ?? null;
        return {
          code: c,
          name: aimagName(c),
          priceA,
          priceB,
          pct: pctChange(priceA, priceB),
        };
      });

      if (includeUb) {
        const sh = SHARED.find((s) => s.aimag === productCode);
        if (sh) {
          const labelA = aimagWeeks.find((w) => w.code === weekA)?.name;
          const labelB = aimagWeeks.find((w) => w.code === weekB)?.name;
          const uA = ubWeeks.find((w) => w.name === labelA)?.code || ubWeeks[0]?.code;
          const uB = ubWeeks.find((w) => w.name === labelB)?.code || ubWeeks[1]?.code;
          const priceA = ubData[sh.ub]?.[uA] ?? null;
          const priceB = ubData[sh.ub]?.[uB] ?? null;
          rows.push({
            code: "511",
            name: "Улаанбаатар",
            priceA,
            priceB,
            pct: pctChange(priceA, priceB),
          });
        }
      }
      return rows;
    }

    function highlightMapTableRow(code) {
      const tb = $("mapTable")?.querySelector("tbody");
      if (!tb) return;
      tb.querySelectorAll("tr").forEach((tr) => {
        tr.classList.toggle("hl", tr.dataset.code === code);
      });
      const tr = tb.querySelector(`tr[data-code="${code}"]`);
      if (tr) tr.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }

    function renderMapPanel() {
      const pCode = $("mapProduct").value || aimagProducts[0]?.code;
      const wA = $("mapWeekA").value || aimagWeeks[0]?.code;
      const wB = $("mapWeekB").value || aimagWeeks[1]?.code;
      const pName = aimagProducts.find((p) => p.code === pCode)?.name || "";
      const labelA = aimagWeeks.find((w) => w.code === wA)?.name || "A";
      const labelB = aimagWeeks.find((w) => w.code === wB)?.name || "B";
      const canUb = SHARED.some((s) => s.aimag === pCode);

      $("mapTitle").textContent = pName + " · " + labelA;
      $("mapNote").innerHTML = canUb
        ? `Тойргийн <strong>хэмжээ</strong> = үнийн түвшин · <strong>өнгө</strong> = ${
            mapMode === "price" ? "үнийн түвшин (хөх→улаан)" : "7 хоногийн өөрчлөлт (улаан өссөн, ногоон буурсан)"
          }. <strong>Улаанбаатар</strong> (алтан хүрээ) нийтлэг бүтээгдэхүүн дээр харагдана.`
        : `Тойргийн <strong>хэмжээ</strong> = үнийн түвшин · <strong>өнгө</strong> = ${
            mapMode === "price" ? "үнийн түвшин" : "7 хоногийн өөрчлөлт"
          }. Энэ бүтээгдэхүүн зөвхөн аймгийн хүснэгтэд байгаа тул УБ зурагт орохгүй.`;

      const rows = buildMapRows(pCode, wA, wB, true);
      const valid = rows.filter((r) => r.priceA != null);
      const avg = mean(valid.filter((r) => r.code !== "511").map((r) => r.priceA));
      const avgPct = mean(valid.filter((r) => r.code !== "511").map((r) => r.pct));
      const maxR = [...valid].sort((a, b) => (b.priceA || 0) - (a.priceA || 0))[0];
      const minR = [...valid].sort((a, b) => (a.priceA || 0) - (b.priceA || 0))[0];
      const ubRow = rows.find((r) => r.code === "511");

      $("mapKpis").innerHTML = `
        <div class="kpi"><div class="label">Аймаг дундаж</div><div class="value">${fmtMNT(avg)}</div><div class="sub">${labelA}</div></div>
        <div class="kpi"><div class="label">Хамгийн үнэтэй</div><div class="value" style="font-size:1rem">${maxR ? escapeHtml(maxR.name) : "—"}</div><div class="sub">${maxR ? fmtMNT(maxR.priceA) : ""}</div></div>
        <div class="kpi"><div class="label">Хамгийн хямд</div><div class="value" style="font-size:1rem">${minR ? escapeHtml(minR.name) : "—"}</div><div class="sub">${minR ? fmtMNT(minR.priceA) : ""}</div></div>
        <div class="kpi"><div class="label">Дундаж өөрчлөлт</div><div class="value chg ${chgClass(avgPct)}">${fmtPct(avgPct)}</div><div class="sub">A vs B</div></div>
        <div class="kpi"><div class="label">Улаанбаатар</div><div class="value" style="font-size:1rem">${ubRow && ubRow.priceA != null ? fmtMNT(ubRow.priceA) : "—"}</div><div class="sub chg ${chgClass(ubRow?.pct)}">${ubRow && ubRow.priceA != null ? fmtPct(ubRow.pct) : canUb ? "" : "нийтлэг биш"}</div></div>
      `;

      ensureMap("main");
      paintMap("main", rows, {
        productName: pName,
        labelA,
        labelB,
        mode: mapMode,
        legendEl: $("mapLegend"),
        onSelect: (code) => highlightMapTableRow(code),
      });

      const sorted = [...rows].filter((r) => r.priceA != null).sort((a, b) => b.priceA - a.priceA);
      const tb = $("mapTable").querySelector("tbody");
      tb.innerHTML = sorted.map((r, i) => `
        <tr data-code="${r.code}" style="cursor:pointer">
          <td class="rank">${i + 1}</td>
          <td>${escapeHtml(r.name)}${r.code === "511" ? " ★" : ""}</td>
          <td class="num">${fmtMNT(r.priceA)}</td>
          <td class="num">${fmtMNT(r.priceB)}</td>
          <td class="num chg ${chgClass(r.pct)}">${fmtPct(r.pct)}</td>
        </tr>
      `).join("") || `<tr><td colspan="5" class="empty">Мэдээлэл алга</td></tr>`;

      tb.querySelectorAll("tr[data-code]").forEach((tr) => {
        tr.addEventListener("click", () => {
          const code = tr.dataset.code;
          highlightMapTableRow(code);
          const marker = mapMarkersByCode[code];
          if (marker) {
            marker.openPopup();
            const inst = mapInstances.main;
            if (inst) inst.map.panTo(marker.getLatLng(), { animate: true });
          }
        });
      });
    }

    // ─── API ──────────────────────────────────────────────────
    async function fetchMeta(url) {
      const res = await fetch(url);
      if (!res.ok) throw new Error("Метадата алдаа: " + res.status + " · " + url);
      return res.json();
    }

    async function postQuery(url, query) {
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ query, response: { format: "json-stat2" } }),
      });
      if (!res.ok) throw new Error("Өгөгдөл татах алдаа: " + res.status);
      return res.json();
    }

    function parseMetaVariables(meta) {
      const vars = {};
      for (const v of meta.variables || []) {
        const items = [];
        for (let i = 0; i < v.values.length; i++) {
          const code = String(v.values[i] ?? "");
          if (code === "" && !v.valueTexts[i]) continue;
          items.push({ code, name: String(v.valueTexts[i] || code).trim() });
        }
        vars[v.code] = items;
        vars[v.text] = items; // alias by mongolian label too
      }
      return vars;
    }

    /**
     * json-stat2: dimensions order in js.id, values flat array
     * Index formula: for dims d0,d1,d2 with sizes s0,s1,s2:
     * idx = i0*s1*s2 + i1*s2 + i2
     */
    function parseJsonStat(js) {
      const ids = js.id || [];
      const size = js.size || [];
      const values = js.value || [];
      const dim = js.dimension || {};

      const dimCodes = {}; // dimName -> ordered codes by position
      const dimLabels = {};
      for (const id of ids) {
        const cat = dim[id]?.category;
        if (!cat) continue;
        const index = cat.index; // code -> position
        const label = cat.label || {};
        const ordered = new Array(Object.keys(index).length);
        for (const [code, pos] of Object.entries(index)) {
          ordered[pos] = String(code);
        }
        dimCodes[id] = ordered;
        dimLabels[id] = label;
      }

      const n = values.length;
      const strides = new Array(ids.length);
      let acc = 1;
      for (let d = ids.length - 1; d >= 0; d--) {
        strides[d] = acc;
        acc *= size[d];
      }

      function getCoords(flatIdx) {
        const coords = {};
        for (let d = 0; d < ids.length; d++) {
          const pos = Math.floor(flatIdx / strides[d]) % size[d];
          coords[ids[d]] = dimCodes[ids[d]][pos];
        }
        return coords;
      }

      const rows = [];
      for (let i = 0; i < n; i++) {
        const val = values[i];
        const num = val == null || val === "" ? null : Number(val);
        rows.push({ ...getCoords(i), value: Number.isFinite(num) ? num : null });
      }
      return { rows, dimCodes, dimLabels, updated: js.updated, label: js.label };
    }

    async function loadAllData() {
      showLoading("Метадата татаж байна…");
      metaAimag = await fetchMeta(API_AIMAG);
      metaUb = await fetchMeta(API_UB);

      const va = parseMetaVariables(metaAimag);
      const vu = parseMetaVariables(metaUb);

      aimagProducts = (va["Бүтээгдэхүүн"] || []).filter((p) => p.code !== "");
      ubProducts = (vu["Бүтээгдэхүүн"] || []).filter((p) => p.code !== "");
      aimagWeeks = (va["Хугацаа"] || []).slice(); // already newest first in meta
      ubWeeks = (vu["Хугацаа"] || []).slice();

      // Fetch recent history: enough for charts (up to 52 weeks + all products)
      const weekCountA = Math.min(aimagWeeks.length, 60);
      const weekCountU = Math.min(ubWeeks.length, 60);
      const weekCodesA = aimagWeeks.slice(0, weekCountA).map((w) => w.code);
      const weekCodesU = ubWeeks.slice(0, weekCountU).map((w) => w.code);
      const prodCodesA = aimagProducts.map((p) => p.code);
      const prodCodesU = ubProducts.map((p) => p.code);

      showLoading("Аймгуудын үнэ татаж байна…");
      const jsA = await postQuery(API_AIMAG, [
        { code: "Бүтээгдэхүүн", selection: { filter: "item", values: prodCodesA } },
        { code: "Бүс", selection: { filter: "item", values: AIMAG_CODES } },
        { code: "Хугацаа", selection: { filter: "item", values: weekCodesA } },
      ]);

      showLoading("Улаанбаатарын үнэ татаж байна…");
      const jsU = await postQuery(API_UB, [
        { code: "Бүтээгдэхүүн", selection: { filter: "item", values: prodCodesU } },
        { code: "Хугацаа", selection: { filter: "item", values: weekCodesU } },
      ]);

      // Build nested maps
      aimagData = {};
      const parsedA = parseJsonStat(jsA);
      for (const r of parsedA.rows) {
        const p = r["Бүтээгдэхүүн"];
        const a = r["Бүс"];
        const t = r["Хугацаа"];
        if (!aimagData[p]) aimagData[p] = {};
        if (!aimagData[p][a]) aimagData[p][a] = {};
        aimagData[p][a][t] = r.value;
      }

      ubData = {};
      const parsedU = parseJsonStat(jsU);
      for (const r of parsedU.rows) {
        const p = r["Бүтээгдэхүүн"];
        const t = r["Хугацаа"];
        if (!ubData[p]) ubData[p] = {};
        ubData[p][t] = r.value;
      }

      const uDate = ubWeeks[0]?.name || "—";
      const aDate = aimagWeeks[0]?.name || "—";
      elUpdated.textContent =
        "Сүүлийн өгөгдөл: " + uDate + (uDate !== aDate ? " (УБ) / " + aDate + " (аймаг)" : "");
      return { parsedA, parsedU };
    }

    // ─── Populate selects ─────────────────────────────────────
    function fillSelect(el, items, mapFn) {
      el.innerHTML = "";
      for (const it of items) {
        const opt = document.createElement("option");
        const m = mapFn(it);
        opt.value = m.value;
        opt.textContent = m.label;
        el.appendChild(opt);
      }
    }

    function setupControls() {
      fillSelect($("ubProduct"), ubProducts, (p) => ({ value: p.code, label: p.name }));
      fillSelect($("aimagProduct"), aimagProducts, (p) => ({ value: p.code, label: p.name }));
      fillSelect($("mapProduct"), aimagProducts, (p) => ({ value: p.code, label: p.name }));
      fillSelect($("ovProduct"), SHARED, (p) => ({ value: p.key, label: p.label }));
      fillSelect($("cmpProduct"), SHARED, (p) => ({ value: p.key, label: p.label }));
      fillSelect($("aimagWeekA"), aimagWeeks, (w) => ({ value: w.code, label: w.name }));
      fillSelect($("aimagWeekB"), aimagWeeks, (w) => ({ value: w.code, label: w.name }));
      fillSelect($("mapWeekA"), aimagWeeks, (w) => ({ value: w.code, label: w.name }));
      fillSelect($("mapWeekB"), aimagWeeks, (w) => ({ value: w.code, label: w.name }));
      fillSelect($("cmpWeek"), aimagWeeks.slice(0, 40), (w) => ({ value: w.code, label: w.name }));
      fillSelect($("trAimag"), AIMAG_CODES.map((c) => ({ code: c, name: aimagName(c) })), (a) => ({
        value: a.code, label: a.name,
      }));

      if (aimagWeeks.length > 1) {
        $("aimagWeekA").value = aimagWeeks[0].code;
        $("aimagWeekB").value = aimagWeeks[1].code;
        $("mapWeekA").value = aimagWeeks[0].code;
        $("mapWeekB").value = aimagWeeks[1].code;
      }

      // UB chips — quick categories (detail view)
      const chips = [
        { q: "", label: "Бүгд" },
        { q: "мах", label: "Мах" },
        { q: "гурил", label: "Гурил" },
        { q: "талх", label: "Талх" },
        { q: "сүү|тараг|цөцгийн", label: "Сүү/цагаан идээ" },
        { q: "төмс|лууван|байцаа|сонгино|манжин|алим", label: "Ногоо/жимс" },
        { q: "будаа|чихэр|тос|өндөг|цай", label: "Үндсэн" },
        { q: "бензин|дизель|аи-", label: "Шатахуун" },
      ];
      const chipBox = $("ubChips");
      chipBox.innerHTML = "";
      chips.forEach((c, i) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "chip" + (i === 0 ? " active" : "");
        b.textContent = c.label;
        b.dataset.q = c.q;
        b.addEventListener("click", () => {
          chipBox.querySelectorAll(".chip").forEach((x) => x.classList.remove("active"));
          b.classList.add("active");
          filterUbProductSelect(c.q);
          if (ubActiveSub === "detail") renderUbDetail();
        });
        chipBox.appendChild(b);
      });

      // Multi chips (same categories → select matching products)
      const multiChips = $("ubMultiChips");
      if (multiChips) {
        multiChips.innerHTML = "";
        chips.forEach((c) => {
          const b = document.createElement("button");
          b.type = "button";
          b.className = "chip";
          b.textContent = c.label || "Бүгд";
          b.addEventListener("click", () => {
            const sel = $("ubMultiProducts");
            const re = c.q ? new RegExp(c.q, "i") : null;
            [...sel.options].forEach((opt) => {
              opt.selected = !re || re.test(opt.textContent);
            });
            // limit to 8 for readability
            let n = 0;
            [...sel.options].forEach((opt) => {
              if (opt.selected) {
                n++;
                if (n > 8) opt.selected = false;
              }
            });
            renderUbMulti();
          });
          multiChips.appendChild(b);
        });
      }

      // Table category select
      const tableCat = $("ubTableCat");
      if (tableCat) {
        tableCat.innerHTML = UB_CATEGORIES.map(
          (c) => `<option value="${c.id === "all" ? "" : c.id}">${c.label}</option>`
        ).join("");
      }

      // Multi product select
      fillSelect($("ubMultiProducts"), ubProducts, (p) => ({ value: p.code, label: p.name }));
      [...$("ubMultiProducts").options].slice(0, 4).forEach((o) => { o.selected = true; });

      // Trend multi-select defaults: first 3 ub products
      fillTrProducts("ub");
      const trSel = $("trProducts");
      [...trSel.options].slice(0, 3).forEach((o) => { o.selected = true; });
    }

    function filterUbProductSelect(q) {
      const sel = $("ubProduct");
      const re = q ? new RegExp(q, "i") : null;
      let first = null;
      [...sel.options].forEach((opt) => {
        const show = !re || re.test(opt.textContent);
        opt.hidden = !show;
        if (show && !first) first = opt;
      });
      if (first && re) {
        sel.value = first.value;
      }
    }

    function fillTrProducts(source) {
      const sel = $("trProducts");
      const list = source === "ub" ? ubProducts : aimagProducts;
      fillSelect(sel, list, (p) => ({ value: p.code, label: p.name }));
    }

    // ─── Tabs ─────────────────────────────────────────────────
    function activateTab(name, opts = {}) {
      const tab = document.querySelector('.tab[data-tab="' + name + '"]');
      const panel = $("panel-" + name);
      if (!tab || !panel) return;
      document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
      tab.classList.add("active");
      panel.classList.add("active");
      if (!opts.skipHash) {
        const hash = name === "overview" ? "" : "#" + name;
        if (location.hash !== hash) history.replaceState(null, "", hash || location.pathname + location.search);
      }
      if (name === "overview") renderOverview();
      if (name === "map") {
        renderMapPanel();
        setTimeout(invalidateMaps, 100);
      }
      if (name === "ub") renderUb();
      if (name === "aimag") {
        renderAimag();
        setTimeout(invalidateMaps, 100);
      }
      if (name === "compare") renderCompare();
      if (name === "trend") renderTrend();
    }

    function setupTabs() {
      document.querySelectorAll(".tab").forEach((tab) => {
        tab.addEventListener("click", () => {
          activateTab(tab.dataset.tab);
        });
      });
    }

    // ─── Overview ─────────────────────────────────────────────
    function renderOverview() {
      const w0 = ubWeeks[0]?.code;
      const w1 = ubWeeks[1]?.code;
      const a0 = aimagWeeks[0]?.code;
      const a1 = aimagWeeks[1]?.code;

      // KPI: count products, avg weekly change UB, biggest movers
      const ubRows = ubProducts.map((p) => {
        const cur = ubData[p.code]?.[w0] ?? null;
        const prev = ubData[p.code]?.[w1] ?? null;
        return { ...p, cur, prev, pct: pctChange(cur, prev) };
      }).filter((r) => r.cur != null);

      const up = ubRows.filter((r) => r.pct != null && r.pct > 0.05).length;
      const down = ubRows.filter((r) => r.pct != null && r.pct < -0.05).length;
      const avgPct = mean(ubRows.map((r) => r.pct));
      const topUp = [...ubRows].filter((r) => r.pct != null).sort((a, b) => b.pct - a.pct)[0];
      const topDown = [...ubRows].filter((r) => r.pct != null).sort((a, b) => a.pct - b.pct)[0];

      // Aimag national mean for sheep meat as sample
      const sheep = SHARED[0];
      const sheepPrices = AIMAG_CODES.map((c) => aimagData[sheep.aimag]?.[c]?.[a0] ?? null);
      const sheepMean = mean(sheepPrices);
      const sheepUb = ubData[sheep.ub]?.[w0] ?? null;

      $("overviewKpis").innerHTML = `
        <div class="kpi"><div class="label">УБ бүтээгдэхүүн</div><div class="value">${ubProducts.length}</div><div class="sub">нийслэлийн жагсаалт</div></div>
        <div class="kpi"><div class="label">Аймгийн бүтээгдэхүүн</div><div class="value">${aimagProducts.length}</div><div class="sub">21 аймаг</div></div>
        <div class="kpi"><div class="label">Сүүлийн 7 хоног</div><div class="value" style="font-size:1.05rem">${ubWeeks[0]?.name || "—"}</div><div class="sub">шинэчлэгдсэн өдөр</div></div>
        <div class="kpi"><div class="label">УБ дундаж өөрчлөлт</div><div class="value chg ${chgClass(avgPct)}">${fmtPct(avgPct)}</div><div class="sub">өсөлт ${up} · бууралт ${down}</div></div>
        <div class="kpi"><div class="label">Хамгийн их өссөн (УБ)</div><div class="value" style="font-size:.95rem">${topUp ? topUp.name.split(",")[0] : "—"}</div><div class="sub chg up">${topUp ? fmtPct(topUp.pct) + " · " + fmtMNT(topUp.cur) : ""}</div></div>
        <div class="kpi"><div class="label">Хамгийн их буурсан (УБ)</div><div class="value" style="font-size:.95rem">${topDown ? topDown.name.split(",")[0] : "—"}</div><div class="sub chg down">${topDown ? fmtPct(topDown.pct) + " · " + fmtMNT(topDown.cur) : ""}</div></div>
        <div class="kpi"><div class="label">Хонины мах — аймаг дундаж</div><div class="value" style="font-size:1.05rem">${fmtMNT(sheepMean)}</div><div class="sub">УБ: ${fmtMNT(sheepUb)}</div></div>
        <div class="kpi"><div class="label">Хугацааны цуваа</div><div class="value">${Math.min(ubWeeks.length, 60)}</div><div class="sub">7-хоног (ачаалсан)</div></div>
      `;

      // Top change table
      const sorted = [...ubRows].filter((r) => r.pct != null).sort((a, b) => Math.abs(b.pct) - Math.abs(a.pct)).slice(0, 15);
      const tb = $("ovTopChgTable").querySelector("tbody");
      tb.innerHTML = sorted.map((r, i) => `
        <tr>
          <td class="rank">${i + 1}</td>
          <td>${escapeHtml(r.name)}</td>
          <td class="num">${fmtMNT(r.cur)}</td>
          <td class="num chg ${chgClass(r.pct)}">${fmtPct(r.pct)}</td>
        </tr>
      `).join("") || `<tr><td colspan="4" class="empty">Мэдээлэл алга</td></tr>`;

      // Shared product bar: aimags + UB
      renderOvBar();
      renderOvUbPrices(ubRows);
      renderBriefing(ubRows);
    }

    function renderOvBar() {
      const key = $("ovProduct").value;
      const sh = SHARED.find((s) => s.key === key) || SHARED[0];
      $("ovSharedProductHint").textContent = sh.label;
      const week = aimagWeeks[0]?.code;
      const weekUb = ubWeeks[0]?.code;

      const labels = [];
      const data = [];
      const bg = [];

      // UB first
      labels.push("Улаанбаатар");
      data.push(ubData[sh.ub]?.[weekUb] ?? null);
      bg.push("#fbbf24");

      const rows = AIMAG_CODES.map((c) => ({
        code: c,
        name: aimagName(c),
        price: aimagData[sh.aimag]?.[c]?.[week] ?? null,
      })).filter((r) => r.price != null).sort((a, b) => a.price - b.price);

      for (const r of rows) {
        labels.push(r.name);
        data.push(r.price);
        bg.push("#38bdf8");
      }

      destroyChart("ovBar");
      charts.ovBar = new Chart($("ovBarChart"), {
        type: "bar",
        data: {
          labels,
          datasets: [{
            label: "Үнэ (₮)",
            data,
            backgroundColor: bg,
            borderRadius: 4,
            maxBarThickness: 18,
          }],
        },
        options: {
          indexAxis: "y",
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: (ctx) => fmtMNT(ctx.raw),
              },
            },
          },
          scales: {
            x: {
              ticks: { callback: (v) => new Intl.NumberFormat("mn-MN", { notation: "compact" }).format(v) },
              grid: { color: "rgba(42,59,85,.5)" },
            },
            y: { grid: { display: false }, ticks: { font: { size: 10 } } },
          },
        },
      });
    }

    function renderOvUbPrices(ubRows) {
      const rows = [...ubRows].sort((a, b) => (b.cur || 0) - (a.cur || 0));
      destroyChart("ovUb");
      charts.ovUb = new Chart($("ovUbPrices"), {
        type: "bar",
        data: {
          labels: rows.map((r) => shortName(r.name)),
          datasets: [{
            label: "Үнэ (₮)",
            data: rows.map((r) => r.cur),
            backgroundColor: rows.map((_, i) => CHART_COLORS[i % CHART_COLORS.length] + "cc"),
            borderRadius: 4,
            maxBarThickness: 22,
          }],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                title: (items) => rows[items[0].dataIndex]?.name || "",
                label: (ctx) => fmtMNT(ctx.raw),
              },
            },
          },
          scales: {
            x: { ticks: { maxRotation: 60, minRotation: 40, font: { size: 9 } }, grid: { display: false } },
            y: {
              ticks: { callback: (v) => new Intl.NumberFormat("mn-MN", { notation: "compact" }).format(v) },
              grid: { color: "rgba(42,59,85,.5)" },
            },
          },
        },
      });
    }

    function shortName(s) {
      return String(s).split(",")[0].slice(0, 22);
    }
    function escapeHtml(s) {
      return String(s)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
    }

    // ─── UB panel ─────────────────────────────────────────────
    const UB_CATEGORIES = [
      { id: "all", label: "Бүгд", re: null },
      { id: "flour", label: "Гурил", re: /гурил/i },
      { id: "bread", label: "Талх", re: /талх/i },
      { id: "meat", label: "Мах", re: /мах/i },
      { id: "dairy", label: "Сүү / цагаан идээ", re: /сүү|тараг|цөцгийн/i },
      { id: "veg", label: "Ногоо / жимс", re: /төмс|лууван|байцаа|манжин|сонгино|алим/i },
      { id: "staple", label: "Үндсэн хүнс", re: /будаа|чихэр|тос|өндөг|цай/i },
      { id: "fuel", label: "Шатахуун", re: /бензин|дизель|аи-/i },
    ];
    let ubActiveCat = "all";
    let ubActiveSub = "dash";
    let ubSort = { key: "pct", dir: -1 };

    function ubPriceAt(code, weekOffset) {
      const w = ubWeeks[weekOffset];
      if (!w) return null;
      return ubData[code]?.[w.code] ?? null;
    }

    function ubSeries(code, n) {
      const count = Math.min(n || 12, ubWeeks.length);
      // oldest -> newest for sparkline
      return ubWeeks.slice(0, count).map((w) => ubData[code]?.[w.code] ?? null).reverse();
    }

    function ubProductRows() {
      return ubProducts.map((p) => {
        const price = ubPriceAt(p.code, 0);
        const prev = ubPriceAt(p.code, 1);
        const w4 = ubPriceAt(p.code, 4);
        const w12 = ubPriceAt(p.code, 12);
        return {
          code: p.code,
          name: p.name,
          price,
          prev,
          diff: price != null && prev != null ? price - prev : null,
          pct: pctChange(price, prev),
          pct4: pctChange(price, w4),
          pct12: pctChange(price, w12),
          spark: ubSeries(p.code, 12),
        };
      });
    }

    function ubCatOf(name) {
      for (const c of UB_CATEGORIES) {
        if (c.id === "all" || !c.re) continue;
        if (c.re.test(name)) return c;
      }
      return { id: "other", label: "Бусад", re: null };
    }

    function filterByUbCat(rows, catId) {
      if (!catId || catId === "all") return rows;
      const cat = UB_CATEGORIES.find((c) => c.id === catId);
      if (!cat || !cat.re) return rows;
      return rows.filter((r) => cat.re.test(r.name));
    }

    function sparkHtml(values) {
      const nums = values.filter((v) => v != null);
      if (!nums.length) return "";
      const min = Math.min(...nums);
      const max = Math.max(...nums);
      const first = nums[0];
      const last = nums[nums.length - 1];
      const trend = last > first ? "up" : last < first ? "down" : "mid";
      return `<span class="spark" title="сүүлийн ${values.length} 7-хоног">${values.map((v) => {
        if (v == null) return `<i style="height:2px;opacity:.3"></i>`;
        const h = max === min ? 12 : 4 + ((v - min) / (max - min)) * 18;
        return `<i class="${trend}" style="height:${h.toFixed(1)}px"></i>`;
      }).join("")}</span>`;
    }

    function setUbSub(sub) {
      ubActiveSub = sub;
      document.querySelectorAll(".ub-subtab").forEach((t) => {
        t.classList.toggle("active", t.dataset.sub === sub);
      });
      document.querySelectorAll(".ub-sub").forEach((el) => {
        el.classList.toggle("active", el.id === "ubSub-" + sub);
      });
      renderUb();
    }

    function renderUb() {
      if (ubActiveSub === "dash") renderUbDash();
      else if (ubActiveSub === "detail") renderUbDetail();
      else if (ubActiveSub === "table") renderUbAllTable();
      else if (ubActiveSub === "multi") renderUbMulti();
    }

    function renderUbDash() {
      const rows = ubProductRows().filter((r) => r.price != null);
      const up = rows.filter((r) => r.pct != null && r.pct > 0.05);
      const down = rows.filter((r) => r.pct != null && r.pct < -0.05);
      const flat = rows.filter((r) => r.pct != null && Math.abs(r.pct) <= 0.05);
      const avgPct = mean(rows.map((r) => r.pct));
      const avg4 = mean(rows.map((r) => r.pct4));
      const avg12 = mean(rows.map((r) => r.pct12));

      $("ubDashKpis").innerHTML = `
        <div class="kpi"><div class="label">Бүтээгдэхүүн</div><div class="value">${rows.length}</div><div class="sub">нийслэлийн 7 хоногийн үнэ</div></div>
        <div class="kpi"><div class="label">Сүүлийн 7 хоног</div><div class="value" style="font-size:1.05rem">${ubWeeks[0]?.name || "—"}</div><div class="sub">шинэчлэгдсэн</div></div>
        <div class="kpi"><div class="label">Дундаж өөрчлөлт (7х)</div><div class="value chg ${chgClass(avgPct)}">${fmtPct(avgPct)}</div><div class="sub">өсөлт ${up.length} · бууралт ${down.length} · тогтвортой ${flat.length}</div></div>
        <div class="kpi"><div class="label">4×7 хоногийн дундаж</div><div class="value chg ${chgClass(avg4)}">${fmtPct(avg4)}</div><div class="sub">~1 сар</div></div>
        <div class="kpi"><div class="label">12×7 хоногийн дундаж</div><div class="value chg ${chgClass(avg12)}">${fmtPct(avg12)}</div><div class="sub">~3 сар</div></div>
        <div class="kpi"><div class="label">Хамгийн их өссөн</div><div class="value" style="font-size:.92rem">${(() => { const t=[...rows].filter(r=>r.pct!=null).sort((a,b)=>b.pct-a.pct)[0]; return t?escapeHtml(shortName(t.name)):"—"; })()}</div><div class="sub chg up">${(() => { const t=[...rows].filter(r=>r.pct!=null).sort((a,b)=>b.pct-a.pct)[0]; return t?fmtPct(t.pct)+" · "+fmtMNT(t.price):""; })()}</div></div>
      `;

      // Category cards
      const catBox = $("ubCatGrid");
      catBox.innerHTML = UB_CATEGORIES.map((c) => {
        const subset = c.id === "all" ? rows : rows.filter((r) => c.re && c.re.test(r.name));
        const ap = mean(subset.map((r) => r.pct));
        const active = ubActiveCat === c.id ? " active" : "";
        return `<div class="ub-cat-card${active}" data-cat="${c.id}">
          <div class="cat-name">${escapeHtml(c.label)}</div>
          <div class="cat-meta">${subset.length} бүтээгдэхүүн</div>
          <div class="cat-pct chg ${chgClass(ap)}">${fmtPct(ap)}</div>
        </div>`;
      }).join("");
      catBox.querySelectorAll(".ub-cat-card").forEach((el) => {
        el.addEventListener("click", () => {
          ubActiveCat = el.dataset.cat;
          renderUbDash();
        });
      });

      const filtered = filterByUbCat(rows, ubActiveCat);
      const catLabel = UB_CATEGORIES.find((c) => c.id === ubActiveCat)?.label || "Бүгд";
      $("ubCatChartHint").textContent = catLabel;

      // Top up / down
      const topUp = [...rows].filter((r) => r.pct != null).sort((a, b) => b.pct - a.pct).slice(0, 8);
      const topDown = [...rows].filter((r) => r.pct != null).sort((a, b) => a.pct - b.pct).slice(0, 8);
      const fillTop = (id, list) => {
        $(id).querySelector("tbody").innerHTML = list.map((r, i) => `
          <tr class="clickable" data-code="${r.code}">
            <td class="rank">${i + 1}</td>
            <td>${escapeHtml(r.name)}</td>
            <td class="num">${fmtMNT(r.price)}</td>
            <td class="num chg ${chgClass(r.pct)}">${fmtPct(r.pct)}</td>
          </tr>`).join("") || `<tr><td colspan="4" class="empty">—</td></tr>`;
        $(id).querySelectorAll("tr[data-code]").forEach((tr) => {
          tr.addEventListener("click", () => {
            $("ubProduct").value = tr.dataset.code;
            setUbSub("detail");
          });
        });
      };
      fillTop("ubTopUpTable", topUp);
      fillTop("ubTopDownTable", topDown);

      // Category product price bars
      const byPrice = [...filtered].sort((a, b) => (b.price || 0) - (a.price || 0));
      destroyChart("ubCat");
      charts.ubCat = new Chart($("ubCatChart"), {
        type: "bar",
        data: {
          labels: byPrice.map((r) => shortName(r.name)),
          datasets: [{
            label: "Үнэ",
            data: byPrice.map((r) => r.price),
            backgroundColor: byPrice.map((r) => {
              const c = chgClass(r.pct);
              return c === "up" ? "#f87171aa" : c === "down" ? "#4ade80aa" : "#38bdf8aa";
            }),
            borderRadius: 4,
            maxBarThickness: 22,
          }],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                title: (items) => byPrice[items[0].dataIndex]?.name || "",
                label: (c) => {
                  const r = byPrice[c.dataIndex];
                  return `${fmtMNT(c.raw)} · 7х: ${fmtPct(r?.pct)}`;
                },
              },
            },
          },
          onClick: (_e, els) => {
            if (!els.length) return;
            const r = byPrice[els[0].index];
            if (r) { $("ubProduct").value = r.code; setUbSub("detail"); }
          },
          scales: {
            x: { ticks: { maxRotation: 50, minRotation: 30, font: { size: 9 } }, grid: { display: false } },
            y: {
              ticks: { callback: (v) => new Intl.NumberFormat("mn-MN", { notation: "compact" }).format(v) },
              grid: { color: "rgba(42,59,85,.5)" },
            },
          },
        },
      });

      // Avg change by category
      const catStats = UB_CATEGORIES.filter((c) => c.id !== "all").map((c) => {
        const sub = rows.filter((r) => c.re && c.re.test(r.name));
        return { label: c.label, pct: mean(sub.map((r) => r.pct)), n: sub.length };
      }).filter((c) => c.n > 0);
      destroyChart("ubCatChg");
      charts.ubCatChg = new Chart($("ubCatChgChart"), {
        type: "bar",
        data: {
          labels: catStats.map((c) => c.label),
          datasets: [{
            data: catStats.map((c) => c.pct),
            backgroundColor: catStats.map((c) =>
              c.pct == null ? "#94a3b8" : c.pct > 0 ? "#f87171" : c.pct < 0 ? "#4ade80" : "#94a3b8"
            ),
            borderRadius: 6,
            maxBarThickness: 36,
          }],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: (c) => `${fmtPct(c.raw)} · ${catStats[c.dataIndex].n} бараа`,
              },
            },
          },
          scales: {
            x: { grid: { display: false } },
            y: { ticks: { callback: (v) => v + "%" }, grid: { color: "rgba(42,59,85,.5)" } },
          },
        },
      });
    }

    function renderUbDetail() {
      const pCode = $("ubProduct").value || ubProducts[0]?.code;
      const win = $("ubWindow").value;
      const pName = ubProducts.find((p) => p.code === pCode)?.name || "";
      $("ubChartTitle").textContent = pName;

      // Product cards for active category filter (chips)
      renderUbProdGrid(pCode);

      let weeks = ubWeeks.slice();
      if (win !== "all") weeks = weeks.slice(0, Number(win));
      const chron = [...weeks].reverse();
      const series = chron.map((w) => ubData[pCode]?.[w.code] ?? null);
      const labels = chron.map((w) => w.name);

      const cur = series[series.length - 1];
      const prev = series.length > 1 ? series[series.length - 2] : null;
      const first = series.find((x) => x != null);
      const w4 = ubPriceAt(pCode, 4);
      const w12 = ubPriceAt(pCode, 12);
      const w26 = ubPriceAt(pCode, 26);
      const pctW = pctChange(cur, prev);
      const pctP = pctChange(cur, first);
      const pct4 = pctChange(cur, w4);
      const pct12 = pctChange(cur, w12);
      const pct26 = pctChange(cur, w26);
      const validS = series.filter((x) => x != null);
      const maxV = validS.length ? Math.max(...validS) : null;
      const minV = validS.length ? Math.min(...validS) : null;
      const avgV = mean(validS);
      const maxIdx = series.lastIndexOf(maxV);
      const minIdx = series.lastIndexOf(minV);

      $("ubKpis").innerHTML = `
        <div class="kpi"><div class="label">Одоогийн үнэ</div><div class="value">${fmtMNT(cur)}</div><div class="sub">${ubWeeks[0]?.name || ""}</div></div>
        <div class="kpi"><div class="label">7 хоног</div><div class="value chg ${chgClass(pctW)}">${fmtPct(pctW)}</div><div class="sub">${prev != null ? fmtMNT(prev) + " → " + fmtMNT(cur) : ""}</div></div>
        <div class="kpi"><div class="label">~1 сар (4×7х)</div><div class="value chg ${chgClass(pct4)}">${fmtPct(pct4)}</div><div class="sub">${w4 != null ? fmtMNT(w4) : "—"}</div></div>
        <div class="kpi"><div class="label">~3 сар (12×7х)</div><div class="value chg ${chgClass(pct12)}">${fmtPct(pct12)}</div><div class="sub">${w12 != null ? fmtMNT(w12) : "—"}</div></div>
        <div class="kpi"><div class="label">~6 сар (26×7х)</div><div class="value chg ${chgClass(pct26)}">${fmtPct(pct26)}</div><div class="sub">${w26 != null ? fmtMNT(w26) : "—"}</div></div>
        <div class="kpi"><div class="label">Цонхны дундаж</div><div class="value" style="font-size:1.05rem">${fmtMNT(avgV)}</div><div class="sub">дээд ${fmtMNT(maxV)} · доод ${fmtMNT(minV)}</div></div>
      `;

      $("ubStatPills").innerHTML = `
        <span class="stat-pill">Оргил: <b>${fmtMNT(maxV)}</b> ${maxIdx >= 0 ? "(" + labels[maxIdx] + ")" : ""}</span>
        <span class="stat-pill">Доод: <b>${fmtMNT(minV)}</b> ${minIdx >= 0 ? "(" + labels[minIdx] + ")" : ""}</span>
        <span class="stat-pill">Цонх: <b>${chron.length}</b> 7-хоног</span>
        <span class="stat-pill">Цонхны нийт өөрчлөлт: <b class="chg ${chgClass(pctP)}">${fmtPct(pctP)}</b></span>
        <span class="stat-pill">Ангилал: <b>${escapeHtml(ubCatOf(pName).label)}</b></span>
      `;

      destroyChart("ubTrend");
      charts.ubTrend = new Chart($("ubTrendChart"), {
        type: "line",
        data: {
          labels,
          datasets: [{
            label: pName,
            data: series,
            borderColor: "#38bdf8",
            backgroundColor: "rgba(56,189,248,.15)",
            fill: true,
            tension: 0.25,
            pointRadius: chron.length > 40 ? 0 : 2,
            pointHoverRadius: 5,
            borderWidth: 2,
          }],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: "index", intersect: false },
          plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: (c) => fmtMNT(c.raw) } },
          },
          scales: {
            x: { ticks: { maxTicksLimit: 10, font: { size: 10 } }, grid: { display: false } },
            y: {
              ticks: { callback: (v) => new Intl.NumberFormat("mn-MN", { notation: "compact" }).format(v) },
              grid: { color: "rgba(42,59,85,.5)" },
            },
          },
        },
      });

      const chgLabels = [];
      const chgData = [];
      const chgColors = [];
      for (let i = 1; i < chron.length; i++) {
        const pct = pctChange(series[i], series[i - 1]);
        chgLabels.push(chron[i].name);
        chgData.push(pct);
        chgColors.push(pct == null ? "#94a3b8" : pct > 0 ? "#f87171" : pct < 0 ? "#4ade80" : "#94a3b8");
      }
      destroyChart("ubChg");
      charts.ubChg = new Chart($("ubChgChart"), {
        type: "bar",
        data: {
          labels: chgLabels,
          datasets: [{
            label: "% өөрчлөлт",
            data: chgData,
            backgroundColor: chgColors,
            borderRadius: 3,
            maxBarThickness: 16,
          }],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: (c) => fmtPct(c.raw) } },
          },
          scales: {
            x: { ticks: { maxTicksLimit: 10, font: { size: 9 } }, grid: { display: false } },
            y: {
              ticks: { callback: (v) => v + "%" },
              grid: { color: "rgba(42,59,85,.5)" },
            },
          },
        },
      });

      // History table (newest first)
      $("ubHistHint").textContent = pName;
      const histBody = $("ubHistTable").querySelector("tbody");
      const histRows = [...weeks]; // newest first
      histBody.innerHTML = histRows.map((w, i) => {
        const price = ubData[pCode]?.[w.code] ?? null;
        const older = weeks[i + 1] ? ubData[pCode]?.[weeks[i + 1].code] ?? null : null;
        const diff = price != null && older != null ? price - older : null;
        const pct = pctChange(price, older);
        const fromPeak = maxV != null && price != null ? price - maxV : null;
        return `<tr>
          <td>${escapeHtml(w.name)}</td>
          <td class="num">${fmtMNT(price)}</td>
          <td class="num chg ${chgClass(pct)}">${diff == null ? "—" : (diff > 0 ? "+" : "") + fmtNum(diff)}</td>
          <td class="num chg ${chgClass(pct)}">${fmtPct(pct)}</td>
          <td class="num">${fromPeak == null ? "—" : (fromPeak > 0 ? "+" : "") + fmtNum(fromPeak)}</td>
        </tr>`;
      }).join("");

      renderUbVsAimag(pCode, pName, cur);
    }

    function renderUbProdGrid(selectedCode) {
      const rows = filterByUbCat(ubProductRows(), ubActiveCat === "all" ? null : ubActiveCat);
      // If chip filter via ubChips sets search-like filter, product select drives list of related
      const qChip = $("ubChips")?.querySelector(".chip.active")?.dataset.q || "";
      let list = rows;
      if (qChip) {
        try {
          const re = new RegExp(qChip, "i");
          list = ubProductRows().filter((r) => re.test(r.name));
        } catch (_) { /* ignore */ }
      } else if (ubActiveCat !== "all") {
        list = filterByUbCat(ubProductRows(), ubActiveCat);
      } else {
        list = ubProductRows();
      }

      const box = $("ubProdGrid");
      if (!box) return;
      box.innerHTML = list.slice(0, 24).map((r) => `
        <div class="ub-prod-card${r.code === selectedCode ? " active" : ""}" data-code="${r.code}">
          <div class="pn">${escapeHtml(r.name)}</div>
          <div class="pp">${fmtMNT(r.price)}</div>
          <div class="pm">
            <span class="chg ${chgClass(r.pct)}">${fmtPct(r.pct)}</span>
            ${sparkHtml(r.spark)}
          </div>
        </div>
      `).join("");
      box.querySelectorAll(".ub-prod-card").forEach((card) => {
        card.addEventListener("click", () => {
          $("ubProduct").value = card.dataset.code;
          renderUbDetail();
        });
      });
    }

    function renderUbVsAimag(ubCode, pName, ubPrice) {
      const wrap = $("ubVsAimagWrap");
      const sh = SHARED.find((s) => s.ub === ubCode);
      if (!sh || !aimagWeeks.length) {
        wrap.style.display = "none";
        return;
      }
      wrap.style.display = "";
      const weekA = aimagWeeks[0]?.code;
      const weekLabel = aimagWeeks[0]?.name;
      const weekU = ubWeeks.find((w) => w.name === weekLabel)?.code || ubWeeks[0]?.code;
      const ubP = ubData[sh.ub]?.[weekU] ?? ubPrice;

      const aimagPrices = AIMAG_CODES.map((c) => ({
        name: aimagName(c),
        price: aimagData[sh.aimag]?.[c]?.[weekA] ?? null,
      })).filter((r) => r.price != null).sort((a, b) => a.price - b.price);

      const avg = mean(aimagPrices.map((r) => r.price));
      const minR = aimagPrices[0];
      const maxR = aimagPrices[aimagPrices.length - 1];
      const vs = pctChange(ubP, avg);
      // rank among aimags+UB
      const all = [
        ...aimagPrices.map((r) => ({ name: r.name, price: r.price, isUb: false })),
        { name: "Улаанбаатар", price: ubP, isUb: true },
      ].filter((r) => r.price != null).sort((a, b) => a.price - b.price);
      const rank = all.findIndex((r) => r.isUb) + 1;

      $("ubVsAimagKpis").innerHTML = `
        <div class="kpi"><div class="label">УБ үнэ</div><div class="value" style="font-size:1.1rem">${fmtMNT(ubP)}</div></div>
        <div class="kpi"><div class="label">Аймаг дундаж</div><div class="value" style="font-size:1.1rem">${fmtMNT(avg)}</div></div>
        <div class="kpi"><div class="label">УБ vs дундаж</div><div class="value chg ${chgClass(vs)}">${fmtPct(vs)}</div></div>
        <div class="kpi"><div class="label">Зэрэглэл (хямд→үнэтэй)</div><div class="value">${rank || "—"} / ${all.length}</div><div class="sub">хямд: ${minR ? escapeHtml(minR.name) : "—"} · үнэтэй: ${maxR ? escapeHtml(maxR.name) : "—"}</div></div>
      `;

      // Trend: last 26 weeks UB vs aimag mean
      const n = Math.min(26, aimagWeeks.length, ubWeeks.length);
      const chronA = aimagWeeks.slice(0, n).reverse();
      const labels = chronA.map((w) => w.name);
      const ubLine = chronA.map((w) => {
        const uCode = ubWeeks.find((u) => u.name === w.name)?.code;
        return uCode ? ubData[sh.ub]?.[uCode] ?? null : null;
      });
      const aimagLine = chronA.map((w) => {
        const vals = AIMAG_CODES.map((c) => aimagData[sh.aimag]?.[c]?.[w.code] ?? null);
        return mean(vals);
      });

      destroyChart("ubVsAimag");
      charts.ubVsAimag = new Chart($("ubVsAimagChart"), {
        type: "line",
        data: {
          labels,
          datasets: [
            {
              label: "Улаанбаатар",
              data: ubLine,
              borderColor: "#fbbf24",
              backgroundColor: "rgba(251,191,36,.12)",
              fill: false,
              tension: 0.25,
              borderWidth: 2.5,
              pointRadius: 0,
              pointHoverRadius: 4,
            },
            {
              label: "Аймаг дундаж",
              data: aimagLine,
              borderColor: "#38bdf8",
              backgroundColor: "rgba(56,189,248,.1)",
              fill: false,
              tension: 0.25,
              borderWidth: 2,
              borderDash: [5, 4],
              pointRadius: 0,
              pointHoverRadius: 4,
            },
          ],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: "index", intersect: false },
          plugins: {
            legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
            tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${fmtMNT(c.raw)}` } },
          },
          scales: {
            x: { ticks: { maxTicksLimit: 8, font: { size: 9 } }, grid: { display: false } },
            y: {
              ticks: { callback: (v) => new Intl.NumberFormat("mn-MN", { notation: "compact" }).format(v) },
              grid: { color: "rgba(42,59,85,.5)" },
            },
          },
        },
      });

      // Rank bar: cheapest 5 + UB + most expensive 5 (unique)
      const show = [];
      const take = (arr) => arr.forEach((x) => {
        if (!show.find((s) => s.name === x.name)) show.push(x);
      });
      take(all.slice(0, 5));
      const ubItem = all.find((x) => x.isUb);
      if (ubItem) take([ubItem]);
      take(all.slice(-5));
      show.sort((a, b) => a.price - b.price);

      destroyChart("ubRank");
      charts.ubRank = new Chart($("ubRankChart"), {
        type: "bar",
        data: {
          labels: show.map((r) => r.name),
          datasets: [{
            data: show.map((r) => r.price),
            backgroundColor: show.map((r) => (r.isUb ? "#fbbf24" : "#818cf8cc")),
            borderRadius: 4,
            maxBarThickness: 18,
          }],
        },
        options: {
          indexAxis: "y",
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: (c) => fmtMNT(c.raw) } },
          },
          scales: {
            x: {
              ticks: { callback: (v) => new Intl.NumberFormat("mn-MN", { notation: "compact" }).format(v) },
              grid: { color: "rgba(42,59,85,.5)" },
            },
            y: { grid: { display: false }, ticks: { font: { size: 10 } } },
          },
        },
      });
    }

    function renderUbAllTable() {
      const q = ($("ubSearch").value || "").trim().toLowerCase();
      const catId = $("ubTableCat")?.value || "";
      const sortKey = $("ubTableSort")?.value || "pct";

      let rows = ubProductRows();
      if (catId) rows = filterByUbCat(rows, catId);
      if (q) rows = rows.filter((r) => r.name.toLowerCase().includes(q));

      const dir = sortKey === "name" ? 1 : -1;
      rows.sort((a, b) => {
        const ka = a[sortKey], kb = b[sortKey];
        if (ka == null && kb == null) return 0;
        if (ka == null) return 1;
        if (kb == null) return -1;
        if (typeof ka === "string") return ka.localeCompare(kb, "mn") * dir;
        return (ka - kb) * dir;
      });

      const avgPct = mean(rows.map((r) => r.pct));
      $("ubTablePills").innerHTML = `
        <span class="stat-pill">Илэрц: <b>${rows.length}</b></span>
        <span class="stat-pill">Дундаж 7х: <b class="chg ${chgClass(avgPct)}">${fmtPct(avgPct)}</b></span>
        <span class="stat-pill">Сүүлийн өгөгдөл: <b>${ubWeeks[0]?.name || "—"}</b></span>
      `;

      const tb = $("ubAllTable").querySelector("tbody");
      const sel = $("ubProduct")?.value;
      tb.innerHTML = rows.map((r) => `
        <tr class="clickable${r.code === sel ? " selected" : ""}" data-code="${r.code}">
          <td>${escapeHtml(r.name)}</td>
          <td class="num">${fmtMNT(r.price)}</td>
          <td class="num chg ${chgClass(r.pct)}">${fmtPct(r.pct)}</td>
          <td class="num chg ${chgClass(r.pct4)}">${fmtPct(r.pct4)}</td>
          <td class="num chg ${chgClass(r.pct12)}">${fmtPct(r.pct12)}</td>
          <td>${sparkHtml(r.spark)}</td>
        </tr>
      `).join("") || `<tr><td colspan="6" class="empty">Илэрц алга</td></tr>`;

      tb.querySelectorAll("tr[data-code]").forEach((tr) => {
        tr.addEventListener("click", () => {
          $("ubProduct").value = tr.dataset.code;
          setUbSub("detail");
        });
      });
    }

    function renderUbMulti() {
      const selected = [...$("ubMultiProducts").selectedOptions].map((o) => o.value);
      const win = Number($("ubMultiWindow").value) || 26;
      const mode = $("ubMultiMode").value;
      if (!selected.length) {
        destroyChart("ubMulti");
        return;
      }
      const weeks = ubWeeks.slice(0, win);
      const chron = [...weeks].reverse();
      const labels = chron.map((w) => w.name);
      const datasets = selected.map((code, i) => {
        const name = ubProducts.find((p) => p.code === code)?.name || code;
        const series = chron.map((w) => ubData[code]?.[w.code] ?? null);
        const color = CHART_COLORS[i % CHART_COLORS.length];
        let data = series;
        if (mode === "index") {
          const base = series.find((x) => x != null);
          data = series.map((v) => (v != null && base ? (v / base) * 100 : null));
        }
        return {
          label: shortName(name),
          data,
          borderColor: color,
          backgroundColor: color + "22",
          tension: 0.25,
          pointRadius: 0,
          pointHoverRadius: 4,
          borderWidth: 2,
        };
      });

      destroyChart("ubMulti");
      charts.ubMulti = new Chart($("ubMultiChart"), {
        type: "line",
        data: { labels, datasets },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: "index", intersect: false },
          plugins: {
            legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
            tooltip: {
              callbacks: {
                label: (c) =>
                  mode === "index"
                    ? `${c.dataset.label}: ${c.raw == null ? "—" : c.raw.toFixed(1)}`
                    : `${c.dataset.label}: ${fmtMNT(c.raw)}`,
              },
            },
          },
          scales: {
            x: { ticks: { maxTicksLimit: 12, font: { size: 10 } }, grid: { display: false } },
            y: {
              title: { display: true, text: mode === "index" ? "Индекс (100=эхлэл)" : "Үнэ (₮)" },
              ticks: {
                callback: (v) =>
                  mode === "index"
                    ? v
                    : new Intl.NumberFormat("mn-MN", { notation: "compact" }).format(v),
              },
              grid: { color: "rgba(42,59,85,.5)" },
            },
          },
        },
      });
    }

    // ─── Aimag panel ──────────────────────────────────────────
    let aimagSort = { key: "priceA", dir: -1 };
    function renderAimag() {
      const pCode = $("aimagProduct").value || aimagProducts[0]?.code;
      const wA = $("aimagWeekA").value || aimagWeeks[0]?.code;
      const wB = $("aimagWeekB").value || aimagWeeks[1]?.code;

      const rows = AIMAG_CODES.map((c) => {
        const priceA = aimagData[pCode]?.[c]?.[wA] ?? null;
        const priceB = aimagData[pCode]?.[c]?.[wB] ?? null;
        const diff = priceA != null && priceB != null ? priceA - priceB : null;
        const pct = pctChange(priceA, priceB);
        return { code: c, name: aimagName(c), priceA, priceB, diff, pct };
      });

      const valid = rows.filter((r) => r.priceA != null);
      const avg = mean(valid.map((r) => r.priceA));
      const avgPct = mean(valid.map((r) => r.pct));
      const maxR = [...valid].sort((a, b) => b.priceA - a.priceA)[0];
      const minR = [...valid].sort((a, b) => a.priceA - b.priceA)[0];
      const maxChg = [...valid].filter((r) => r.pct != null).sort((a, b) => b.pct - a.pct)[0];

      // Mini map on aimag tab
      const pName = aimagProducts.find((p) => p.code === pCode)?.name || "";
      const labelA = aimagWeeks.find((w) => w.code === wA)?.name || "A";
      const labelB = aimagWeeks.find((w) => w.code === wB)?.name || "B";
      ensureMap("aimag");
      paintMap("aimag", rows, {
        productName: pName,
        labelA,
        labelB,
        mode: "pct",
        legendEl: $("aimagMapLegend"),
      });

      $("aimagKpis").innerHTML = `
        <div class="kpi"><div class="label">Аймаг дундаж</div><div class="value">${fmtMNT(avg)}</div><div class="sub">${valid.length} / 21 аймаг</div></div>
        <div class="kpi"><div class="label">Хамгийн үнэтэй</div><div class="value" style="font-size:1rem">${maxR ? maxR.name : "—"}</div><div class="sub">${maxR ? fmtMNT(maxR.priceA) : ""}</div></div>
        <div class="kpi"><div class="label">Хамгийн хямд</div><div class="value" style="font-size:1rem">${minR ? minR.name : "—"}</div><div class="sub">${minR ? fmtMNT(minR.priceA) : ""}</div></div>
        <div class="kpi"><div class="label">Дундаж өөрчлөлт</div><div class="value chg ${chgClass(avgPct)}">${fmtPct(avgPct)}</div><div class="sub">A vs B</div></div>
        <div class="kpi"><div class="label">Хамгийн их өссөн</div><div class="value" style="font-size:1rem">${maxChg ? maxChg.name : "—"}</div><div class="sub chg up">${maxChg ? fmtPct(maxChg.pct) : ""}</div></div>
      `;

      const byPrice = [...valid].sort((a, b) => b.priceA - a.priceA);
      destroyChart("aimagBar");
      charts.aimagBar = new Chart($("aimagBarChart"), {
        type: "bar",
        data: {
          labels: byPrice.map((r) => r.name),
          datasets: [{
            label: "Үнэ A",
            data: byPrice.map((r) => r.priceA),
            backgroundColor: "#818cf8cc",
            borderRadius: 4,
            maxBarThickness: 20,
          }],
        },
        options: {
          indexAxis: "y",
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: (c) => fmtMNT(c.raw) } },
          },
          scales: {
            x: {
              ticks: { callback: (v) => new Intl.NumberFormat("mn-MN", { notation: "compact" }).format(v) },
              grid: { color: "rgba(42,59,85,.5)" },
            },
            y: { grid: { display: false }, ticks: { font: { size: 10 } } },
          },
        },
      });

      const byChg = [...valid].filter((r) => r.pct != null).sort((a, b) => b.pct - a.pct);
      destroyChart("aimagChg");
      charts.aimagChg = new Chart($("aimagChgChart"), {
        type: "bar",
        data: {
          labels: byChg.map((r) => r.name),
          datasets: [{
            label: "%",
            data: byChg.map((r) => r.pct),
            backgroundColor: byChg.map((r) => r.pct > 0 ? "#f87171" : r.pct < 0 ? "#4ade80" : "#94a3b8"),
            borderRadius: 4,
            maxBarThickness: 20,
          }],
        },
        options: {
          indexAxis: "y",
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: (c) => fmtPct(c.raw) } },
          },
          scales: {
            x: { ticks: { callback: (v) => v + "%" }, grid: { color: "rgba(42,59,85,.5)" } },
            y: { grid: { display: false }, ticks: { font: { size: 10 } } },
          },
        },
      });

      // Table
      const sorted = [...rows].sort((a, b) => {
        const ka = a[aimagSort.key], kb = b[aimagSort.key];
        if (ka == null && kb == null) return 0;
        if (ka == null) return 1;
        if (kb == null) return -1;
        if (typeof ka === "string") return ka.localeCompare(kb, "mn") * aimagSort.dir;
        return (ka - kb) * aimagSort.dir;
      });
      const maxP = Math.max(...valid.map((r) => r.priceA), 1);
      const tb = $("aimagTable").querySelector("tbody");
      tb.innerHTML = sorted.map((r, i) => {
        const w = r.priceA != null ? (r.priceA / maxP) * 100 : 0;
        return `
          <tr>
            <td class="rank">${r.priceA != null ? i + 1 : "—"}</td>
            <td>${escapeHtml(r.name)}</td>
            <td class="num">${fmtMNT(r.priceA)}</td>
            <td class="num">${fmtMNT(r.priceB)}</td>
            <td class="num chg ${chgClass(r.pct)}">${r.diff == null ? "—" : (r.diff > 0 ? "+" : "") + fmtNum(r.diff)}</td>
            <td class="num chg ${chgClass(r.pct)}">${fmtPct(r.pct)}</td>
            <td class="bar-cell"><div class="bar"><i style="width:${w}%"></i></div></td>
          </tr>`;
      }).join("");

      renderHeatmap();
    }

    function renderHeatmap() {
      const week = aimagWeeks[0]?.code;
      const products = aimagProducts;
      const box = $("heatMap");
      // Find min/max per product for color scale
      const scales = {};
      for (const p of products) {
        const vals = AIMAG_CODES.map((c) => aimagData[p.code]?.[c]?.[week]).filter((x) => x != null);
        scales[p.code] = {
          min: Math.min(...vals),
          max: Math.max(...vals),
        };
      }

      const cols = products.length + 1;
      const labelW = isMobileView() ? 72 : isTabletView() ? 96 : 120;
      const cellMin = isMobileView() ? 40 : 56;
      box.style.gridTemplateColumns = `${labelW}px repeat(${products.length}, minmax(${cellMin}px, 1fr))`;
      let html = `<div class="heat-cell head"></div>`;
      for (const p of products) {
        html += `<div class="heat-cell head" title="${escapeHtml(p.name)}">${escapeHtml(shortName(p.name))}</div>`;
      }
      for (const code of AIMAG_CODES) {
        html += `<div class="heat-cell label">${escapeHtml(aimagName(code))}</div>`;
        for (const p of products) {
          const v = aimagData[p.code]?.[code]?.[week];
          const { min, max } = scales[p.code];
          let bg = "var(--surface2)";
          let color = "var(--muted)";
          if (v != null && max > min) {
            const t = (v - min) / (max - min);
            // blue (cheap) -> red (expensive)
            const r = Math.round(56 + t * (248 - 56));
            const g = Math.round(189 - t * (189 - 113));
            const b = Math.round(248 - t * (248 - 113));
            bg = `rgba(${r},${g},${b},0.55)`;
            color = "var(--text)";
          }
          const title = v == null ? "—" : `${aimagName(code)} · ${p.name}: ${fmtMNT(v)}`;
          html += `<div class="heat-cell" style="background:${bg};color:${color}" title="${escapeHtml(title)}">${v == null ? "—" : fmtNum(v, 0)}</div>`;
        }
      }
      box.innerHTML = html;
    }

    // ─── Compare ──────────────────────────────────────────────
    function renderCompare() {
      const key = $("cmpProduct").value;
      const sh = SHARED.find((s) => s.key === key) || SHARED[0];
      const weekA = $("cmpWeek").value || aimagWeeks[0]?.code;
      // Align UB week by date label
      const weekLabel = aimagWeeks.find((w) => w.code === weekA)?.name;
      const weekU = ubWeeks.find((w) => w.name === weekLabel)?.code || ubWeeks[0]?.code;

      const aimagPrices = AIMAG_CODES.map((c) => ({
        code: c,
        name: aimagName(c),
        price: aimagData[sh.aimag]?.[c]?.[weekA] ?? null,
      })).filter((r) => r.price != null).sort((a, b) => a.price - b.price);

      const ubPrice = ubData[sh.ub]?.[weekU] ?? null;
      const avgA = mean(aimagPrices.map((r) => r.price));
      const minA = aimagPrices[0];
      const maxA = aimagPrices[aimagPrices.length - 1];
      const vsAvg = pctChange(ubPrice, avgA);

      $("cmpKpis").innerHTML = `
        <div class="kpi"><div class="label">Улаанбаатар</div><div class="value">${fmtMNT(ubPrice)}</div><div class="sub">${weekLabel || ""}</div></div>
        <div class="kpi"><div class="label">Аймаг дундаж</div><div class="value">${fmtMNT(avgA)}</div><div class="sub">21 аймаг</div></div>
        <div class="kpi"><div class="label">УБ vs дундаж</div><div class="value chg ${chgClass(vsAvg)}">${fmtPct(vsAvg)}</div><div class="sub">${vsAvg != null && vsAvg > 0 ? "УБ илүү үнэтэй" : vsAvg != null && vsAvg < 0 ? "УБ хямд" : ""}</div></div>
        <div class="kpi"><div class="label">Хамгийн хямд аймаг</div><div class="value" style="font-size:1rem">${minA ? minA.name : "—"}</div><div class="sub">${minA ? fmtMNT(minA.price) : ""}</div></div>
        <div class="kpi"><div class="label">Хамгийн үнэтэй аймаг</div><div class="value" style="font-size:1rem">${maxA ? maxA.name : "—"}</div><div class="sub">${maxA ? fmtMNT(maxA.price) : ""}</div></div>
      `;

      const labels = ["Улаанбаатар", ...aimagPrices.map((r) => r.name)];
      const data = [ubPrice, ...aimagPrices.map((r) => r.price)];
      const colors = ["#fbbf24", ...aimagPrices.map(() => "#38bdf8")];

      destroyChart("cmp");
      charts.cmp = new Chart($("cmpChart"), {
        type: "bar",
        data: {
          labels,
          datasets: [{
            data,
            backgroundColor: colors,
            borderRadius: 4,
            maxBarThickness: 18,
          }],
        },
        options: {
          indexAxis: "y",
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: (c) => fmtMNT(c.raw) } },
          },
          scales: {
            x: {
              ticks: { callback: (v) => new Intl.NumberFormat("mn-MN", { notation: "compact" }).format(v) },
              grid: { color: "rgba(42,59,85,.5)" },
            },
            y: { grid: { display: false }, ticks: { font: { size: 10 } } },
          },
        },
      });

      // Table all shared products
      const tb = $("cmpTable").querySelector("tbody");
      tb.innerHTML = SHARED.map((s) => {
        const prices = AIMAG_CODES.map((c) => ({
          name: aimagName(c),
          price: aimagData[s.aimag]?.[c]?.[weekA] ?? null,
        })).filter((r) => r.price != null);
        const avg = mean(prices.map((r) => r.price));
        const ub = ubData[s.ub]?.[weekU] ?? null;
        const pct = pctChange(ub, avg);
        const mn = [...prices].sort((a, b) => a.price - b.price)[0];
        const mx = [...prices].sort((a, b) => b.price - a.price)[0];
        return `
          <tr>
            <td>${escapeHtml(s.label)}</td>
            <td class="num">${fmtMNT(ub)}</td>
            <td class="num">${fmtMNT(avg)}</td>
            <td class="num chg ${chgClass(pct)}">${fmtPct(pct)}</td>
            <td class="num">${mn ? escapeHtml(mn.name) + " · " + fmtMNT(mn.price) : "—"}</td>
            <td class="num">${mx ? escapeHtml(mx.name) + " · " + fmtMNT(mx.price) : "—"}</td>
          </tr>`;
      }).join("");
    }

    // ─── Trend multi ──────────────────────────────────────────
    function renderTrend() {
      const source = $("trSource").value;
      const win = $("trWindow").value;
      const selected = [...$("trProducts").selectedOptions].map((o) => o.value);
      if (!selected.length) {
        destroyChart("trIdx");
        destroyChart("trPrice");
        return;
      }

      let weeks = source === "ub" ? ubWeeks.slice() : aimagWeeks.slice();
      if (win !== "all") weeks = weeks.slice(0, Number(win));
      const chron = [...weeks].reverse();
      const labels = chron.map((w) => w.name);
      const aimagCode = $("trAimag").value || AIMAG_CODES[0];
      const prodList = source === "ub" ? ubProducts : aimagProducts;

      const datasetsPrice = [];
      const datasetsIdx = [];
      selected.forEach((code, i) => {
        const name = prodList.find((p) => p.code === code)?.name || code;
        const series = chron.map((w) => {
          if (source === "ub") return ubData[code]?.[w.code] ?? null;
          return aimagData[code]?.[aimagCode]?.[w.code] ?? null;
        });
        const color = CHART_COLORS[i % CHART_COLORS.length];
        datasetsPrice.push({
          label: shortName(name),
          data: series,
          borderColor: color,
          backgroundColor: color + "33",
          tension: 0.25,
          pointRadius: 0,
          pointHoverRadius: 4,
          borderWidth: 2,
        });
        const base = series.find((x) => x != null);
        const idx = series.map((v) => (v != null && base ? (v / base) * 100 : null));
        datasetsIdx.push({
          label: shortName(name),
          data: idx,
          borderColor: color,
          backgroundColor: color + "33",
          tension: 0.25,
          pointRadius: 0,
          pointHoverRadius: 4,
          borderWidth: 2,
        });
      });

      destroyChart("trIdx");
      charts.trIdx = new Chart($("trIndexChart"), {
        type: "line",
        data: { labels, datasets: datasetsIdx },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: "index", intersect: false },
          plugins: {
            legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
            tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${c.raw == null ? "—" : c.raw.toFixed(1)}` } },
          },
          scales: {
            x: { ticks: { maxTicksLimit: 12, font: { size: 10 } }, grid: { display: false } },
            y: {
              title: { display: true, text: "Индекс (эхлэл=100)" },
              grid: { color: "rgba(42,59,85,.5)" },
            },
          },
        },
      });

      destroyChart("trPrice");
      charts.trPrice = new Chart($("trPriceChart"), {
        type: "line",
        data: { labels, datasets: datasetsPrice },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: "index", intersect: false },
          plugins: {
            legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
            tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${fmtMNT(c.raw)}` } },
          },
          scales: {
            x: { ticks: { maxTicksLimit: 12, font: { size: 10 } }, grid: { display: false } },
            y: {
              ticks: { callback: (v) => new Intl.NumberFormat("mn-MN", { notation: "compact" }).format(v) },
              grid: { color: "rgba(42,59,85,.5)" },
            },
          },
        },
      });
    }

    // ─── Excel / CSV export (client-side, public, no login) ───
    function fileStamp() {
      const d = ubWeeks[0]?.name || aimagWeeks[0]?.name || new Date().toISOString().slice(0, 10);
      return String(d).replace(/[^\d-]/g, "") || "export";
    }

    function hasExportData() {
      return ubProducts.length > 0 || aimagProducts.length > 0;
    }

    function setExportEnabled(on) {
      const btn = $("btnDlMenu");
      if (btn) btn.disabled = !on;
      ["btnDlUbTable", "btnDlUbHist", "btnDlAimag", "btnDlCompare"].forEach((id) => {
        const el = $(id);
        if (el) el.disabled = !on;
      });
    }

    function downloadBlob(filename, blob) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1500);
    }

    /** CSV with UTF-8 BOM so Excel opens Mongolian correctly */
    function downloadCsv(filename, rows) {
      const esc = (v) => {
        if (v == null || v === "") return "";
        const s = String(v);
        if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
        return s;
      };
      const body = rows.map((r) => r.map(esc).join(",")).join("\r\n");
      const bom = "\uFEFF";
      downloadBlob(filename, new Blob([bom + body], { type: "text/csv;charset=utf-8" }));
    }

    function sheetFromAoA(aoa) {
      if (typeof XLSX === "undefined") throw new Error("XLSX library not loaded");
      return XLSX.utils.aoa_to_sheet(aoa);
    }

    function downloadXlsx(filename, sheets) {
      // sheets: { "SheetName": [[...], ...] }
      if (typeof XLSX === "undefined") {
        // Fallback: first sheet as CSV
        const first = Object.keys(sheets)[0];
        downloadCsv(filename.replace(/\.xlsx$/i, ".csv"), sheets[first] || []);
        alert("Excel сан ачаалагдаагүй тул CSV болгон татлаа. Интернэт холболтоо шалгаад дахин оролдоно уу.");
        return;
      }
      const wb = XLSX.utils.book_new();
      for (const [name, aoa] of Object.entries(sheets)) {
        const safe = String(name).slice(0, 31) || "Sheet";
        const ws = sheetFromAoA(aoa);
        // rough column widths
        const cols = [];
        const width = Math.min(20, (aoa[0] || []).length);
        for (let c = 0; c < width; c++) {
          let max = 10;
          for (let r = 0; r < Math.min(aoa.length, 80); r++) {
            const cell = aoa[r]?.[c];
            if (cell != null) max = Math.max(max, String(cell).length);
          }
          cols.push({ wch: Math.min(40, max + 2) });
        }
        ws["!cols"] = cols;
        XLSX.utils.book_append_sheet(wb, ws, safe);
      }
      XLSX.writeFile(wb, filename);
    }

    function metaSheet() {
      return [
        ["Талбар", "Утга"],
        ["Гарчиг", "7 хоногийн үнийн дашбоард — ҮСХ 1212.mn"],
        ["Сүүлийн 7 хоног (УБ)", ubWeeks[0]?.name || ""],
        ["Сүүлийн 7 хоног (аймаг)", aimagWeeks[0]?.name || ""],
        ["УБ бүтээгдэхүүн", ubProducts.length],
        ["Аймгийн бүтээгдэхүүн", aimagProducts.length],
        ["Аймаг", AIMAG_CODES.length],
        ["Ачаалсан 7-хоног (УБ)", ubWeeks.length],
        ["Ачаалсан 7-хоног (аймаг)", aimagWeeks.length],
        ["УБ эх", "DT_NSO_0600_001V4"],
        ["Аймаг эх", "DT_NSO_0300_010V5"],
        ["API", "https://data.1212.mn/"],
        ["Вэб", "https://batsukh1111.github.io/mongolia-weekly-price-dashboard/"],
        ["Татсан огноо (UTC)", new Date().toISOString()],
        ["Нэгж", "Төгрөг (₮)"],
        ["Тэмдэглэл", "Хэн ч нэвтрэлгүйгээр татаж болно. Өгөгдөл ҮСХ-ийн PxWeb API-аас амьдаар татагдсан."],
      ];
    }

    function buildUbLatestAoA() {
      const rows = [["Бүтээгдэхүүн", "Код", "Ангилал", "Үнэ_сүүлийн", "Огноо_сүүлийн", "Үнэ_өмнөх", "Огноо_өмнөх", "Зөрүү_₮", "Өөрчлөлт_%_7х", "Өөрчлөлт_%_4x7х", "Өөрчлөлт_%_12x7х"]];
      const w0 = ubWeeks[0]?.name || "";
      const w1 = ubWeeks[1]?.name || "";
      for (const r of ubProductRows()) {
        rows.push([
          r.name,
          r.code,
          ubCatOf(r.name).label,
          r.price,
          w0,
          r.prev,
          w1,
          r.diff,
          r.pct != null ? Number(r.pct.toFixed(3)) : null,
          r.pct4 != null ? Number(r.pct4.toFixed(3)) : null,
          r.pct12 != null ? Number(r.pct12.toFixed(3)) : null,
        ]);
      }
      return rows;
    }

    function buildUbSeriesAoA() {
      // Wide: product | date1 | date2 | ...
      const weeks = ubWeeks.slice(); // newest first labels as columns
      const header = ["Бүтээгдэхүүн", "Код", ...weeks.map((w) => w.name)];
      const rows = [header];
      for (const p of ubProducts) {
        rows.push([
          p.name,
          p.code,
          ...weeks.map((w) => ubData[p.code]?.[w.code] ?? null),
        ]);
      }
      return rows;
    }

    function buildUbLongAoA() {
      const rows = [["Бүтээгдэхүүн", "Код", "Огноо", "Үнэ_₮"]];
      for (const p of ubProducts) {
        for (const w of ubWeeks) {
          const v = ubData[p.code]?.[w.code];
          if (v == null) continue;
          rows.push([p.name, p.code, w.name, v]);
        }
      }
      return rows;
    }

    function buildAimagLatestAoA() {
      const w0 = aimagWeeks[0];
      const w1 = aimagWeeks[1];
      const rows = [["Бүтээгдэхүүн", "Бүтээгдэхүүн_код", "Аймаг", "Аймаг_код", "Үнэ_сүүлийн", "Огноо_сүүлийн", "Үнэ_өмнөх", "Огноо_өмнөх", "Зөрүү_₮", "Өөрчлөлт_%"]];
      for (const p of aimagProducts) {
        for (const code of AIMAG_CODES) {
          const a = aimagData[p.code]?.[code]?.[w0?.code] ?? null;
          const b = aimagData[p.code]?.[code]?.[w1?.code] ?? null;
          const diff = a != null && b != null ? a - b : null;
          const pct = pctChange(a, b);
          rows.push([
            p.name,
            p.code,
            aimagName(code),
            code,
            a,
            w0?.name || "",
            b,
            w1?.name || "",
            diff,
            pct != null ? Number(pct.toFixed(3)) : null,
          ]);
        }
      }
      return rows;
    }

    function buildAimagLongAoA() {
      const rows = [["Бүтээгдэхүүн", "Бүтээгдэхүүн_код", "Аймаг", "Аймаг_код", "Огноо", "Үнэ_₮"]];
      for (const p of aimagProducts) {
        for (const code of AIMAG_CODES) {
          for (const w of aimagWeeks) {
            const v = aimagData[p.code]?.[code]?.[w.code];
            if (v == null) continue;
            rows.push([p.name, p.code, aimagName(code), code, w.name, v]);
          }
        }
      }
      return rows;
    }

    function buildCompareAoA() {
      const weekA = aimagWeeks[0]?.code;
      const weekLabel = aimagWeeks[0]?.name;
      const weekU = ubWeeks.find((w) => w.name === weekLabel)?.code || ubWeeks[0]?.code;
      const rows = [["Бүтээгдэхүүн", "УБ_үнэ", "Аймаг_дундаж", "Зөрүү_%_УБ_vs_дундаж", "Хамгийн_хямд_аймаг", "Хамгийн_хямд_үнэ", "Хамгийн_үнэтэй_аймаг", "Хамгийн_үнэтэй_үнэ", "Огноо"]];
      for (const s of SHARED) {
        const prices = AIMAG_CODES.map((c) => ({
          name: aimagName(c),
          price: aimagData[s.aimag]?.[c]?.[weekA] ?? null,
        })).filter((r) => r.price != null);
        const avg = mean(prices.map((r) => r.price));
        const ub = ubData[s.ub]?.[weekU] ?? null;
        const pct = pctChange(ub, avg);
        const mn = [...prices].sort((a, b) => a.price - b.price)[0];
        const mx = [...prices].sort((a, b) => b.price - a.price)[0];
        rows.push([
          s.label,
          ub,
          avg != null ? Number(avg.toFixed(2)) : null,
          pct != null ? Number(pct.toFixed(3)) : null,
          mn?.name || "",
          mn?.price ?? null,
          mx?.name || "",
          mx?.price ?? null,
          weekLabel || "",
        ]);
      }
      return rows;
    }

    function buildUbHistAoA(pCode) {
      const pName = ubProducts.find((p) => p.code === pCode)?.name || pCode;
      const win = $("ubWindow")?.value;
      let weeks = ubWeeks.slice();
      if (win && win !== "all") weeks = weeks.slice(0, Number(win));
      const rows = [["Бүтээгдэхүүн", "Огноо", "Үнэ_₮", "Өмнөхөөс_₮", "Өөрчлөлт_%"]];
      for (let i = 0; i < weeks.length; i++) {
        const w = weeks[i];
        const price = ubData[pCode]?.[w.code] ?? null;
        const older = weeks[i + 1] ? ubData[pCode]?.[weeks[i + 1].code] ?? null : null;
        const diff = price != null && older != null ? price - older : null;
        const pct = pctChange(price, older);
        rows.push([
          pName,
          w.name,
          price,
          diff,
          pct != null ? Number(pct.toFixed(3)) : null,
        ]);
      }
      return rows;
    }

    function buildAimagViewAoA() {
      const pCode = $("aimagProduct")?.value || aimagProducts[0]?.code;
      const wA = $("aimagWeekA")?.value || aimagWeeks[0]?.code;
      const wB = $("aimagWeekB")?.value || aimagWeeks[1]?.code;
      const pName = aimagProducts.find((p) => p.code === pCode)?.name || "";
      const labelA = aimagWeeks.find((w) => w.code === wA)?.name || "A";
      const labelB = aimagWeeks.find((w) => w.code === wB)?.name || "B";
      const rows = [["Бүтээгдэхүүн", "Аймаг", "Үнэ_" + labelA, "Үнэ_" + labelB, "Зөрүү_₮", "Өөрчлөлт_%"]];
      for (const c of AIMAG_CODES) {
        const a = aimagData[pCode]?.[c]?.[wA] ?? null;
        const b = aimagData[pCode]?.[c]?.[wB] ?? null;
        const diff = a != null && b != null ? a - b : null;
        const pct = pctChange(a, b);
        rows.push([
          pName,
          aimagName(c),
          a,
          b,
          diff,
          pct != null ? Number(pct.toFixed(3)) : null,
        ]);
      }
      return rows;
    }

    function exportFullXlsx() {
      downloadXlsx(`une-dashboard-${fileStamp()}-bugd.xlsx`, {
        "Метадата": metaSheet(),
        "УБ_сүүлийн": buildUbLatestAoA(),
        "УБ_цуврал": buildUbSeriesAoA(),
        "Аймаг_сүүлийн": buildAimagLatestAoA(),
        "Аймаг_цуврал_урт": buildAimagLongAoA(),
        "УБ_vs_аймаг": buildCompareAoA(),
      });
    }

    function exportUbXlsx() {
      downloadXlsx(`une-UB-${fileStamp()}.xlsx`, {
        "Метадата": metaSheet(),
        "УБ_сүүлийн": buildUbLatestAoA(),
        "УБ_цуврал": buildUbSeriesAoA(),
        "УБ_урт": buildUbLongAoA(),
      });
    }

    function exportAimagXlsx() {
      downloadXlsx(`une-aimag-${fileStamp()}.xlsx`, {
        "Метадата": metaSheet(),
        "Аймаг_сүүлийн": buildAimagLatestAoA(),
        "Аймаг_цуврал_урт": buildAimagLongAoA(),
      });
    }

    function exportCurrentView() {
      const tab = document.querySelector(".tab.active")?.dataset.tab || "overview";
      if (tab === "ub") {
        if (ubActiveSub === "detail") {
          const code = $("ubProduct")?.value || ubProducts[0]?.code;
          downloadXlsx(`une-UB-tuuh-${fileStamp()}.xlsx`, {
            "Түүх": buildUbHistAoA(code),
            "УБ_сүүлийн": buildUbLatestAoA(),
          });
        } else if (ubActiveSub === "table") {
          downloadXlsx(`une-UB-jagsaalt-${fileStamp()}.xlsx`, {
            "УБ_сүүлийн": buildUbLatestAoA(),
          });
        } else if (ubActiveSub === "multi") {
          const selected = [...($("ubMultiProducts")?.selectedOptions || [])].map((o) => o.value);
          const win = Number($("ubMultiWindow")?.value) || 26;
          const weeks = ubWeeks.slice(0, win);
          const header = ["Бүтээгдэхүүн", ...weeks.map((w) => w.name).reverse()];
          const rows = [header];
          for (const code of selected) {
            const name = ubProducts.find((p) => p.code === code)?.name || code;
            const chron = [...weeks].reverse();
            rows.push([name, ...chron.map((w) => ubData[code]?.[w.code] ?? null)]);
          }
          downloadXlsx(`une-UB-olon-${fileStamp()}.xlsx`, { "Олон_бүтээгдэхүүн": rows });
        } else {
          exportUbXlsx();
        }
      } else if (tab === "aimag" || tab === "map") {
        downloadXlsx(`une-aimag-haragdal-${fileStamp()}.xlsx`, {
          "Аймаг_сонгосон": buildAimagViewAoA(),
          "Аймаг_сүүлийн": buildAimagLatestAoA(),
        });
      } else if (tab === "compare") {
        downloadXlsx(`une-haritsuulalt-${fileStamp()}.xlsx`, {
          "УБ_vs_аймаг": buildCompareAoA(),
        });
      } else {
        exportFullXlsx();
      }
    }

    function runExport(kind) {
      if (!hasExportData()) {
        alert("Өгөгдөл хараахан ачаалагдаагүй байна. Түр хүлээгээд дахин оролдоно уу.");
        return;
      }
      try {
        if (kind === "full-xlsx") exportFullXlsx();
        else if (kind === "ub-xlsx") exportUbXlsx();
        else if (kind === "aimag-xlsx") exportAimagXlsx();
        else if (kind === "ub-csv") downloadCsv(`une-UB-${fileStamp()}.csv`, buildUbLatestAoA());
        else if (kind === "aimag-csv") downloadCsv(`une-aimag-${fileStamp()}.csv`, buildAimagLatestAoA());
        else if (kind === "full-csv") {
          downloadCsv(`une-UB-${fileStamp()}.csv`, buildUbLatestAoA());
          setTimeout(() => downloadCsv(`une-aimag-${fileStamp()}.csv`, buildAimagLatestAoA()), 400);
        } else if (kind === "current") exportCurrentView();
        else if (kind === "ub-hist") {
          const code = $("ubProduct")?.value || ubProducts[0]?.code;
          downloadXlsx(`une-UB-tuuh-${fileStamp()}.xlsx`, { "Түүх": buildUbHistAoA(code) });
        } else if (kind === "ub-table") {
          downloadXlsx(`une-UB-jagsaalt-${fileStamp()}.xlsx`, { "УБ_сүүлийн": buildUbLatestAoA() });
        } else if (kind === "aimag-view") {
          downloadXlsx(`une-aimag-haragdal-${fileStamp()}.xlsx`, { "Аймаг_сонгосон": buildAimagViewAoA() });
        } else if (kind === "compare") {
          downloadXlsx(`une-haritsuulalt-${fileStamp()}.xlsx`, { "УБ_vs_аймаг": buildCompareAoA() });
        }
      } catch (err) {
        console.error(err);
        alert("Таталт амжилтгүй: " + (err.message || err));
      }
    }

    // ─── Events ───────────────────────────────────────────────
    function bindEvents() {
      $("btnRefresh").addEventListener("click", () => init(true));

      // Download menu
      const dlWrap = $("dlWrap");
      const dlBtn = $("btnDlMenu");
      dlBtn?.addEventListener("click", (e) => {
        e.stopPropagation();
        dlWrap?.classList.toggle("open");
      });
      document.addEventListener("click", (e) => {
        if (dlWrap && !dlWrap.contains(e.target)) dlWrap.classList.remove("open");
      });
      $("dlMenu")?.querySelectorAll("button[data-export]").forEach((b) => {
        b.addEventListener("click", () => {
          dlWrap?.classList.remove("open");
          runExport(b.dataset.export);
        });
      });
      $("btnDlUbTable")?.addEventListener("click", () => runExport("ub-table"));
      $("btnDlUbHist")?.addEventListener("click", () => runExport("ub-hist"));
      $("btnDlAimag")?.addEventListener("click", () => runExport("aimag-view"));
      $("btnDlCompare")?.addEventListener("click", () => runExport("compare"));
      $("ovProduct").addEventListener("change", renderOvBar);
      $("ubApply").addEventListener("click", () => renderUbDetail());
      $("ubProduct").addEventListener("change", () => renderUbDetail());
      $("ubWindow").addEventListener("change", () => renderUbDetail());
      $("ubSearch").addEventListener("input", renderUbAllTable);
      $("ubTableCat")?.addEventListener("change", renderUbAllTable);
      $("ubTableSort")?.addEventListener("change", renderUbAllTable);
      $("ubMultiApply")?.addEventListener("click", renderUbMulti);
      $("ubMultiWindow")?.addEventListener("change", renderUbMulti);
      $("ubMultiMode")?.addEventListener("change", renderUbMulti);
      $("ubMultiProducts")?.addEventListener("change", renderUbMulti);
      document.querySelectorAll(".ub-subtab").forEach((tab) => {
        tab.addEventListener("click", () => setUbSub(tab.dataset.sub));
      });
      $("aimagApply").addEventListener("click", renderAimag);
      $("aimagProduct").addEventListener("change", renderAimag);
      $("mapApply").addEventListener("click", renderMapPanel);
      $("mapProduct").addEventListener("change", renderMapPanel);
      $("mapWeekA").addEventListener("change", renderMapPanel);
      $("mapWeekB").addEventListener("change", renderMapPanel);
      $("mapModeToggle").querySelectorAll(".chip").forEach((chip) => {
        chip.addEventListener("click", () => {
          $("mapModeToggle").querySelectorAll(".chip").forEach((c) => c.classList.remove("active"));
          chip.classList.add("active");
          mapMode = chip.dataset.mode || "pct";
          renderMapPanel();
        });
      });
      $("cmpApply").addEventListener("click", renderCompare);
      $("cmpProduct").addEventListener("change", renderCompare);
      $("cmpWeek").addEventListener("change", renderCompare);
      $("trApply").addEventListener("click", renderTrend);
      $("trSource").addEventListener("change", () => {
        const src = $("trSource").value;
        $("trAimagField").style.display = src === "aimag" ? "" : "none";
        fillTrProducts(src);
        [...$("trProducts").options].slice(0, 3).forEach((o) => { o.selected = true; });
        renderTrend();
      });

      $("ubAllTable")?.querySelectorAll("th[data-k]").forEach((th) => {
        th.addEventListener("click", () => {
          const k = th.dataset.k;
          if ($("ubTableSort")) {
            // map header click to sort dropdown
            const map = { name: "name", price: "price", pct: "pct", pct4: "pct4", pct12: "pct12" };
            if (map[k]) $("ubTableSort").value = map[k];
          }
          if (ubSort.key === k) ubSort.dir *= -1;
          else { ubSort.key = k; ubSort.dir = k === "name" ? 1 : -1; }
          renderUbAllTable();
        });
      });
      $("aimagTable").querySelectorAll("th[data-k]").forEach((th) => {
        th.addEventListener("click", () => {
          const k = th.dataset.k;
          if (aimagSort.key === k) aimagSort.dir *= -1;
          else { aimagSort.key = k; aimagSort.dir = k === "name" ? 1 : -1; }
          renderAimag();
        });
      });
    }


    // ─── Weekly briefing + global search ──────────────────────
    function renderBriefing(ubRows) {
      const el = $("briefingText");
      const meta = $("briefingMeta");
      if (!el) return;
      const rows = (ubRows || []).filter((r) => r.pct != null);
      const up = rows.filter((r) => r.pct > 0.05);
      const down = rows.filter((r) => r.pct < -0.05);
      const avgPct = mean(rows.map((r) => r.pct));
      const topUp = [...rows].sort((a, b) => b.pct - a.pct)[0];
      const topDown = [...rows].sort((a, b) => a.pct - b.pct)[0];
      const week = ubWeeks[0]?.name || "";
      if (meta) meta.textContent = week ? (week + " · Улаанбаатар") : "";
      if (!rows.length) {
        el.textContent = "Өгөгдөл ачаалагдахад энд 7 хоногийн тойм гарна.";
        return;
      }
      const tone = avgPct != null && avgPct > 0.05 ? "өсөх" : avgPct != null && avgPct < -0.05 ? "буурах" : "тогтвортой";
      el.innerHTML =
        "Энэ 7 хоногт Улаанбаатарын <b>" + rows.length + "</b> бүтээгдэхүүний дундаж хөдөлгөөн <b class=\"chg " + chgClass(avgPct) + "\">" + fmtPct(avgPct) + "</b> (" + tone + "). " +
        "<b>" + up.length + "</b> нэр төрөл өсөж, <b>" + down.length + "</b> буурсан." +
        (topUp ? " Хамгийн их өссөн нь <button type=\"button\" class=\"inline-link\" data-jump=\"" + topUp.code + "\">" + escapeHtml(shortName(topUp.name)) + "</button> (" + fmtPct(topUp.pct) + ", " + fmtMNT(topUp.cur) + ")." : "") +
        (topDown ? " Хамгийн их буурсан нь <button type=\"button\" class=\"inline-link\" data-jump=\"" + topDown.code + "\">" + escapeHtml(shortName(topDown.name)) + "</button> (" + fmtPct(topDown.pct) + ", " + fmtMNT(topDown.cur) + ")." : "");
      el.querySelectorAll("[data-jump]").forEach((btn) => {
        btn.addEventListener("click", () => {
          activateTab("ub");
          $("ubProduct").value = btn.dataset.jump;
          setUbSub("detail");
        });
      });
    }

    function jumpToUbProduct(code) {
      if (!$("ubProduct")) return;
      activateTab("ub");
      $("ubProduct").value = code;
      setUbSub("detail");
    }

    function setupGlobalSearch() {
      const input = $("globalSearch");
      const box = $("searchResults");
      if (!input || !box) return;
      const close = () => { box.hidden = true; box.innerHTML = ""; };
      const render = () => {
        const q = input.value.trim().toLowerCase();
        if (q.length < 1) { close(); return; }
        const hits = [];
        for (const p of ubProducts) {
          if (p.name.toLowerCase().includes(q)) hits.push({ kind: "УБ", name: p.name, go: () => jumpToUbProduct(p.code) });
        }
        for (const p of aimagProducts) {
          if (p.name.toLowerCase().includes(q)) hits.push({ kind: "Аймаг", name: p.name, go: () => { activateTab("map"); $("mapProduct").value = p.code; renderMapPanel(); } });
        }
        const shown = hits.slice(0, 8);
        if (!shown.length) {
          box.hidden = false;
          box.innerHTML = "<div class=\"search-empty\">Илэрц алга</div>";
          return;
        }
        box.hidden = false;
        box.innerHTML = shown.map((h, i) =>
          "<button type=\"button\" class=\"search-hit\" data-i=\"" + i + "\"><span class=\"search-kind\">" + h.kind + "</span><span>" + escapeHtml(h.name) + "</span></button>"
        ).join("");
        box.querySelectorAll(".search-hit").forEach((btn) => {
          btn.addEventListener("click", () => {
            shown[Number(btn.dataset.i)].go();
            input.value = "";
            close();
          });
        });
      };
      input.addEventListener("input", render);
      input.addEventListener("keydown", (e) => {
        if (e.key === "Escape") { input.blur(); close(); }
        if (e.key === "Enter") {
          const first = box.querySelector(".search-hit");
          if (first) first.click();
        }
      });
      document.addEventListener("click", (e) => {
        if (!e.target.closest(".search-wrap")) close();
      });
    }

    function applyHashTab() {
      const raw = (location.hash || "#overview").replace("#", "") || "overview";
      const name = document.querySelector('.tab[data-tab="' + raw + '"]') ? raw : "overview";
      activateTab(name, { skipHash: true });
    }

    // ─── Init ─────────────────────────────────────────────────
    let eventsBound = false;
    async function init() {
      clearError();
      showLoading("1212.mn API-тай холбогдож байна…");
      try {
        chartDefaults();
        await loadAllData();
        setupControls();
        if (!eventsBound) {
          setupTabs();
          bindEvents();
          setupGlobalSearch();
          window.addEventListener("hashchange", applyHashTab);
          eventsBound = true;
        }
        applyHashTab();
        setTimeout(invalidateMaps, 120);
        setExportEnabled(true);
        hideLoading();
      } catch (err) {
        console.error(err);
        hideLoading();
        setExportEnabled(false);
        showError(
          "1212.mn API-тай холбогдож чадсангүй. Интернэт холболт болон CORS-ийг шалгана уу. " +
          "GitHub Pages дээр эсвэл локал серверээр (python -m http.server) нээнэ үү. " +
          "Дэлгэрэнгүй: " + (err && err.message ? err.message : err)
        );
      }
    }

    init();
