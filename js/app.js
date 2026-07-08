/**
 * WL2 Performance Dashboard — application layer.
 * Wires up: file upload -> engine.js (parsing/KPI calc) -> components/charts.js
 * (rendering) -> DOM. Holds UI state (filters, granularity) and orchestrates
 * re-rendering. Labels/colors/thresholds are NOT hardcoded here — they come
 * from js/config/config.js so this file rarely needs editing.
 */
(function () {
  "use strict";

  var CONFIG = window.WL2_CONFIG;
  var lineChart = Charts.lineChart, barChart = Charts.barChart;

  var CONTRACTOR_LABELS = CONFIG.CONTRACTOR_LABELS;
  var FLEET_DEFS = CONFIG.FLEET_DEFS;
  var CRUSHER_DEFS = CONFIG.CRUSHER_DEFS;
  var CAUSE_GROUPS = CONFIG.CAUSE_GROUPS;
  var THRESH = CONFIG.STATUS_THRESHOLDS;

  var state = {
    fileName: null, uploadedAt: null,
    assumptions: null, planMonthly: null, planWeekly: null, machines: null, contractors: null,
    rawDays: [], kpis: [], dataQuality: [],
    granularity: "day", // day | week | month
    rangeFrom: null, rangeTo: null, rangePreset: "all",
    monthlyMethod: "weighted",
    chartTableVisible: {}
  };

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $all(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }
  function fmt(v, digits) { return Charts.fmtNum(v, digits); }
  function pct(v, digits) { return v === null || v === undefined || isNaN(v) ? "-" : v.toFixed(digits === undefined ? 1 : digits) + "%"; }
  function dateKey(d) { return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
  function shortDate(d) { return (d.getMonth() + 1) + "/" + d.getDate(); }
  var MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  // ---------------- Theme ----------------
  function initTheme() {
    var saved = localStorage.getItem(CONFIG.STORAGE_KEYS.THEME);
    if (saved) document.documentElement.setAttribute("data-theme", saved);
    $("#themeToggle").addEventListener("click", function () {
      var cur = document.documentElement.getAttribute("data-theme");
      var prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
      var effectiveCur = cur || (prefersDark ? "dark" : "light");
      var next = effectiveCur === "dark" ? "light" : "dark";
      document.documentElement.setAttribute("data-theme", next);
      localStorage.setItem(CONFIG.STORAGE_KEYS.THEME, next);
      if (state.kpis.length) renderAll();
    });
  }

  // ---------------- File handling ----------------
  function arrayBufferToBase64(buf) {
    var bytes = new Uint8Array(buf);
    var chunk = 0x8000, str = "";
    for (var i = 0; i < bytes.length; i += chunk) {
      str += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(str);
  }
  function base64ToArrayBuffer(b64) {
    var str = atob(b64);
    var bytes = new Uint8Array(str.length);
    for (var i = 0; i < str.length; i++) bytes[i] = str.charCodeAt(i);
    return bytes.buffer;
  }

  function processWorkbook(arrayBuffer, fileName, persist) {
    var wb;
    try {
      wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    } catch (e) {
      alert("ไม่สามารถอ่านไฟล์ Excel นี้ได้: " + e.message);
      return;
    }
    function rows(name) {
      if (!wb.Sheets[name]) return [];
      return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    }
    try {
      state.assumptions = WL2.parseAssumptions(rows("Assumptions"));
      state.planMonthly = WL2.parsePlanMonthly(rows("Plan_Monthly"));
      state.planWeekly = WL2.parsePlanWeekly(rows("Plan_Weekly"));
      state.machines = WL2.parseMasterMachine(rows("Master_Machine"));
      state.contractors = WL2.parseMasterContractor(rows("Master_Contractor"));
      state.rawDays = WL2.parseDailyInput(rows("Daily_Input"));
      state.kpis = WL2.computeAllDays(state.rawDays, state.assumptions, state.planMonthly);
      state.dataQuality = WL2.computeDataQuality(state.rawDays);
    } catch (e) {
      alert("โครงสร้างไฟล์ไม่ตรงกับเทมเพลต WL2 Performance: " + e.message);
      return;
    }
    if (!state.kpis.length) {
      alert("ไม่พบข้อมูลรายวันในไฟล์นี้ (ชีต Daily_Input ว่างเปล่า)");
      return;
    }
    state.fileName = fileName;
    state.uploadedAt = new Date();
    state.rangePreset = "all";
    applyPreset("all");

    if (persist !== false) {
      try {
        var b64 = arrayBufferToBase64(arrayBuffer);
        localStorage.setItem(CONFIG.STORAGE_KEYS.FILE, b64);
        localStorage.setItem(CONFIG.STORAGE_KEYS.META, JSON.stringify({ fileName: fileName, uploadedAt: state.uploadedAt.toISOString() }));
      } catch (e) { /* storage quota - ignore, dashboard still works this session */ }
    }
    showDashboard();
    renderAll();
  }

  function showDashboard() {
    $("#emptyState").style.display = "none";
    $("#dashboard").classList.add("visible");
    $("#fileMeta").textContent = state.fileName + " • อัปเดตล่าสุด " + state.uploadedAt.toLocaleString("th-TH");
    $("#fileMeta").style.display = "inline-flex";
    $("#clearBtn").style.display = "inline-flex";
  }

  function tryRestoreFromStorage() {
    var b64 = localStorage.getItem(CONFIG.STORAGE_KEYS.FILE);
    var metaRaw = localStorage.getItem(CONFIG.STORAGE_KEYS.META);
    if (!b64 || !metaRaw) return;
    try {
      var meta = JSON.parse(metaRaw);
      var buf = base64ToArrayBuffer(b64);
      state.uploadedAt = new Date(meta.uploadedAt);
      processWorkbook(buf, meta.fileName, false);
    } catch (e) { /* ignore corrupt storage */ }
  }

  function clearData() {
    localStorage.removeItem(CONFIG.STORAGE_KEYS.FILE);
    localStorage.removeItem(CONFIG.STORAGE_KEYS.META);
    state.kpis = []; state.rawDays = [];
    $("#dashboard").classList.remove("visible");
    $("#emptyState").style.display = "block";
    $("#fileMeta").style.display = "none";
    $("#clearBtn").style.display = "none";
  }

  // ---------------- Filtering ----------------
  function dataMinMax() {
    var dates = state.kpis.map(function (k) { return k.date; });
    return { min: new Date(Math.min.apply(null, dates)), max: new Date(Math.max.apply(null, dates)) };
  }

  function applyPreset(preset) {
    var mm = dataMinMax();
    state.rangePreset = preset;
    if (preset === "all") { state.rangeFrom = mm.min; state.rangeTo = mm.max; }
    else if (preset === "ytd") { state.rangeFrom = new Date(mm.max.getFullYear(), 0, 1); state.rangeTo = mm.max; }
    else if (preset === "month") { state.rangeFrom = new Date(mm.max.getFullYear(), mm.max.getMonth(), 1); state.rangeTo = mm.max; }
    else if (preset === "last30") { var d = new Date(mm.max); d.setDate(d.getDate() - 29); state.rangeFrom = d < mm.min ? mm.min : d; state.rangeTo = mm.max; }
  }

  function filteredKpis() {
    var from = state.rangeFrom, to = state.rangeTo;
    return state.kpis.filter(function (k) { return k.date >= from && k.date <= to; });
  }

  // ---------------- KPI cards ----------------
  function statusForOEE(v, target) {
    if (v === null) return "neutral";
    if (v >= target) return "good";
    if (v >= target * 0.85) return "warning";
    return "bad";
  }
  function statusForThreshold(v, key) {
    var t = THRESH[key];
    if (v === null || v === undefined) return "neutral";
    if (v >= t.good) return "good";
    if (v >= t.warning) return "warning";
    return "bad";
  }
  function deltaBadge(status) {
    return status === "good" ? "good" : status === "bad" ? "bad" : "neutral";
  }

  // For "all data" / "YTD" the workbook's own Yearly_Summary sheet averages the
  // plan rate across all 12 Plan_Monthly rows (including future months still at
  // their default rate) rather than only the months actually filled in. Mirror
  // that exactly so the headline KPI cards tie back to the Excel's own YTD block.
  function ytdAsPeriodSummary() {
    var monthly = WL2.computeMonthlySummary(state.kpis, state.planMonthly, state.assumptions);
    var ytd = WL2.computeYTD(monthly, state.assumptions);
    return {
      days: ytd.days, sumPT: ytd.sumPT, production: ytd.production, rate: ytd.rate,
      MTBF: ytd.MTBF, MTTR: ytd.MTTR,
      w: { UO: ytd.UO, UA: ytd.UA, OA: ytd.OA, MA: ytd.MA, RE: ytd.RE, CU: ytd.CU, OEE: ytd.OEE },
      avgMethod: { UO: ytd.UO, UA: ytd.UA, CU: ytd.CU, OEE: ytd.OEE, OA: ytd.OA, MA: ytd.MA }
    };
  }

  function renderKPICards() {
    var subset = filteredKpis();
    var isYTD = state.rangePreset === "all" || state.rangePreset === "ytd";
    var ps = isYTD ? ytdAsPeriodSummary() : WL2.computePeriodSummary(subset, state.planWeekly, state.assumptions);
    // YTD/All matches the workbook's own Yearly_Summary block, which is
    // weighted-only. Any other range (custom, this month, last 30 days)
    // matches the contractor's own daily-average report style instead —
    // same distinction the workbook's own Period_Summary sheet documents
    // as block "① weighted" vs block "② daily-average, matches the report".
    var m = isYTD ? ps.w : ps.avgMethod;
    var target = state.assumptions.OEE_tgt;
    var wrap = $("#kpiCards");
    wrap.innerHTML = "";
    $("#kpiMethodNote").textContent = isYTD
      ? "คำนวณแบบถ่วงน้ำหนักตามชั่วโมงจริง — ตรงกับชีต Yearly_Summary"
      : "คำนวณแบบเฉลี่ย %รายวัน — ตรงกับรายงานผู้รับเหมา/PPT (ชีต Period_Summary บล็อก②)";

    var usingTrueBD = subset.some(function (k) { return k.usingTrueBreakdownCount; });
    var mtbfSub = usingTrueBD ? "OT / จำนวนครั้ง Breakdown จริง" : "ยังไม่มีข้อมูลจำนวนครั้ง Breakdown (ใช้ค่าประมาณ)";

    var tiles = [
      { label: "OEE", value: m.OEE, isPct: true, sub: "เป้าหมาย " + target + "% • UO × Performance", status: statusForOEE(m.OEE, target) },
      { label: "Operating Availability (OA)", value: m.OA, isPct: true, sub: "AT / ST", status: statusForThreshold(m.OA, "OA") },
      { label: "Mechanical Avail. (MA)", value: m.MA, isPct: true, sub: "OT / (OT+MT)", status: statusForThreshold(m.MA, "MA") },
      { label: "Utilization Avail. (UA)", value: m.UA, isPct: true, sub: "OT / (OT+IT)", status: statusForThreshold(m.UA, "UA") },
      { label: "Working Utilization (UO)", value: m.UO, isPct: true, sub: "WT / (WT+DT)", status: statusForThreshold(m.UO, "UO") },
      { label: "Reliability (RE)", value: m.RE, isPct: true, sub: "AT / (AT+Unplanned MT)", status: statusForThreshold(m.RE, "RE") },
      { label: "Performance (CU)", value: m.CU, isPct: true, sub: "Actual Output / Ideal Output", status: statusForThreshold(m.CU, "CU") },
      { label: "MTBF", value: ps.MTBF, unit: "hr", isNum: true, sub: mtbfSub, status: "neutral" },
      { label: "MTTR", value: ps.MTTR, unit: "hr", isNum: true, sub: "Unplanned MT / จำนวนครั้ง Breakdown", status: "neutral" },
      { label: "Production", value: ps.production, unit: "BCM", isNum: true, sub: "Rate " + fmt(ps.rate) + " BCM/hr", status: "neutral" },
      { label: "Productive Time", value: ps.sumPT, unit: "hr", isNum: true, sub: subset.length + " วันที่มีข้อมูลในช่วงนี้", status: "neutral" }
    ];

    tiles.forEach(function (t) {
      var el = document.createElement("div");
      el.className = "card stat-tile";
      var valueStr = t.isPct ? pct(t.value) : (t.value === null ? "-" : fmt(t.value) + (t.unit ? '<span class="unit">' + t.unit + "</span>" : ""));
      var color = t.status === "good" ? "var(--good)" : t.status === "warning" ? "var(--warning)" : t.status === "bad" ? "var(--critical)" : "var(--text-muted)";
      el.innerHTML =
        '<div class="label">' + t.label + '</div>' +
        '<div class="value">' + valueStr + '</div>' +
        '<div class="delta ' + deltaBadge(t.status) + '"><span class="status-dot" style="background:' + color + '"></span>' + t.sub + '</div>' +
        (t.isPct ? '<div class="meter-track"><div class="meter-fill" style="width:' + Math.max(0, Math.min(100, t.value || 0)) + '%;background:' + color + '"></div></div>' : '');
      wrap.appendChild(el);
    });
  }

  // ---------------- Trend chart ----------------
  function bucketByGranularity(kpis, granularity) {
    if (granularity === "day") {
      return kpis.map(function (k) { return { label: k.date, key: dateKey(k.date), items: [k] }; });
    }
    var groups = {}, order = [];
    kpis.forEach(function (k) {
      var key = granularity === "week" ? k.date.getFullYear() + "-W" + k.week : k.date.getFullYear() + "-" + k.month;
      if (!groups[key]) { groups[key] = []; order.push(key); }
      groups[key].push(k);
    });
    return order.map(function (key) { return { label: key, key: key, items: groups[key] }; });
  }

  function renderTrendChart() {
    var subset = filteredKpis();
    var buckets = bucketByGranularity(subset, state.granularity);
    var xLabels = buckets.map(function (b) { return b.key; });

    // Use the granularity-appropriate reference so each point matches its
    // validated source: a single day already carries its own OEE (monthly
    // plan rate, matches Calc_Daily); a week aggregate matches Weekly_Summary
    // (weekly plan rate); a month aggregate matches Monthly_Summary (monthly
    // plan rate).
    var oeeVals = [], prodVals = [];
    buckets.forEach(function (b) {
      var oee, production;
      if (state.granularity === "day") {
        oee = b.items[0].OEE;
        production = b.items[0].S;
      } else if (state.granularity === "week") {
        var ws = WL2.summarizeWeek(b.items, b.items[0].week, state.planWeekly);
        oee = ws.OEE;
        production = ws.production;
      } else {
        var ms = WL2.summarizeMonth(b.items, b.items[0].month, state.planMonthly, state.assumptions);
        oee = ms.w.OEE;
        production = ms.production;
      }
      oeeVals.push(oee);
      prodVals.push(production);
    });

    var xFormat = function (label, i) {
      if (state.granularity === "day") return shortDate(buckets[i].items[0].date);
      if (state.granularity === "week") return "W" + buckets[i].items[0].week;
      return MONTH_NAMES[buckets[i].items[0].month - 1];
    };

    lineChart($("#trendOEEChart"), {
      height: 260,
      series: [
        { name: "OEE %", color: "var(--series-1)", area: true, values: oeeVals.map(function (v) { return { y: v }; }) }
      ],
      xLabels: xLabels, xFormat: xFormat,
      targetLine: { value: state.assumptions.OEE_tgt, label: "เป้าหมาย " + state.assumptions.OEE_tgt + "%" },
      yFormat: function (v) { return v + "%"; }, yMax: 100
    });
    renderChartTable("trendOEETable", ["ช่วง", "OEE %"], xLabels.map(function (l, i) { return [xFormat(l, i), pct(oeeVals[i])]; }));

    barChart($("#trendProdChart"), {
      height: 240, mode: "grouped",
      categories: xLabels,
      series: [{ name: "Production", color: "var(--series-1)", values: prodVals }],
      xFormat: xFormat, yFormat: function (v) { return fmt(v); }
    });
    renderChartTable("trendProdTable", ["ช่วง", "Production (BCM)"], xLabels.map(function (l, i) { return [xFormat(l, i), fmt(prodVals[i])]; }));
  }

  // ---------------- Monthly Actual vs Plan ----------------
  function renderMonthlySection() {
    var monthly = WL2.computeMonthlySummary(state.kpis, state.planMonthly, state.assumptions);
    var activeMonths = monthly.filter(function (m) { return m.production > 0 || m.planProd > 0; });
    var cats = activeMonths.map(function (m) { return m.name; });

    barChart($("#monthlyProdChart"), {
      height: 260, mode: "grouped",
      categories: cats,
      series: [
        { name: "Actual", color: "var(--series-1)", values: activeMonths.map(function (m) { return m.production; }) },
        { name: "Plan", color: "var(--series-3)", values: activeMonths.map(function (m) { return m.planProd; }) }
      ],
      yFormat: function (v) { return fmt(v); }
    });
    renderLegend("monthlyProdLegend", [{ name: "Actual", color: "var(--series-1)" }, { name: "Plan", color: "var(--series-3)" }]);

    var methodKey = state.monthlyMethod === "weighted" ? "w" : "avgMethod";
    lineChart($("#monthlyOEEChart"), {
      height: 240,
      series: [{ name: "OEE %", color: "var(--series-1)", area: true, values: activeMonths.map(function (m) { return { y: m[methodKey].OEE }; }) }],
      xLabels: cats,
      targetLine: { value: state.assumptions.OEE_tgt, label: "เป้าหมาย" }, yFormat: function (v) { return v + "%"; }, yMax: 100
    });

    var rowsHtml = activeMonths.map(function (m) {
      var mk = m[methodKey];
      var varClass = m.varProd === null ? "" : (m.varProd >= 0 ? "good" : "critical");
      return "<tr>" +
        "<td>" + m.name + "</td>" +
        '<td class="num">' + fmt(m.production) + "</td>" +
        '<td class="num">' + fmt(m.planProd) + "</td>" +
        '<td class="num"><span class="badge ' + varClass + '">' + (m.varProd === null ? "-" : (m.varProd >= 0 ? "+" : "") + fmt(m.varProd)) + "</span></td>" +
        '<td class="num">' + fmt(m.rate) + "</td>" +
        '<td class="num">' + fmt(m.planRate) + "</td>" +
        '<td class="num">' + pct(mk.OA) + "</td>" +
        '<td class="num">' + pct(mk.UO) + "</td>" +
        '<td class="num">' + pct(mk.CU) + "</td>" +
        '<td class="num">' + pct(mk.OEE) + "</td>" +
        "</tr>";
    }).join("");
    $("#monthlyTableBody").innerHTML = rowsHtml;
  }

  // ---------------- Time-loss breakdown ----------------
  function renderTimeLossSection() {
    var subset = filteredKpis();
    var buckets = bucketByGranularity(subset, state.granularity === "day" && subset.length > 60 ? "week" : state.granularity);
    var xFormat = function (label, i) {
      var first = buckets[i].items[0];
      if (buckets[i].items.length === 1) return shortDate(first.date);
      return "W" + first.week;
    };
    function sumField(items, f) { return items.reduce(function (a, k) { return a + k[f]; }, 0); }

    barChart($("#timeLossChart"), {
      height: 280, mode: "stacked",
      categories: buckets.map(function (b) { return b.key; }),
      xFormat: xFormat,
      series: [
        { name: "Productive (PT)", color: "var(--series-1)", values: buckets.map(function (b) { return sumField(b.items, "PT"); }) },
        { name: "Maintenance (MT)", color: "var(--series-6)", values: buckets.map(function (b) { return sumField(b.items, "MT"); }) },
        { name: "Outage+Reloc (PA)", color: "var(--series-5)", values: buckets.map(function (b) { return sumField(b.items, "PA"); }) },
        { name: "Idle (IT)", color: "var(--series-3)", values: buckets.map(function (b) { return sumField(b.items, "IT"); }) },
        { name: "Delay (DT)", color: "var(--series-8)", values: buckets.map(function (b) { return sumField(b.items, "DT"); }) }
      ],
      yFormat: function (v) { return fmt(v) + "h"; }
    });
    renderLegend("timeLossLegend", [
      { name: "Productive (PT)", color: "var(--series-1)" }, { name: "Maintenance (MT)", color: "var(--series-6)" },
      { name: "Outage+Reloc (PA)", color: "var(--series-5)" }, { name: "Idle (IT)", color: "var(--series-3)" },
      { name: "Delay (DT)", color: "var(--series-8)" }
    ]);

    // Pareto of detailed causes across the filtered subset
    var causeTotals = [];
    CAUSE_GROUPS.forEach(function (grp) {
      grp.fields.forEach(function (fld) {
        var total = subset.reduce(function (a, k) { return a + (k.raw[fld[0]] || 0); }, 0);
        if (total > 0) causeTotals.push({ label: fld[1], value: total, color: grp.color, group: grp.key });
      });
    });
    Charts.paretoChart($("#paretoChart"), {
      items: causeTotals, limit: 12, labelWidth: 190, rowH: 26,
      valueFormat: function (v) { return fmt(v) + " hr"; }
    });
    var groupLegend = CAUSE_GROUPS.map(function (g) { return { name: g.key, color: g.color }; });
    renderLegend("paretoLegend", groupLegend);
  }

  // ---------------- Contractor production ----------------
  function renderContractorSection() {
    var subset = filteredKpis();
    var buckets = bucketByGranularity(subset, state.granularity);
    function sumField(items, f) { return items.reduce(function (a, k) { return a + k[f]; }, 0); }
    var xFormat = function (label, i) {
      var first = buckets[i].items[0];
      if (state.granularity === "day") return shortDate(first.date);
      if (state.granularity === "week") return "W" + first.week;
      return MONTH_NAMES[first.month - 1];
    };
    barChart($("#contractorChart"), {
      height: 260, mode: "stacked",
      categories: buckets.map(function (b) { return b.key; }),
      xFormat: xFormat,
      series: [
        { name: CONTRACTOR_LABELS.S1, color: "var(--series-1)", values: buckets.map(function (b) { return sumField(b.items, "S1"); }) },
        { name: CONTRACTOR_LABELS.S2, color: "var(--series-2)", values: buckets.map(function (b) { return sumField(b.items, "S2"); }) },
        { name: CONTRACTOR_LABELS.S3, color: "var(--series-3)", values: buckets.map(function (b) { return sumField(b.items, "S3"); }) }
      ],
      yFormat: function (v) { return fmt(v); }
    });
    renderLegend("contractorLegend", [
      { name: CONTRACTOR_LABELS.S1, color: "var(--series-1)" }, { name: CONTRACTOR_LABELS.S2, color: "var(--series-2)" }, { name: CONTRACTOR_LABELS.S3, color: "var(--series-3)" }
    ]);

    var totalS1 = subset.reduce(function (a, k) { return a + k.S1; }, 0);
    var totalS2 = subset.reduce(function (a, k) { return a + k.S2; }, 0);
    var totalS3 = subset.reduce(function (a, k) { return a + k.S3; }, 0);
    var total = totalS1 + totalS2 + totalS3;
    var rows = [
      [CONTRACTOR_LABELS.S1, totalS1], [CONTRACTOR_LABELS.S2, totalS2], [CONTRACTOR_LABELS.S3, totalS3]
    ];
    $("#contractorTableBody").innerHTML = rows.map(function (r) {
      var share = total ? (r[1] / total * 100) : 0;
      return "<tr><td>" + r[0] + "</td><td class='num'>" + fmt(r[1]) + "</td><td class='num'>" + share.toFixed(1) + "%</td></tr>";
    }).join("");
  }

  // ---------------- Crusher performance ----------------
  function renderCrusherSection() {
    var subset = filteredKpis();
    var wrap = $("#crusherGrid");
    wrap.innerHTML = "";
    CRUSHER_DEFS.forEach(function (cr, idx) {
      var prod = subset.reduce(function (a, k) { return a + k.crusherStats[cr.key].prod; }, 0);
      var hrs = subset.reduce(function (a, k) { return a + k.crusherStats[cr.key].hrs; }, 0);
      var rate = hrs > 0 ? prod / hrs : null;
      var util = cr.capacity > 0 && rate !== null ? (rate / cr.capacity * 100) : null;
      var sparkVals = subset.map(function (k) { return k.crusherStats[cr.key].prod; });
      var div = document.createElement("div");
      div.className = "card small-mult";
      var color = "var(--series-" + ((idx % 8) + 1) + ")";
      div.innerHTML =
        "<h4>" + cr.label + "</h4>" +
        '<div class="kpi-line"><span>Production</span><b class="num">' + fmt(prod) + " BCM</b></div>" +
        '<div class="kpi-line"><span>Hours run</span><b class="num">' + fmt(hrs) + " hr</b></div>" +
        '<div class="kpi-line"><span>Avg rate</span><b class="num">' + (rate === null ? "-" : fmt(rate) + " BCM/hr") + '</b></div>' +
        '<div class="kpi-line"><span>vs capacity</span><b class="num">' + (util === null ? "n/a" : util.toFixed(0) + "%") + '</b></div>' +
        '<div class="spark" style="margin-top:6px"></div>';
      wrap.appendChild(div);
      Charts.sparkline(div.querySelector(".spark"), sparkVals, { width: 220, height: 32, color: color });
    });
  }

  // ---------------- Truck fleet ----------------
  function renderFleetSection() {
    var subset = filteredKpis();
    var TRIP_FIELD = { HD785: "hd785Trips", Tonly: "tonlyTrips", "12DT-D": "dt12dTrips", "10DT": "dt10Trips", "12DT-F": "dt12fTrips", Tianma: "tianmaTrips" };
    var rows = FLEET_DEFS.map(function (f) {
      var prod = subset.reduce(function (a, k) { return a + (k.fleetProd[f.key] || 0); }, 0);
      var tripsTotal = subset.reduce(function (a, k) { return a + k.raw[TRIP_FIELD[f.key]]; }, 0);
      var avgRate = tripsTotal > 0 ? prod / tripsTotal : null;
      return { label: f.label, trips: tripsTotal, prod: prod, avgRate: avgRate };
    });
    var totalProd = rows.reduce(function (a, r) { return a + r.prod; }, 0);
    $("#fleetTableBody").innerHTML = rows.map(function (r) {
      var share = totalProd ? (r.prod / totalProd * 100) : 0;
      return "<tr><td>" + r.label + "</td><td class='num'>" + fmt(r.trips) + "</td><td class='num'>" + (r.avgRate === null ? "-" : fmt(r.avgRate, 1)) + "</td><td class='num'>" + fmt(r.prod) + "</td><td class='num'>" + share.toFixed(1) + "%</td></tr>";
    }).join("");

    Charts.paretoChart($("#fleetChart"), {
      items: rows.filter(function (r) { return r.prod > 0; }).map(function (r, i) { return { label: r.label, value: r.prod, color: "var(--series-" + ((i % 8) + 1) + ")" }; }),
      labelWidth: 130, rowH: 28, valueFormat: function (v) { return fmt(v) + " BCM"; }
    });
  }

  // ---------------- Data quality ----------------
  function renderDataQuality() {
    var subset = state.dataQuality.filter(function (d) { return d.date >= state.rangeFrom && d.date <= state.rangeTo; });
    var negative = subset.filter(function (d) { return d.negative; }).length;
    var over24 = subset.filter(function (d) { return d.over24h; }).length;
    var noProd = subset.filter(function (d) { return d.noProduction; }).length;
    var dup = subset.filter(function (d) { return d.duplicate; }).length;
    var missingBD = subset.filter(function (d) { return d.missingBreakdownCount; }).length;
    var check = subset.filter(function (d) { return d.status !== "PASS"; }).length;

    var chips = [
      { label: "วันที่ตรวจสอบ", value: subset.length, cls: "" },
      { label: "PASS", value: subset.length - check, cls: "good" },
      { label: "ต้องแก้ (CHECK)", value: check, cls: check ? "critical" : "good" },
      { label: "ไม่มี Production", value: noProd, cls: noProd ? "warning" : "good" },
      { label: "ไม่มีจำนวนครั้ง Breakdown", value: missingBD, cls: missingBD ? "warning" : "good" },
      { label: "ค่าติดลบ", value: negative, cls: negative ? "critical" : "good" },
      { label: "เกิน 24 ชม./วัน", value: over24, cls: over24 ? "critical" : "good" },
      { label: "วันซ้ำ", value: dup, cls: dup ? "critical" : "good" }
    ];
    $("#dqChips").innerHTML = chips.map(function (c) {
      var color = c.cls === "good" ? "var(--good)" : c.cls === "critical" ? "var(--critical)" : c.cls === "warning" ? "var(--warning)" : "var(--text-muted)";
      return '<span class="dq-chip" style="border-color:' + color + '55;color:' + (c.cls ? color : "var(--text-primary)") + '">' + c.label + ": " + c.value + "</span>";
    }).join("");

    var issues = subset.filter(function (d) { return d.status !== "PASS" || d.noProduction || d.missingBreakdownCount; });
    var tbody = $("#dqTableBody");
    if (!issues.length) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text-muted);padding:20px">ไม่มีปัญหาข้อมูลในช่วงที่เลือก — ข้อมูลผ่านการตรวจสอบทั้งหมด</td></tr>';
    } else {
      tbody.innerHTML = issues.map(function (d) {
        return "<tr><td>" + d.date.toLocaleDateString("th-TH") + "</td>" +
          "<td>" + (d.missingSeason ? "⚠" : "-") + "</td>" +
          "<td>" + (d.negative ? "⚠" : "-") + "</td>" +
          "<td>" + (d.over24h ? "⚠" : "-") + "</td>" +
          "<td>" + (d.noProduction ? "⚠" : "-") + "</td>" +
          "<td>" + (d.missingBreakdownCount ? "⚠" : "-") + "</td>" +
          "<td>" + (d.duplicate ? "⚠" : "-") + "</td>" +
          '<td><span class="badge ' + (d.status === "PASS" ? "good" : "critical") + '">' + d.status + "</span></td></tr>";
      }).join("");
    }
  }

  // ---------------- Raw data table ----------------
  var tableState = { sortKey: "date", sortDir: -1, search: "" };
  function renderRawTable() {
    var subset = filteredKpis().slice();
    if (tableState.search) {
      var s = tableState.search.toLowerCase();
      subset = subset.filter(function (k) { return (k.season || "").toLowerCase().indexOf(s) >= 0 || shortDate(k.date).indexOf(s) >= 0; });
    }
    subset.sort(function (a, b) {
      var av = tableState.sortKey === "date" ? a.date : a[tableState.sortKey];
      var bv = tableState.sortKey === "date" ? b.date : b[tableState.sortKey];
      if (av === null) av = -Infinity; if (bv === null) bv = -Infinity;
      return av > bv ? tableState.sortDir : av < bv ? -tableState.sortDir : 0;
    });
    var rows = subset.slice(0, 500);
    $("#rawTableBody").innerHTML = rows.map(function (k) {
      return "<tr><td>" + k.date.toLocaleDateString("th-TH") + "</td>" +
        "<td class='num'>" + k.week + "</td>" +
        "<td>" + (k.season || "-") + "</td>" +
        "<td class='num'>" + fmt(k.ST) + "</td>" +
        "<td class='num'>" + fmt(k.AT) + "</td>" +
        "<td class='num'>" + fmt(k.OT) + "</td>" +
        "<td class='num'>" + fmt(k.PT) + "</td>" +
        "<td class='num'>" + fmt(k.S) + "</td>" +
        "<td class='num'>" + (k.Qvs === null ? "-" : fmt(k.Qvs)) + "</td>" +
        "<td class='num'>" + pct(k.OA) + "</td>" +
        "<td class='num'>" + pct(k.UO) + "</td>" +
        "<td class='num'>" + pct(k.CU) + "</td>" +
        "<td class='num'>" + pct(k.OEE) + "</td></tr>";
    }).join("");
    $("#rawTableCount").textContent = subset.length + " วัน" + (subset.length > 500 ? " (แสดง 500 แถวแรก)" : "");
  }

  function wireTableSort() {
    $all("#rawTableHead th[data-key]").forEach(function (th) {
      th.addEventListener("click", function () {
        var key = th.getAttribute("data-key");
        if (tableState.sortKey === key) tableState.sortDir *= -1; else { tableState.sortKey = key; tableState.sortDir = 1; }
        renderRawTable();
      });
    });
    $("#tableSearch").addEventListener("input", function (e) { tableState.search = e.target.value; renderRawTable(); });
    $("#exportCsvBtn").addEventListener("click", exportCsv);
  }

  function exportCsv() {
    var subset = filteredKpis();
    var header = ["Date", "Week", "Season", "ST", "AT", "OT", "PT", "Production_BCM", "Rate_BCM_hr", "OA_pct", "UO_pct", "CU_pct", "OEE_pct"];
    var lines = [header.join(",")];
    subset.forEach(function (k) {
      lines.push([dateKey(k.date), k.week, k.season, k.ST, k.AT, k.OT, k.PT, k.S, k.Qvs || "", k.OA || "", k.UO || "", k.CU || "", k.OEE || ""].join(","));
    });
    var blob = new Blob([lines.join("\n")], { type: "text/csv" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "wl2_performance_export_" + dateKey(state.rangeFrom) + "_to_" + dateKey(state.rangeTo) + ".csv";
    a.click();
  }

  // ---------------- Chart helpers: legend + table-view toggle ----------------
  function renderLegend(containerId, items) {
    var el = document.getElementById(containerId);
    if (!el) return;
    el.innerHTML = items.map(function (it) {
      return '<span class="item"><span class="swatch" style="background:' + it.color + '"></span>' + it.name + '</span>';
    }).join("");
  }

  function renderChartTable(bodyId, headers, rows) {
    var el = document.getElementById(bodyId);
    if (!el) return;
    el.innerHTML = "<table><thead><tr>" + headers.map(function (h) { return "<th>" + h + "</th>"; }).join("") + "</tr></thead><tbody>" +
      rows.map(function (r) { return "<tr>" + r.map(function (c) { return "<td class='num'>" + c + "</td>"; }).join("") + "</tr>"; }).join("") +
      "</tbody></table>";
  }

  function wireChartTableToggles() {
    $all(".chart-toggle-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var targetId = btn.getAttribute("data-target");
        var target = document.getElementById(targetId);
        target.classList.toggle("visible");
        btn.textContent = target.classList.contains("visible") ? "ซ่อนตาราง" : "แสดงตาราง";
      });
    });
  }

  // ---------------- Filter bar wiring ----------------
  function populateMonthSelect() {
    var months = {};
    state.kpis.forEach(function (k) { months[k.month] = true; });
    var sel = $("#monthSelect");
    sel.innerHTML = '<option value="">— เลือกเดือน —</option>' + Object.keys(months).sort(function (a, b) { return a - b; }).map(function (m) {
      return '<option value="' + m + '">' + MONTH_NAMES[m - 1] + " 2026</option>";
    }).join("");
  }

  function wireFilterBar() {
    $all("#granularitySeg button").forEach(function (btn) {
      btn.addEventListener("click", function () {
        $all("#granularitySeg button").forEach(function (b) { b.classList.remove("active"); });
        btn.classList.add("active");
        state.granularity = btn.getAttribute("data-val");
        renderAll();
      });
    });
    $all("#presetSeg button").forEach(function (btn) {
      btn.addEventListener("click", function () {
        $all("#presetSeg button").forEach(function (b) { b.classList.remove("active"); });
        btn.classList.add("active");
        applyPreset(btn.getAttribute("data-val"));
        $("#monthSelect").value = "";
        $("#fromDate").value = dateKey(state.rangeFrom);
        $("#toDate").value = dateKey(state.rangeTo);
        renderAll();
      });
    });
    $("#monthSelect").addEventListener("change", function (e) {
      var m = Number(e.target.value);
      if (!m) return;
      var year = dataMinMax().max.getFullYear();
      state.rangeFrom = new Date(year, m - 1, 1);
      state.rangeTo = new Date(year, m, 0);
      state.rangePreset = "custom";
      $all("#presetSeg button").forEach(function (b) { b.classList.remove("active"); });
      $("#fromDate").value = dateKey(state.rangeFrom);
      $("#toDate").value = dateKey(state.rangeTo);
      renderAll();
    });
    $("#fromDate").addEventListener("change", function (e) {
      state.rangeFrom = new Date(e.target.value + "T00:00:00");
      state.rangePreset = "custom";
      $all("#presetSeg button").forEach(function (b) { b.classList.remove("active"); });
      renderAll();
    });
    $("#toDate").addEventListener("change", function (e) {
      state.rangeTo = new Date(e.target.value + "T00:00:00");
      state.rangePreset = "custom";
      $all("#presetSeg button").forEach(function (b) { b.classList.remove("active"); });
      renderAll();
    });
    $all("#methodSeg button").forEach(function (btn) {
      btn.addEventListener("click", function () {
        $all("#methodSeg button").forEach(function (b) { b.classList.remove("active"); });
        btn.classList.add("active");
        state.monthlyMethod = btn.getAttribute("data-val");
        renderMonthlySection();
      });
    });
  }

  function renderAll() {
    populateMonthSelect();
    renderKPICards();
    renderTrendChart();
    renderMonthlySection();
    renderTimeLossSection();
    renderContractorSection();
    renderCrusherSection();
    renderFleetSection();
    renderDataQuality();
    renderRawTable();
  }

  // ---------------- Upload wiring ----------------
  function wireUpload() {
    var input = $("#fileInput");
    $("#uploadBtn").addEventListener("click", function () { input.click(); });
    $("#uploadBtnEmpty").addEventListener("click", function () { input.click(); });
    input.addEventListener("change", function (e) {
      var file = e.target.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function () { processWorkbook(reader.result, file.name); };
      reader.readAsArrayBuffer(file);
      input.value = "";
    });
    var dz = $("#emptyState");
    ["dragenter", "dragover"].forEach(function (evt) {
      dz.addEventListener(evt, function (e) { e.preventDefault(); dz.classList.add("dragover"); });
    });
    ["dragleave", "drop"].forEach(function (evt) {
      dz.addEventListener(evt, function (e) { e.preventDefault(); dz.classList.remove("dragover"); });
    });
    dz.addEventListener("drop", function (e) {
      var file = e.dataTransfer.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function () { processWorkbook(reader.result, file.name); };
      reader.readAsArrayBuffer(file);
    });
    $("#clearBtn").addEventListener("click", function () {
      if (confirm("ล้างข้อมูลที่บันทึกไว้ในเบราว์เซอร์นี้? (ไฟล์ Excel ต้นฉบับจะไม่ถูกลบ)")) clearData();
    });
  }

  function init() {
    initTheme();
    wireUpload();
    wireFilterBar();
    wireTableSort();
    wireChartTableToggles();
    tryRestoreFromStorage();
  }

  document.addEventListener("DOMContentLoaded", init);
})();
