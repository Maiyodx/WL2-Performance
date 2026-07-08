/* Minimal SVG chart toolkit: line/area, grouped+stacked bar, sparkline.
   Design follows the dataviz skill: thin 2px lines, rounded caps, hairline
   gridlines, hover crosshair + tooltip, legend with toggle-to-isolate,
   table-view twin for every chart. No external dependencies. */
(function (root) {
  "use strict";

  var NS = "http://www.w3.org/2000/svg";
  function el(tag, attrs) {
    var e = document.createElementNS(NS, tag);
    if (attrs) for (var k in attrs) e.setAttribute(k, attrs[k]);
    return e;
  }
  function fmtNum(v, digits) {
    if (v === null || v === undefined || isNaN(v)) return "-";
    if (Math.abs(v) >= 1000) return Math.round(v).toLocaleString("en-US");
    return v.toFixed(digits === undefined ? 1 : digits);
  }

  function niceTicks(min, max, count) {
    if (min === max) { min -= 1; max += 1; }
    var span = max - min;
    var step = Math.pow(10, Math.floor(Math.log10(span / count)));
    var err = (span / count) / step;
    if (err >= 7.5) step *= 10; else if (err >= 3.5) step *= 5; else if (err >= 1.5) step *= 2;
    var niceMin = Math.floor(min / step) * step;
    var niceMax = Math.ceil(max / step) * step;
    var ticks = [];
    for (var v = niceMin; v <= niceMax + step * 1e-6; v += step) ticks.push(Math.round(v * 1e6) / 1e6);
    return ticks;
  }

  function ensureTooltip(container) {
    var tip = container.querySelector(".tooltip");
    if (!tip) {
      tip = document.createElement("div");
      tip.className = "tooltip";
      container.appendChild(tip);
    }
    return tip;
  }

  function showTooltip(container, tip, x, y, html) {
    tip.innerHTML = html;
    tip.style.left = x + "px";
    tip.style.top = y + "px";
    tip.classList.add("show");
  }
  function hideTooltip(tip) { tip.classList.remove("show"); }

  /**
   * Multi-series line/area chart.
   * opts: { width, height, margin, series:[{name,color,values:[{x,y}], area:bool}],
   *         xLabels:[...], yFormat, xFormat, targetLine:{value,label}, yTitle }
   */
  function lineChart(container, opts) {
    container.innerHTML = "";
    var W = opts.width || container.clientWidth || 640;
    var H = opts.height || 260;
    var margin = Object.assign({ top: 10, right: 16, bottom: 26, left: 46 }, opts.margin || {});
    var innerW = W - margin.left - margin.right;
    var innerH = H - margin.top - margin.bottom;

    var svg = el("svg", { viewBox: "0 0 " + W + " " + H, width: "100%", height: H });
    var series = opts.series.filter(function (s) { return s.values && s.values.length; });
    var n = opts.xLabels.length;

    var allY = [];
    series.forEach(function (s) { s.values.forEach(function (v) { if (v.y !== null && v.y !== undefined) allY.push(v.y); }); });
    if (opts.targetLine) allY.push(opts.targetLine.value);
    if (opts.yMin !== undefined) allY.push(opts.yMin);
    if (opts.yMax !== undefined) allY.push(opts.yMax);
    var yMin = allY.length ? Math.min.apply(null, allY) : 0;
    var yMax = allY.length ? Math.max.apply(null, allY) : 1;
    if (opts.forceZero !== false) yMin = Math.min(0, yMin);
    var ticks = niceTicks(yMin, yMax, 4);
    yMin = ticks[0]; yMax = ticks[ticks.length - 1];

    function xScale(i) { return n <= 1 ? margin.left + innerW / 2 : margin.left + (innerW * i) / (n - 1); }
    function yScale(v) { return margin.top + innerH - ((v - yMin) / (yMax - yMin || 1)) * innerH; }

    var gGrid = el("g");
    ticks.forEach(function (t) {
      var y = yScale(t);
      gGrid.appendChild(el("line", { x1: margin.left, x2: margin.left + innerW, y1: y, y2: y, stroke: "var(--gridline)", "stroke-width": 1 }));
      var lbl = el("text", { x: margin.left - 8, y: y + 4, "text-anchor": "end", "font-size": 10.5, fill: "var(--text-muted)" });
      lbl.textContent = opts.yFormat ? opts.yFormat(t) : fmtNum(t);
      gGrid.appendChild(lbl);
    });
    svg.appendChild(gGrid);

    // x-axis labels (subsample to avoid overlap)
    var maxLabels = Math.max(4, Math.floor(innerW / 70));
    var stride = Math.max(1, Math.ceil(n / maxLabels));
    var gX = el("g");
    for (var i = 0; i < n; i += stride) {
      var x = xScale(i);
      var t = el("text", { x: x, y: H - 6, "text-anchor": "middle", "font-size": 10.5, fill: "var(--text-muted)" });
      t.textContent = opts.xFormat ? opts.xFormat(opts.xLabels[i], i) : opts.xLabels[i];
      gX.appendChild(t);
    }
    svg.appendChild(gX);
    svg.appendChild(el("line", { x1: margin.left, x2: margin.left + innerW, y1: margin.top + innerH, y2: margin.top + innerH, stroke: "var(--baseline)", "stroke-width": 1 }));

    if (opts.targetLine) {
      var ty = yScale(opts.targetLine.value);
      var tl = el("line", { x1: margin.left, x2: margin.left + innerW, y1: ty, y2: ty, stroke: "var(--text-muted)", "stroke-width": 1.25, "stroke-dasharray": "3,3" });
      svg.appendChild(tl);
      var tlabel = el("text", { x: margin.left + innerW, y: ty - 4, "text-anchor": "end", "font-size": 10, fill: "var(--text-muted)", "font-weight": 700 });
      tlabel.textContent = opts.targetLine.label || fmtNum(opts.targetLine.value);
      svg.appendChild(tlabel);
    }

    function pathFor(values, defined) {
      var d = "", started = false;
      for (var i = 0; i < values.length; i++) {
        var v = values[i];
        if (v.y === null || v.y === undefined) { started = false; continue; }
        var x = xScale(i), y = yScale(v.y);
        d += (started ? " L " : " M ") + x.toFixed(2) + " " + y.toFixed(2);
        started = true;
      }
      return d;
    }
    function areaFor(values) {
      var segs = [], cur = null;
      values.forEach(function (v, i) {
        if (v.y === null || v.y === undefined) { if (cur) { segs.push(cur); cur = null; } return; }
        if (!cur) cur = [];
        cur.push({ i: i, y: v.y });
      });
      if (cur) segs.push(cur);
      return segs.map(function (seg) {
        var d = "M " + xScale(seg[0].i) + " " + yScale(seg[0].y);
        seg.forEach(function (p, idx) { if (idx) d += " L " + xScale(p.i) + " " + yScale(p.y); });
        d += " L " + xScale(seg[seg.length - 1].i) + " " + yScale(yMin);
        d += " L " + xScale(seg[0].i) + " " + yScale(yMin) + " Z";
        return d;
      }).join(" ");
    }

    var seriesGroups = [];
    series.forEach(function (s) {
      var g = el("g", { "data-series": s.name });
      if (s.area) {
        var ap = areaFor(s.values);
        if (ap) g.appendChild(el("path", { d: ap, fill: s.color, opacity: 0.14, stroke: "none" }));
      }
      var path = pathFor(s.values);
      if (path) g.appendChild(el("path", { d: path, fill: "none", stroke: s.color, "stroke-width": 2, "stroke-linecap": "round", "stroke-linejoin": "round" }));
      svg.appendChild(g);
      seriesGroups.push(g);
    });

    var hoverLine = el("line", { y1: margin.top, y2: margin.top + innerH, stroke: "var(--text-muted)", "stroke-width": 1, opacity: 0 });
    svg.appendChild(hoverLine);
    var hoverDots = series.map(function (s) { return el("circle", { r: 3.5, fill: s.color, stroke: "var(--surface-1)", "stroke-width": 1.5, opacity: 0 }); });
    hoverDots.forEach(function (d) { svg.appendChild(d); });

    var hitRect = el("rect", { x: margin.left, y: margin.top, width: Math.max(innerW, 1), height: Math.max(innerH, 1), fill: "transparent" });
    svg.appendChild(hitRect);

    var wrap = document.createElement("div");
    wrap.className = "chart-body";
    wrap.style.position = "relative";
    wrap.appendChild(svg);
    var tip = ensureTooltip(wrap);
    container.appendChild(wrap);

    hitRect.addEventListener("mousemove", function (evt) {
      var rect = svg.getBoundingClientRect();
      var scaleX = W / rect.width;
      var mx = (evt.clientX - rect.left) * scaleX;
      var idx = n <= 1 ? 0 : Math.round(((mx - margin.left) / innerW) * (n - 1));
      idx = Math.max(0, Math.min(n - 1, idx));
      var x = xScale(idx);
      hoverLine.setAttribute("x1", x); hoverLine.setAttribute("x2", x); hoverLine.setAttribute("opacity", 1);
      var rows = "";
      series.forEach(function (s, si) {
        var v = s.values[idx];
        var y = v && v.y !== null && v.y !== undefined ? yScale(v.y) : null;
        if (y !== null) { hoverDots[si].setAttribute("cx", x); hoverDots[si].setAttribute("cy", y); hoverDots[si].setAttribute("opacity", 1); }
        else hoverDots[si].setAttribute("opacity", 0);
        rows += '<div class="t-row"><span class="t-swatch" style="background:' + s.color + '"></span>' + s.name + ": <b>" + (v && v.y !== null && v.y !== undefined ? (opts.yFormat ? opts.yFormat(v.y) : fmtNum(v.y)) : "-") + "</b></div>";
      });
      var scaleY = H / rect.height;
      var px = evt.clientX - rect.left, py = evt.clientY - rect.top;
      showTooltip(wrap, tip, x / scaleX, (margin.top + 6) / scaleY, '<div class="t-title">' + (opts.xFormat ? opts.xFormat(opts.xLabels[idx], idx) : opts.xLabels[idx]) + "</div>" + rows);
    });
    hitRect.addEventListener("mouseleave", function () {
      hoverLine.setAttribute("opacity", 0);
      hoverDots.forEach(function (d) { d.setAttribute("opacity", 0); });
      hideTooltip(tip);
    });

    return { svg: svg, seriesGroups: seriesGroups };
  }

  /**
   * Grouped or stacked bar chart.
   * opts: { width,height,margin, categories:[...], series:[{name,color,values:[...]}],
   *         mode: 'grouped'|'stacked', yFormat, xFormat, stackAsPercent }
   */
  function barChart(container, opts) {
    container.innerHTML = "";
    var W = opts.width || container.clientWidth || 640;
    var H = opts.height || 260;
    var margin = Object.assign({ top: 10, right: 16, bottom: 30, left: 50 }, opts.margin || {});
    var innerW = W - margin.left - margin.right;
    var innerH = H - margin.top - margin.bottom;
    var mode = opts.mode || "grouped";
    var cats = opts.categories;
    var series = opts.series;
    var n = cats.length;

    var svg = el("svg", { viewBox: "0 0 " + W + " " + H, width: "100%", height: H });

    var yMax, yMin = 0;
    if (mode === "stacked") {
      var totalsPos = cats.map(function (_, i) { return series.reduce(function (a, s) { return a + Math.max(0, s.values[i] || 0); }, 0); });
      var totalsNeg = cats.map(function (_, i) { return series.reduce(function (a, s) { return a + Math.min(0, s.values[i] || 0); }, 0); });
      yMax = Math.max.apply(null, totalsPos.concat(0));
      yMin = Math.min.apply(null, totalsNeg.concat(0));
    } else {
      var allVals = [];
      series.forEach(function (s) { s.values.forEach(function (v) { allVals.push(v || 0); }); });
      yMax = allVals.length ? Math.max.apply(null, allVals) : 1;
      yMin = Math.min(0, allVals.length ? Math.min.apply(null, allVals) : 0);
    }
    var ticks = niceTicks(yMin, yMax, 4);
    yMin = ticks[0]; yMax = ticks[ticks.length - 1];
    function yScale(v) { return margin.top + innerH - ((v - yMin) / (yMax - yMin || 1)) * innerH; }
    var y0 = yScale(0);

    var gGrid = el("g");
    ticks.forEach(function (t) {
      var y = yScale(t);
      gGrid.appendChild(el("line", { x1: margin.left, x2: margin.left + innerW, y1: y, y2: y, stroke: "var(--gridline)", "stroke-width": 1 }));
      var lbl = el("text", { x: margin.left - 8, y: y + 4, "text-anchor": "end", "font-size": 10.5, fill: "var(--text-muted)" });
      lbl.textContent = opts.yFormat ? opts.yFormat(t) : fmtNum(t);
      gGrid.appendChild(lbl);
    });
    svg.appendChild(gGrid);
    svg.appendChild(el("line", { x1: margin.left, x2: margin.left + innerW, y1: y0, y2: y0, stroke: "var(--baseline)", "stroke-width": 1 }));

    var bandW = innerW / n;
    var maxLabels = Math.max(3, Math.floor(innerW / 60));
    var stride = Math.max(1, Math.ceil(n / maxLabels));
    var gX = el("g");
    for (var i = 0; i < n; i += stride) {
      var cx = margin.left + bandW * i + bandW / 2;
      var t = el("text", { x: cx, y: H - 8, "text-anchor": "middle", "font-size": 10.5, fill: "var(--text-muted)" });
      t.textContent = opts.xFormat ? opts.xFormat(cats[i], i) : cats[i];
      gX.appendChild(t);
    }
    svg.appendChild(gX);

    var barPad = bandW < 40 ? 0.22 : 0.3;
    var hitAreas = [];
    var wrap = document.createElement("div");
    wrap.className = "chart-body";
    wrap.style.position = "relative";

    if (mode === "stacked") {
      var gap = 1.5;
      for (i = 0; i < n; i++) {
        var x0 = margin.left + bandW * i + bandW * barPad / 2;
        var bw = bandW * (1 - barPad);
        var stackTop = 0, stackBot = 0;
        var segs = [];
        series.forEach(function (s) {
          var v = s.values[i] || 0;
          if (v >= 0) {
            var yTop = yScale(stackTop + v), yBotY = yScale(stackTop);
            var h = Math.max(0, yBotY - yTop - gap);
            if (h > 0) svg.appendChild(el("rect", { x: x0, y: yTop, width: bw, height: h, rx: 2, fill: s.color }));
            segs.push({ name: s.name, color: s.color, value: v });
            stackTop += v;
          } else {
            var yTopN = yScale(stackBot), yBotN = yScale(stackBot + v);
            var hN = Math.max(0, yBotN - yTopN - gap);
            if (hN > 0) svg.appendChild(el("rect", { x: x0, y: yTopN, width: bw, height: hN, rx: 2, fill: s.color }));
            segs.push({ name: s.name, color: s.color, value: v });
            stackBot += v;
          }
        });
        hitAreas.push({ x: margin.left + bandW * i, w: bandW, cat: cats[i], segs: segs, total: stackTop + stackBot });
      }
    } else {
      var visibleSeries = series;
      var groupW = bandW * (1 - barPad);
      var subW = groupW / visibleSeries.length;
      for (i = 0; i < n; i++) {
        var gx0 = margin.left + bandW * i + bandW * barPad / 2;
        var segs2 = [];
        visibleSeries.forEach(function (s, si) {
          var v = s.values[i] || 0;
          var bx = gx0 + si * subW;
          var yv = yScale(Math.max(0, v)), yb = yScale(Math.min(0, v));
          var h = Math.max(0, yScale(0) - Math.min(yv, yScale(0)));
          var top = v >= 0 ? yv : y0;
          var hh = Math.abs(y0 - (v >= 0 ? yv : yv));
          var rectY = v >= 0 ? yv : y0;
          var rectH = Math.abs(y0 - yv);
          svg.appendChild(el("rect", { x: bx + 1, y: rectY, width: Math.max(0, subW - 2), height: rectH, rx: 2, fill: s.color }));
          segs2.push({ name: s.name, color: s.color, value: v });
        });
        hitAreas.push({ x: margin.left + bandW * i, w: bandW, cat: cats[i], segs: segs2 });
      }
    }

    wrap.appendChild(svg);
    var tip = ensureTooltip(wrap);
    container.appendChild(wrap);

    var hitRect = el("rect", { x: margin.left, y: margin.top, width: Math.max(innerW,1), height: Math.max(innerH,1), fill: "transparent" });
    svg.appendChild(hitRect);
    hitRect.addEventListener("mousemove", function (evt) {
      var rect = svg.getBoundingClientRect();
      var scaleX = W / rect.width, scaleY = H / rect.height;
      var mx = (evt.clientX - rect.left) * scaleX;
      var idx = Math.floor((mx - margin.left) / bandW);
      idx = Math.max(0, Math.min(n - 1, idx));
      var ha = hitAreas[idx];
      if (!ha) return;
      var rows = ha.segs.map(function (s) {
        return '<div class="t-row"><span class="t-swatch" style="background:' + s.color + '"></span>' + s.name + ": <b>" + (opts.yFormat ? opts.yFormat(s.value) : fmtNum(s.value)) + "</b></div>";
      }).join("");
      var px = (ha.x + ha.w / 2) / scaleX;
      var py = margin.top / scaleY;
      showTooltip(wrap, tip, px, py + 6, '<div class="t-title">' + (opts.xFormat ? opts.xFormat(ha.cat, idx) : ha.cat) + "</div>" + rows);
    });
    hitRect.addEventListener("mouseleave", function () { hideTooltip(tip); });

    return { svg: svg };
  }

  /** Horizontal Pareto/ranked bar chart for cause analysis. opts:{items:[{label,value,color}], width,height} */
  function paretoChart(container, opts) {
    container.innerHTML = "";
    var items = opts.items.slice().sort(function (a, b) { return b.value - a.value; }).slice(0, opts.limit || 10);
    var W = opts.width || container.clientWidth || 640;
    var rowH = opts.rowH || 26;
    var margin = { top: 6, right: 50, bottom: 10, left: opts.labelWidth || 150 };
    var H = margin.top + margin.bottom + items.length * rowH;
    var innerW = W - margin.left - margin.right;
    var maxV = Math.max.apply(null, items.map(function (i) { return i.value; }).concat([1]));

    var svg = el("svg", { viewBox: "0 0 " + W + " " + H, width: "100%", height: H });
    items.forEach(function (it, i) {
      var y = margin.top + i * rowH;
      var bw = maxV > 0 ? (it.value / maxV) * innerW : 0;
      var lbl = el("text", { x: margin.left - 8, y: y + rowH / 2 + 4, "text-anchor": "end", "font-size": 11.5, fill: "var(--text-secondary)" });
      lbl.textContent = it.label;
      svg.appendChild(lbl);
      svg.appendChild(el("rect", { x: margin.left, y: y + 4, width: Math.max(bw,2), height: rowH - 10, rx: 3, fill: it.color || "var(--series-6)" }));
      var vlbl = el("text", { x: margin.left + bw + 8, y: y + rowH / 2 + 4, "font-size": 11, "font-weight": 700, fill: "var(--text-primary)" });
      vlbl.textContent = opts.valueFormat ? opts.valueFormat(it.value) : fmtNum(it.value);
      svg.appendChild(vlbl);
    });
    container.appendChild(svg);
    return { svg: svg };
  }

  function sparkline(container, values, opts) {
    opts = opts || {};
    var W = opts.width || 100, H = opts.height || 28;
    container.innerHTML = "";
    var svg = el("svg", { viewBox: "0 0 " + W + " " + H, width: W, height: H });
    var vals = values.filter(function (v) { return v !== null && v !== undefined; });
    if (!vals.length) { container.appendChild(svg); return; }
    var min = Math.min.apply(null, vals), max = Math.max.apply(null, vals);
    if (min === max) { min -= 1; max += 1; }
    var n = values.length;
    var d = "";
    values.forEach(function (v, i) {
      if (v === null || v === undefined) return;
      var x = n <= 1 ? W / 2 : (i / (n - 1)) * (W - 4) + 2;
      var y = H - 3 - ((v - min) / (max - min)) * (H - 6);
      d += (d ? " L " : "M ") + x.toFixed(1) + " " + y.toFixed(1);
    });
    svg.appendChild(el("path", { d: d, fill: "none", stroke: opts.color || "var(--series-1)", "stroke-width": 1.75, "stroke-linecap": "round", "stroke-linejoin": "round" }));
    container.appendChild(svg);
  }

  var Charts = { lineChart: lineChart, barChart: barChart, paretoChart: paretoChart, sparkline: sparkline, fmtNum: fmtNum };
  if (typeof module !== "undefined" && module.exports) module.exports = Charts;
  else root.Charts = Charts;
})(typeof window !== "undefined" ? window : this);
