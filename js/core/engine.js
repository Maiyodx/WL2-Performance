/**
 * WL2 Performance — computation engine.
 * Replicates the Daily_Input -> Calc_Daily -> Weekly/Monthly/Yearly_Summary
 * formulas from WL2_Performance_v6.xlsx. Pure functions, no DOM dependency
 * (usable in Node for testing and in the browser for the dashboard).
 *
 * All tunable settings (column layout, labels, default assumptions) live in
 * js/config/config.js — this file only contains calculation logic.
 */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory(require("../config/config.js"));
  } else {
    root.WL2 = factory(root.WL2_CONFIG);
  }
})(typeof window !== "undefined" ? window : this, function (CONFIG) {
  "use strict";

  var ASSUMPTIONS_DEFAULT = CONFIG.ASSUMPTIONS_DEFAULT;
  var COL = CONFIG.DAILY_INPUT_COLUMNS;
  var DATA_START_ROW = CONFIG.DATA_START_ROW;
  var DATA_END_ROW = CONFIG.DATA_END_ROW;

  function num(v) {
    if (v === null || v === undefined || v === "") return 0;
    var n = Number(v);
    return isNaN(n) ? 0 : n;
  }

  // Excel WEEKNUM(date, 2): week 1 = week containing Jan 1, weeks start Monday
  function weekNum2(date) {
    var year = date.getFullYear();
    var jan1 = new Date(year, 0, 1);
    var jan1Mon = (jan1.getDay() + 6) % 7; // Mon=0..Sun=6
    var dayOfYear0 = Math.round((date - jan1) / 86400000);
    return Math.floor((dayOfYear0 + jan1Mon) / 7) + 1;
  }

  function excelSerialToDate(v) {
    if (v instanceof Date) return v;
    if (typeof v === "number") {
      // Excel serial date (1900 system), matches SheetJS's own epoch calc
      var utc_days = Math.floor(v - 25569);
      var utc_value = utc_days * 86400;
      return new Date(utc_value * 1000);
    }
    return null;
  }

  /**
   * Parse the Daily_Input sheet (array-of-arrays, header:1 layout) into
   * a flat list of raw day rows.
   */
  function parseDailyInput(rows) {
    var out = [];
    for (var r = DATA_START_ROW; r <= DATA_END_ROW; r++) {
      var row = rows[r];
      if (!row) continue;
      var dateRaw = row[COL.date];
      if (dateRaw === undefined || dateRaw === null || dateRaw === "") continue;
      var date = excelSerialToDate(dateRaw);
      if (!date) continue;
      var g = function (idx) { return row[idx]; };
      var d = {
        date: date,
        month: date.getMonth() + 1,
        week: weekNum2(date),
        season: row[COL.season] || "",
        holThai: num(g(COL.holThai)), holLao: num(g(COL.holLao)), holOther: num(g(COL.holOther)),
        pmSystem: num(g(COL.pmSystem)), pmOther: num(g(COL.pmOther)),
        udOperation: num(g(COL.udOperation)), udMechanical: num(g(COL.udMechanical)),
        udElectrical: num(g(COL.udElectrical)), udPLC: num(g(COL.udPLC)), udBelt: num(g(COL.udBelt)),
        faceM8: num(g(COL.faceM8)), faceM8A: num(g(COL.faceM8A)), faceM9: num(g(COL.faceM9)), faceM10: num(g(COL.faceM10)),
        crABD: num(g(COL.crABD)), crBBD: num(g(COL.crBBD)), crCBD: num(g(COL.crCBD)),
        crDBD: num(g(COL.crDBD)), crEBD: num(g(COL.crEBD)), crFBD: num(g(COL.crFBD)),
        outPower: num(g(COL.outPower)), outMDS: num(g(COL.outMDS)), outOther: num(g(COL.outOther)),
        relocMainDump: num(g(COL.relocMainDump)), relocInPit: num(g(COL.relocInPit)),
        relocConvShift: num(g(COL.relocConvShift)), relocOther: num(g(COL.relocOther)),
        idleEmpty: num(g(COL.idleEmpty)), idleRain: num(g(COL.idleRain)), idleWaitTruck: num(g(COL.idleWaitTruck)),
        idleSeqStart: num(g(COL.idleSeqStart)), idleSPD: num(g(COL.idleSPD)), idleSubsidence: num(g(COL.idleSubsidence)), idleOther: num(g(COL.idleOther)),
        delayRain: num(g(COL.delayRain)), delayNoFeed: num(g(COL.delayNoFeed)), delayTrunk: num(g(COL.delayTrunk)), delayOther: num(g(COL.delayOther)),
        prodConD: num(g(COL.prodConD)), prodConF: num(g(COL.prodConF)), prodOther: num(g(COL.prodOther)),
        hd785Trips: num(g(COL.hd785Trips)), hd785Rate: num(g(COL.hd785Rate)),
        tonlyTrips: num(g(COL.tonlyTrips)), tonlyRate: num(g(COL.tonlyRate)),
        dt12dTrips: num(g(COL.dt12dTrips)), dt12dRate: num(g(COL.dt12dRate)),
        dt10Trips: num(g(COL.dt10Trips)), dt10Rate: num(g(COL.dt10Rate)),
        dt12fTrips: num(g(COL.dt12fTrips)), dt12fRate: num(g(COL.dt12fRate)),
        tianmaTrips: num(g(COL.tianmaTrips)), tianmaRate: num(g(COL.tianmaRate)),
        crAProd: num(g(COL.crAProd)), crAHrs: num(g(COL.crAHrs)),
        crBProd: num(g(COL.crBProd)), crBHrs: num(g(COL.crBHrs)),
        crCProd: num(g(COL.crCProd)), crCHrs: num(g(COL.crCHrs)),
        crDProd: num(g(COL.crDProd)), crDHrs: num(g(COL.crDHrs)),
        crEProd: num(g(COL.crEProd)), crEHrs: num(g(COL.crEHrs)),
        crFProd: num(g(COL.crFProd)), crFHrs: num(g(COL.crFHrs)),
        energy: num(g(COL.energy)),
        breakdownCount: num(g(COL.breakdownCount))
      };
      // replicate the "has any data" gate used throughout the workbook
      var sumErc = 0;
      for (var key in COL) {
        if (key === "date" || key === "month" || key === "week" || key === "season") continue;
        sumErc += num(g(COL[key]));
      }
      d.hasData = (d.season !== "" && d.season !== null) || sumErc !== 0;
      out.push(d);
    }
    return out;
  }

  /** Parse Plan_Monthly sheet rows (array-of-arrays) into a map keyed by month 1-12 */
  function parsePlanMonthly(rows) {
    var plan = {};
    for (var r = 4; r <= 15; r++) {
      var row = rows[r];
      if (!row) continue;
      var m = num(row[0]);
      if (!m) continue;
      plan[m] = {
        month: m,
        name: row[1],
        days: num(row[2]),
        wasteProdPlan: num(row[3]),
        conDPlan: num(row[4]),
        conFPlan: num(row[5]),
        qplanRate: num(row[6]),
        crRates: [num(row[7]), num(row[8]), num(row[9]), num(row[10]), num(row[11]), num(row[12])],
        energyPlan: num(row[13])
      };
    }
    return plan;
  }

  function parsePlanWeekly(rows) {
    var plan = {};
    for (var r = 4; r <= 56; r++) {
      var row = rows[r];
      if (!row) continue;
      var w = num(row[0]);
      if (!w) continue;
      plan[w] = num(row[1]);
    }
    return plan;
  }

  function parseAssumptions(rows) {
    var a = Object.assign({}, ASSUMPTIONS_DEFAULT);
    for (var r = 3; r <= 13; r++) {
      var row = rows[r];
      if (!row) continue;
      var symbol = row[2], val = row[3];
      if (symbol === "Qvde") a.Qvde = num(val);
      else if (symbol === "Qvef") a.Qvef = num(val);
      else if (symbol === "Qcap") a.Qcap = num(val);
      else if (symbol === "MTBF_req") a.MTBF_req = num(val);
      else if (symbol === "OEE_tgt") a.OEE_tgt = num(val);
      else if (symbol === "HRS_DAY") a.HRS_DAY = num(val);
    }
    return a;
  }

  /** Compute Calc_Daily-equivalent KPI fields for one parsed day row. */
  function computeDayKPI(d, assumptions, planMonthly) {
    if (!d.hasData) return null;
    var HRS_DAY = assumptions.HRS_DAY;
    var CT = HRS_DAY;
    var HT = (d.holThai + d.holLao + d.holOther) * HRS_DAY;
    var ST = CT - HT;
    var PD = d.pmSystem + d.pmOther;
    var UD = d.udOperation + d.udMechanical + d.udElectrical + d.udPLC + d.udBelt;
    var MT = PD + UD;
    var UT = d.outPower + d.outMDS + d.outOther;
    var PR = d.relocMainDump + d.relocInPit + d.relocConvShift + d.relocOther;
    var PA = UT + PR;
    var AT = ST - MT - PA;
    var IT = d.idleEmpty + d.idleRain + d.idleWaitTruck + d.idleSeqStart + d.idleSPD + d.idleSubsidence + d.idleOther;
    var OT = AT - IT;
    var DT = d.delayRain + d.delayNoFeed + d.delayTrunk + d.delayOther;
    var PT = OT - DT;
    var S1 = d.prodConD, S2 = d.prodConF, S3 = d.prodOther;
    var S = S1 + S2 + S3;
    var Qvs = PT !== 0 ? S / PT : null;
    // Breakdown Frequency: use the real event count (d.breakdownCount) when the
    // day has it filled in; otherwise fall back to the legacy proxy
    // (Operating Time / MTBF_req) so rows from before this column existed
    // still produce a number instead of a blank.
    var usingTrueBreakdownCount = d.breakdownCount > 0;
    var BF = usingTrueBreakdownCount ? d.breakdownCount : (assumptions.MTBF_req !== 0 ? OT / assumptions.MTBF_req : 0);
    var MTTR = (UD !== 0 && BF !== 0) ? UD / BF : null;
    var MTBF = (OT !== 0 && BF !== 0) ? OT / BF : null;
    var UO = OT !== 0 ? (PT / OT) * 100 : null;
    var OA = ST !== 0 ? (AT / ST) * 100 : null;
    var UA = AT !== 0 ? (OT / AT) * 100 : null;
    var MA = (OT + MT) !== 0 ? (OT / (OT + MT)) * 100 : null;
    var RE = (AT + UD) !== 0 ? (AT / (AT + UD)) * 100 : null;
    var plan = planMonthly[d.month];
    var planRate = plan ? plan.qplanRate : 0;
    // Performance (Actual Output / Ideal Output) — mathematically identical to
    // CU below (both are actual-rate / plan-rate), kept as one field.
    var CU = (Qvs !== null && planRate) ? (Qvs / planRate) * 100 : null;
    // OEE = Availability x Performance x Quality. This operation has no
    // reject/defect concept (Quality = 100% always), so OEE = OA x CU / 100.
    var OEE = (OA !== null && CU !== null) ? (OA * CU) / 100 : null;

    // detail/informational breakdown (not part of the main ST waterfall)
    var faceConvBD = d.faceM8 + d.faceM8A + d.faceM9 + d.faceM10;
    var crusherBD = d.crABD + d.crBBD + d.crCBD + d.crDBD + d.crEBD + d.crFBD;

    var fleetProd = {
      HD785: d.hd785Trips * d.hd785Rate,
      Tonly: d.tonlyTrips * d.tonlyRate,
      "12DT-D": d.dt12dTrips * d.dt12dRate,
      "10DT": d.dt10Trips * d.dt10Rate,
      "12DT-F": d.dt12fTrips * d.dt12fRate,
      Tianma: d.tianmaTrips * d.tianmaRate
    };
    var crusherStats = {
      A: { prod: d.crAProd, hrs: d.crAHrs },
      B: { prod: d.crBProd, hrs: d.crBHrs },
      C: { prod: d.crCProd, hrs: d.crCHrs },
      D: { prod: d.crDProd, hrs: d.crDHrs },
      E: { prod: d.crEProd, hrs: d.crEHrs },
      F: { prod: d.crFProd, hrs: d.crFHrs }
    };

    return {
      date: d.date, month: d.month, week: d.week, season: d.season,
      CT: CT, HT: HT, ST: ST, PD: PD, UD: UD, MT: MT, UT: UT, PR: PR, PA: PA,
      AT: AT, IT: IT, OT: OT, DT: DT, PT: PT,
      S: S, S1: S1, S2: S2, S3: S3, Qvs: Qvs, BF: BF, MTTR: MTTR, MTBF: MTBF,
      UO: UO, OA: OA, UA: UA, MA: MA, RE: RE, CU: CU, OEE: OEE, energy: d.energy,
      usingTrueBreakdownCount: usingTrueBreakdownCount,
      faceConvBD: faceConvBD, crusherBD: crusherBD,
      fleetProd: fleetProd, crusherStats: crusherStats,
      raw: d
    };
  }

  function computeAllDays(rawDays, assumptions, planMonthly) {
    var kpis = [];
    for (var i = 0; i < rawDays.length; i++) {
      var k = computeDayKPI(rawDays[i], assumptions, planMonthly);
      if (k) kpis.push(k);
    }
    return kpis;
  }

  function avg(arr) {
    var vals = arr.filter(function (v) { return v !== null && v !== undefined && !isNaN(v); });
    if (!vals.length) return null;
    return vals.reduce(function (a, b) { return a + b; }, 0) / vals.length;
  }

  function groupBy(kpis, keyFn) {
    var map = {};
    kpis.forEach(function (k) {
      var key = keyFn(k);
      if (!map[key]) map[key] = [];
      map[key].push(k);
    });
    return map;
  }

  function summarizeWeek(weekKpis, weekNo, planWeekly) {
    var PT = weekKpis.reduce(function (a, k) { return a + k.PT; }, 0);
    var OT = weekKpis.reduce(function (a, k) { return a + k.OT; }, 0);
    var Production = weekKpis.reduce(function (a, k) { return a + k.S; }, 0);
    var Rate = PT !== 0 ? Production / PT : null;
    var UO = avg(weekKpis.map(function (k) { return k.UO; }));
    var planRate = planWeekly[weekNo] || 0;
    var CU = (UO !== null && planRate) ? avg(weekKpis.map(function (k) { return k.Qvs; })) / planRate * 100 : null;
    var OA = avg(weekKpis.map(function (k) { return k.OA; }));
    // OEE = Availability x Performance x Quality(=100%) -> OA x CU / 100
    var OEE = (OA !== null && CU !== null) ? OA * CU / 100 : null;
    var MA = avg(weekKpis.map(function (k) { return k.MA; }));
    var UA = avg(weekKpis.map(function (k) { return k.UA; }));
    return { week: weekNo, PT: PT, OT: OT, production: Production, rate: Rate, UO: UO, UA: UA, CU: CU, OEE: OEE, OA: OA, MA: MA, planRate: planRate, days: weekKpis.length };
  }

  function summarizeMonth(monthKpis, monthNo, planMonthly, assumptions) {
    var plan = planMonthly[monthNo] || { days: 0, wasteProdPlan: 0, qplanRate: 0, name: "" };
    var sPT = monthKpis.reduce(function (a, k) { return a + k.PT; }, 0);
    var sOT = monthKpis.reduce(function (a, k) { return a + k.OT; }, 0);
    var sST = monthKpis.reduce(function (a, k) { return a + k.ST; }, 0);
    var sAT = monthKpis.reduce(function (a, k) { return a + k.AT; }, 0);
    var sMT = monthKpis.reduce(function (a, k) { return a + k.MT; }, 0);
    var sUD = monthKpis.reduce(function (a, k) { return a + k.UD; }, 0);
    var sIT = monthKpis.reduce(function (a, k) { return a + k.IT; }, 0);
    var sDT = monthKpis.reduce(function (a, k) { return a + k.DT; }, 0);
    var Production = monthKpis.reduce(function (a, k) { return a + k.S; }, 0);
    var S1 = monthKpis.reduce(function (a, k) { return a + k.S1; }, 0);
    var S2 = monthKpis.reduce(function (a, k) { return a + k.S2; }, 0);
    var S3 = monthKpis.reduce(function (a, k) { return a + k.S3; }, 0);
    var Energy = monthKpis.reduce(function (a, k) { return a + k.energy; }, 0);
    var Rate = sPT !== 0 ? Production / sPT : null;

    var UOw = sOT !== 0 ? sPT / sOT * 100 : null;
    var OAw = sST !== 0 ? sAT / sST * 100 : null;
    var UAw = sAT !== 0 ? sOT / sAT * 100 : null;
    var MAw = (sOT + sMT) !== 0 ? sOT / (sOT + sMT) * 100 : null;
    var REw = (sAT + sUD) !== 0 ? sAT / (sAT + sUD) * 100 : null;
    var CUw = (Rate !== null && plan.qplanRate) ? Rate / plan.qplanRate * 100 : null;
    // OEE = Availability x Performance x Quality(=100%) -> OA x CU / 100
    var OEEw = (OAw !== null && CUw !== null) ? OAw * CUw / 100 : null;

    var UOavg = avg(monthKpis.map(function (k) { return k.UO; }));
    var UAavg = avg(monthKpis.map(function (k) { return k.UA; }));
    var CUavg = (plan.qplanRate) ? avg(monthKpis.map(function (k) { return k.Qvs; })) / plan.qplanRate * 100 : null;
    var OAavg = avg(monthKpis.map(function (k) { return k.OA; }));
    var OEEavg = (OAavg !== null && CUavg !== null) ? OAavg * CUavg / 100 : null;
    var MAavg = avg(monthKpis.map(function (k) { return k.MA; }));

    // True MTBF/MTTR aggregated over the period: sum the actual breakdown
    // counts (falling back to the legacy per-day proxy for rows that don't
    // have the count filled in) rather than averaging daily ratios, which
    // avoids single zero-breakdown days distorting the period figure.
    var sumBreakdownCount = monthKpis.reduce(function (a, k) { return a + k.BF; }, 0);
    var MTBFw = (sOT !== 0 && sumBreakdownCount !== 0) ? sOT / sumBreakdownCount : null;
    var MTTRw = (sUD !== 0 && sumBreakdownCount !== 0) ? sUD / sumBreakdownCount : null;

    return {
      month: monthNo, name: plan.name, days: plan.days,
      sumPT: sPT, sumOT: sOT, sumST: sST, sumAT: sAT, sumMT: sMT, sumUD: sUD, sumIT: sIT, sumDT: sDT,
      production: Production, S1: S1, S2: S2, S3: S3, energy: Energy, rate: Rate,
      MTBF: MTBFw, MTTR: MTTRw, sumBreakdownCount: sumBreakdownCount,
      w: { UO: UOw, UA: UAw, OA: OAw, MA: MAw, RE: REw, CU: CUw, OEE: OEEw },
      avgMethod: { UO: UOavg, UA: UAavg, CU: CUavg, OEE: OEEavg, OA: OAavg, MA: MAavg },
      planProd: plan.wasteProdPlan, planRate: plan.qplanRate,
      varProd: Production ? Production - plan.wasteProdPlan : null,
      varOEE: OEEw !== null ? OEEw - assumptions.OEE_tgt : null
    };
  }

  function computeMonthlySummary(kpis, planMonthly, assumptions) {
    var byMonth = groupBy(kpis, function (k) { return k.month; });
    var months = [];
    for (var m = 1; m <= 12; m++) {
      months.push(summarizeMonth(byMonth[m] || [], m, planMonthly, assumptions));
    }
    return months;
  }

  function computeWeeklySummary(kpis, planWeekly) {
    var byWeek = groupBy(kpis, function (k) { return k.week; });
    var weeks = Object.keys(byWeek).map(Number).sort(function (a, b) { return a - b; });
    return weeks.map(function (w) { return summarizeWeek(byWeek[w], w, planWeekly); });
  }

  function parseMasterMachine(rows) {
    var machines = {};
    for (var r = 4; r <= 40; r++) {
      var row = rows[r];
      if (!row || !row[0]) continue;
      machines[row[0]] = { id: row[0], name: row[1], type: row[2], capacity: num(row[3]), unit: row[4], contractor: row[5], status: row[6] };
    }
    return machines;
  }

  function parseMasterContractor(rows) {
    var out = {};
    for (var r = 4; r <= 20; r++) {
      var row = rows[r];
      if (!row || !row[0]) continue;
      out[row[0]] = { id: row[0], name: row[1], scope: row[2], status: row[3] };
    }
    return out;
  }

  /** Replicates Data_Check sheet logic per raw day row. */
  function computeDataQuality(rawDays) {
    var seen = {};
    rawDays.forEach(function (d) {
      var key = d.date.getFullYear() + "-" + d.date.getMonth() + "-" + d.date.getDate();
      seen[key] = (seen[key] || 0) + 1;
    });
    return rawDays.filter(function (d) { return d.hasData; }).map(function (d) {
      var key = d.date.getFullYear() + "-" + d.date.getMonth() + "-" + d.date.getDate();
      var missingSeason = !d.season;
      var allFields = [];
      for (var k in COL) {
        if (k === "date" || k === "month" || k === "week" || k === "season") continue;
        allFields.push(d[k]);
      }
      var negative = allFields.some(function (v) { return v < -0.001; });
      var HT = (d.holThai + d.holLao + d.holOther) * 24;
      var MT = d.pmSystem + d.pmOther + d.udOperation + d.udMechanical + d.udElectrical + d.udPLC + d.udBelt;
      var PAITDT = d.outPower + d.outMDS + d.outOther + d.relocMainDump + d.relocInPit + d.relocConvShift + d.relocOther +
        d.idleEmpty + d.idleRain + d.idleWaitTruck + d.idleSeqStart + d.idleSPD + d.idleSubsidence + d.idleOther +
        d.delayRain + d.delayNoFeed + d.delayTrunk + d.delayOther;
      var over24h = (HT + MT + PAITDT) > 24;
      var noProduction = (d.prodConD + d.prodConF + d.prodOther) === 0;
      var duplicate = seen[key] > 1;
      var unplannedBD = d.udOperation + d.udMechanical + d.udElectrical + d.udPLC + d.udBelt;
      // soft warning only (does not affect PASS/CHECK) — there was unplanned
      // downtime that day but no breakdown-event count was entered, so
      // MTBF/MTTR for that day fall back to the legacy proxy instead of a
      // real figure.
      var missingBreakdownCount = unplannedBD > 0 && !(d.breakdownCount > 0);
      var pass = !missingSeason && !negative && !over24h && !duplicate;
      return {
        date: d.date, missingSeason: missingSeason, negative: negative, over24h: over24h,
        noProduction: noProduction, duplicate: duplicate, missingBreakdownCount: missingBreakdownCount,
        status: pass ? "PASS" : "CHECK"
      };
    });
  }

  /** Generic period aggregator (like Period_Summary) over an arbitrary kpi subset. */
  function computePeriodSummary(kpisSubset, planMonthly, assumptions) {
    var n = kpisSubset.length;
    var sPT = 0, sOT = 0, sST = 0, sAT = 0, sMT = 0, sUD = 0, sIT = 0, sDT = 0, production = 0, energy = 0;
    var monthDayCount = {};
    kpisSubset.forEach(function (k) {
      sPT += k.PT; sOT += k.OT; sST += k.ST; sAT += k.AT; sMT += k.MT; sUD += k.UD; sIT += k.IT; sDT += k.DT;
      production += k.S; energy += k.energy;
      monthDayCount[k.month] = (monthDayCount[k.month] || 0) + 1;
    });
    var rate = sPT !== 0 ? production / sPT : null;
    var UOw = sOT !== 0 ? sPT / sOT * 100 : null;
    var OAw = sST !== 0 ? sAT / sST * 100 : null;
    var UAw = sAT !== 0 ? sOT / sAT * 100 : null;
    var MAw = (sOT + sMT) !== 0 ? sOT / (sOT + sMT) * 100 : null;
    var REw = (sAT + sUD) !== 0 ? sAT / (sAT + sUD) * 100 : null;

    var totalDays = 0, weightedRateSum = 0;
    for (var m in monthDayCount) {
      var plan = planMonthly[m];
      var cnt = monthDayCount[m];
      totalDays += cnt;
      weightedRateSum += (plan ? plan.qplanRate : 0) * cnt;
    }
    var blendedPlanRate = totalDays ? weightedRateSum / totalDays : 0;
    var CUw = (rate !== null && blendedPlanRate) ? rate / blendedPlanRate * 100 : null;
    // OEE = Availability x Performance x Quality(=100%) -> OA x CU / 100
    var OEEw = (OAw !== null && CUw !== null) ? OAw * CUw / 100 : null;

    var UOavg = avg(kpisSubset.map(function (k) { return k.UO; }));
    var UAavg = avg(kpisSubset.map(function (k) { return k.UA; }));
    var CUavg = blendedPlanRate ? avg(kpisSubset.map(function (k) { return k.Qvs; })) / blendedPlanRate * 100 : null;
    var OAavg = avg(kpisSubset.map(function (k) { return k.OA; }));
    var OEEavg = (OAavg !== null && CUavg !== null) ? OAavg * CUavg / 100 : null;
    var MAavg = avg(kpisSubset.map(function (k) { return k.MA; }));
    var REavg = avg(kpisSubset.map(function (k) { return k.RE; }));

    var sumBreakdownCount = kpisSubset.reduce(function (a, k) { return a + k.BF; }, 0);
    var MTBFw = (sOT !== 0 && sumBreakdownCount !== 0) ? sOT / sumBreakdownCount : null;
    var MTTRw = (sUD !== 0 && sumBreakdownCount !== 0) ? sUD / sumBreakdownCount : null;

    return {
      days: n, sumPT: sPT, sumOT: sOT, sumST: sST, sumAT: sAT, sumMT: sMT, sumUD: sUD, sumIT: sIT, sumDT: sDT,
      production: production, energy: energy, rate: rate, planRate: blendedPlanRate,
      MTBF: MTBFw, MTTR: MTTRw,
      w: { UO: UOw, UA: UAw, OA: OAw, MA: MAw, RE: REw, CU: CUw, OEE: OEEw },
      avgMethod: { UO: UOavg, UA: UAavg, CU: CUavg, OEE: OEEavg, OA: OAavg, MA: MAavg, RE: REavg },
      varOEE: OEEw !== null ? OEEw - assumptions.OEE_tgt : null
    };
  }

  function computeYTD(monthlySummaries, assumptions) {
    var sPT = 0, sOT = 0, sST = 0, sAT = 0, sMT = 0, sUD = 0, production = 0, planProdSum = 0, daysSum = 0, planRates = [];
    var sumBreakdownCount = 0;
    monthlySummaries.forEach(function (ms) {
      sPT += ms.sumPT; sOT += ms.sumOT; sST += ms.sumST; sAT += ms.sumAT; sMT += ms.sumMT; sUD += ms.sumUD;
      production += ms.production; planProdSum += ms.planProd; daysSum += ms.days;
      planRates.push(ms.planRate);
      sumBreakdownCount += ms.sumBreakdownCount || 0;
    });
    var rate = sPT !== 0 ? production / sPT : 0;
    var UO = sOT !== 0 ? sPT / sOT * 100 : 0;
    var OA = sST !== 0 ? sAT / sST * 100 : 0;
    var UA = sAT !== 0 ? sOT / sAT * 100 : 0;
    var MA = (sOT + sMT) !== 0 ? sOT / (sOT + sMT) * 100 : 0;
    var RE = (sAT + sUD) !== 0 ? sAT / (sAT + sUD) * 100 : 0;
    var avgPlanRate = avg(planRates) || 0;
    var CU = avgPlanRate ? rate / avgPlanRate * 100 : 0;
    // OEE = Availability x Performance x Quality(=100%) -> OA x CU / 100
    var OEE = OA * CU / 100;
    var achievementPct = planProdSum ? production / planProdSum * 100 : 0;
    var MTBF = sumBreakdownCount !== 0 ? sOT / sumBreakdownCount : null;
    var MTTR = sumBreakdownCount !== 0 ? sUD / sumBreakdownCount : null;
    return {
      production: production, planProd: planProdSum, achievementPct: achievementPct,
      rate: rate, sumPT: sPT, sumOT: sOT, sumST: sST, sumAT: sAT, sumMT: sMT, sumUD: sUD,
      UO: UO, UA: UA, OA: OA, MA: MA, RE: RE, CU: CU, OEE: OEE, MTBF: MTBF, MTTR: MTTR,
      varOEE: OEE - assumptions.OEE_tgt,
      days: daysSum
    };
  }

  return {
    ASSUMPTIONS_DEFAULT: ASSUMPTIONS_DEFAULT,
    COL: COL,
    weekNum2: weekNum2,
    excelSerialToDate: excelSerialToDate,
    parseDailyInput: parseDailyInput,
    parsePlanMonthly: parsePlanMonthly,
    parsePlanWeekly: parsePlanWeekly,
    parseAssumptions: parseAssumptions,
    parseMasterMachine: parseMasterMachine,
    parseMasterContractor: parseMasterContractor,
    computeDataQuality: computeDataQuality,
    computePeriodSummary: computePeriodSummary,
    computeDayKPI: computeDayKPI,
    computeAllDays: computeAllDays,
    computeMonthlySummary: computeMonthlySummary,
    computeWeeklySummary: computeWeeklySummary,
    computeYTD: computeYTD,
    groupBy: groupBy,
    avg: avg
  };
});
