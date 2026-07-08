/**
 * WL2_CONFIG — all the settings someone maintaining this dashboard is likely
 * to need to change. Nothing here contains calculation logic or DOM code —
 * edit this file freely; engine.js / app.js only ever read from it.
 */
(function (root) {
  "use strict";

  var WL2_CONFIG = {

    // ---- localStorage keys (bump the "_v1" suffix if you change the data
    // shape below, so old cached files don't get parsed with new code) ----
    STORAGE_KEYS: {
      FILE: "wl2_dashboard_file_v1",
      META: "wl2_dashboard_meta_v1",
      THEME: "wl2_dashboard_theme_v1"
    },

    // ---- fallback values, used only if the Assumptions sheet is missing a
    // row. Normal operation always reads the real values from the workbook ----
    ASSUMPTIONS_DEFAULT: {
      Qvde: 10398.01,
      Qvef: 8955.22,
      Qcap: 7000,
      MTBF_req: 5,
      OEE_tgt: 85,
      HRS_DAY: 24
    },

    // ---- Daily_Input column layout (0-based index into each row array).
    // Only change this if the Excel template's column order changes. ----
    DAILY_INPUT_COLUMNS: {
      date: 0, month: 1, week: 2, season: 3,
      holThai: 4, holLao: 5, holOther: 6,
      pmSystem: 7, pmOther: 8,
      udOperation: 9, udMechanical: 10, udElectrical: 11, udPLC: 12, udBelt: 13,
      faceM8: 14, faceM8A: 15, faceM9: 16, faceM10: 17,
      crABD: 18, crBBD: 19, crCBD: 20, crDBD: 21, crEBD: 22, crFBD: 23,
      outPower: 24, outMDS: 25, outOther: 26,
      relocMainDump: 27, relocInPit: 28, relocConvShift: 29, relocOther: 30,
      idleEmpty: 31, idleRain: 32, idleWaitTruck: 33, idleSeqStart: 34, idleSPD: 35, idleSubsidence: 36, idleOther: 37,
      delayRain: 38, delayNoFeed: 39, delayTrunk: 40, delayOther: 41,
      prodConD: 42, prodConF: 43, prodOther: 44,
      hd785Trips: 45, hd785Rate: 46,
      tonlyTrips: 47, tonlyRate: 48,
      dt12dTrips: 49, dt12dRate: 50,
      dt10Trips: 51, dt10Rate: 52,
      dt12fTrips: 53, dt12fRate: 54,
      tianmaTrips: 55, tianmaRate: 56,
      crAProd: 57, crAHrs: 58,
      crBProd: 59, crBHrs: 60,
      crCProd: 61, crCHrs: 62,
      crDProd: 63, crDHrs: 64,
      crEProd: 65, crEHrs: 66,
      crFProd: 67, crFHrs: 68,
      energy: 69,
      // Added column (BS) beyond the original template — see README section
      // "MTBF/MTTR" below. Not present in older copies of the workbook;
      // parseDailyInput() treats a missing/blank value as 0 and engine.js
      // falls back to the legacy MTBF proxy for those rows.
      breakdownCount: 70
    },
    // header row is row 4 in Excel (0-based index 3); data starts row 5 (index 4)
    DATA_START_ROW: 4,
    DATA_END_ROW: 368, // row 369 in Excel

    // ---- MTBF/MTTR ----
    // True MTBF = Operating Time / Breakdown Frequency (actual event count),
    // True MTTR = Unplanned Maintenance Time / Breakdown Frequency.
    // Daily_Input needs a "Breakdown Count (Unplanned)" column (BS) — the
    // number of distinct unplanned-breakdown EVENTS that day, not hours.
    // A companion template with this column pre-added and styled is
    // WL2_Performance_v7_with_breakdown_count.xlsx. Until a given row has
    // this filled in, engine.js falls back to the old proxy
    // (Operating Time / MTBF_req) so historical rows still show a number.

    // ---- display labels for the three production sources (S1/S2/S3 in the
    // workbook). Edit these if contractor names change. ----
    CONTRACTOR_LABELS: { S1: "Con-D (SQ)", S2: "Con-F (CPE)", S3: "Other" },

    // ---- truck fleets tracked in Daily_Input. `key` must match the suffix
    // used in fleetProd inside engine.js (do not rename `key`, only `label`). ----
    FLEET_DEFS: [
      { key: "HD785", label: "HD785 Fleet", contractor: "S1" },
      { key: "Tonly", label: "Tonly Fleet", contractor: "S1" },
      { key: "12DT-D", label: "12DT Fleet (D)", contractor: "S1" },
      { key: "10DT", label: "10DT Fleet", contractor: "S1" },
      { key: "12DT-F", label: "12DT Fleet (F)", contractor: "S2" },
      { key: "Tianma", label: "Tianma Fleet", contractor: "S2" }
    ],

    // ---- crushers tracked in Daily_Input, with rated capacity (BCM/hr) used
    // to compute the "vs capacity" utilization figure. Set capacity to 0 for
    // an inactive/decommissioned crusher (shows "n/a" instead of a %). ----
    CRUSHER_DEFS: [
      { key: "A", label: "Crusher A", capacity: 1703.67 },
      { key: "B", label: "Crusher B", capacity: 1703.67 },
      { key: "C", label: "Crusher C", capacity: 1703.67 },
      { key: "D", label: "Crusher D", capacity: 1222.5 },
      { key: "E", label: "Crusher E", capacity: 1222.5 },
      { key: "F", label: "Crusher F", capacity: 0 }
    ],

    // ---- downtime cause groups for the Time-Loss / Pareto sections. Each
    // group gets one fixed categorical color (do not reorder existing groups
    // or their colors will shift for causes already familiar to viewers).
    // `fields` map a Daily_Input field name -> its display label. ----
    CAUSE_GROUPS: [
      { key: "Idle", color: "var(--series-1)", fields: [
          ["idleEmpty", "Idle - Empty/Seq Stop"], ["idleRain", "Idle - Rain Fall"], ["idleWaitTruck", "Idle - Wait Truck"],
          ["idleSeqStart", "Idle - Sequence Start"], ["idleSPD", "Idle - SPD Switch"], ["idleSubsidence", "Idle - Subsidence"], ["idleOther", "Idle - Other"]
        ] },
      { key: "Relocation", color: "var(--series-2)", fields: [
          ["relocMainDump", "Reloc - Main Dump"], ["relocInPit", "Reloc - In-Pit"], ["relocConvShift", "Reloc - Conveyor Shift"], ["relocOther", "Reloc - Other"]
        ] },
      { key: "Planned Maintenance", color: "var(--series-3)", fields: [
          ["pmSystem", "Planned Maint - System"], ["pmOther", "Planned Maint - Other"]
        ] },
      { key: "Delay", color: "var(--series-4)", fields: [
          ["delayRain", "Delay - Clear Rainwater"], ["delayNoFeed", "Delay - Wait/No Feed"], ["delayTrunk", "Delay - Trunk to Crusher"], ["delayOther", "Delay - Other"]
        ] },
      { key: "Power Outage", color: "var(--series-5)", fields: [
          ["outPower", "Outage - Power Plant"], ["outMDS", "Outage - MDS/GIS"], ["outOther", "Outage - Other"]
        ] },
      { key: "Unplanned BD (Main)", color: "var(--series-6)", fields: [
          ["udOperation", "Operation BD"], ["udMechanical", "Mechanical BD"], ["udElectrical", "Electrical BD"], ["udPLC", "PLC/Comm BD"], ["udBelt", "Conveyor Belt BD"]
        ] },
      { key: "Unplanned BD (Crusher, detail)", color: "var(--series-7)", fields: [
          ["crABD", "Crusher A BD"], ["crBBD", "Crusher B BD"], ["crCBD", "Crusher C BD"], ["crDBD", "Crusher D BD"], ["crEBD", "Crusher E BD"], ["crFBD", "Crusher F BD"]
        ] },
      { key: "Unplanned BD (Face Conv, detail)", color: "var(--series-8)", fields: [
          ["faceM8", "Face Conv M8"], ["faceM8A", "Face Conv M8A"], ["faceM9", "Face Conv M9"], ["faceM10", "Face Conv M10"]
        ] }
    ],

    // ---- KPI status thresholds (percent). Used to color the KPI cards
    // good / warning / bad. Edit if management wants stricter/looser bands. ----
    STATUS_THRESHOLDS: {
      OA: { good: 80, warning: 65 },
      UO: { good: 85, warning: 70 },
      UA: { good: 85, warning: 70 },
      CU: { good: 90, warning: 75 },
      MA: { good: 85, warning: 70 },
      RE: { good: 90, warning: 75 }
    }
  };

  if (typeof module !== "undefined" && module.exports) module.exports = WL2_CONFIG;
  else root.WL2_CONFIG = WL2_CONFIG;
})(typeof window !== "undefined" ? window : this);
