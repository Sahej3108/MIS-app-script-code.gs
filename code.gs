/**************************************
 *  MIS DASHBOARD (OPTIMIZED WITH SAFE CACHE)
 **************************************/

const MIS = {
  TZ: "Asia/Kolkata",
  SPREADSHEET_ID: "",
  DATA_SHEET: "MIS Scorer",
  DASH_SHEET: "MIS DASHBOARD",
  REMARKS_SHEET: "MIS Remarks",
  NEXTPLAN_SHEET: "MIS Next Plan",
  DONE_TODAY_SHEET: "MIS Done Today",
  DATE_FORMAT: "DMY",
  MAX_REMARKS_SHOW: 3
};

/**************************************
 * MENU
 **************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 MIS Score")
    .addItem("Refresh MIS DASHBOARD Sheet", "misRefreshDashboardSheet")
    .addToUi();
}

/**************************************
 * WEB APP
 **************************************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("MIS Score Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**************************************
 * 🚀 SAFE CACHE LAYER (ONLY HEAVY PART)
 **************************************/
function misGetCachedCounts_(period) {

  const cache = CacheService.getScriptCache();

  const key = "MIS_COUNT_" +
    Utilities.formatDate(period.start, MIS.TZ, "yyyyMMdd") + "_" +
    Utilities.formatDate(period.end, MIS.TZ, "yyyyMMdd");

  const cached = cache.get(key);

  if (cached) {
    const parsed = JSON.parse(cached);

    const map = new Map();
    Object.keys(parsed.map).forEach(doer => {
      const taskMap = new Map();
      Object.keys(parsed.map[doer]).forEach(task => {
        taskMap.set(task, parsed.map[doer][task]);
      });
      map.set(doer, taskMap);
    });

    return {
      doerTaskMap: map,
      totals: parsed.totals
    };
  }

  // 🔥 original heavy function
  const built = misBuildDoerTaskCounts_(period);

  const obj = {};
  built.doerTaskMap.forEach((taskMap, doer) => {
    obj[doer] = {};
    taskMap.forEach((c, task) => {
      obj[doer][task] = c;
    });
  });

  cache.put(
    key,
    JSON.stringify({ map: obj, totals: built.totals }),
    600 // 10 min
  );

  return built;
}

/**************************************
 * MAIN DASHBOARD API
 **************************************/
function misGetDashboardJson(periodKey) {

  const now = new Date();

  const periodMap = {
    CURRENT_WEEK:  misGetWeekPeriod_("CURRENT_WEEK", now, 0),
    LAST_WEEK:     misGetWeekPeriod_("LAST_WEEK", now, -1),
    WEEK_2_AGO:    misGetWeekPeriod_("WEEK_2_AGO", now, -2),
    WEEK_3_AGO:    misGetWeekPeriod_("WEEK_3_AGO", now, -3),
    WEEK_4_AGO:    misGetWeekPeriod_("WEEK_4_AGO", now, -4),
    NEXT_WEEK:     misGetWeekPeriod_("NEXT_WEEK", now, 1),
    MONTH_TO_DATE: misGetMonthToDatePeriod_(now),
  };

  const period = periodMap[periodKey] || periodMap.CURRENT_WEEK;

  // ⚡ USE CACHE HERE
  const built = misGetCachedCounts_(period);

  const planWeek = {
    key: "PLAN_WEEK",
    start: period.start,
    end: period.end
  };

  const planTeamMap = misGetNextPlanMap_(planWeek);

  const todayStr = Utilities.formatDate(now, MIS.TZ, "yyyy-MM-dd");
  const doneTodaySet = (period.key === "LAST_WEEK")
    ? misGetDoneTodaySet_(todayStr, period.key)
    : new Set();

  const rows = [];
  const doers = Array.from(built.doerTaskMap.keys()).sort();

  for (let doer of doers) {

    if (doneTodaySet.has(doer)) continue;

    const taskMap = built.doerTaskMap.get(doer);
    const tasks = [];

    let dPlanned = 0, dDone = 0, dOnTime = 0;

    taskMap.forEach((c, task) => {

      const planned = c.planned;
      const done = c.done;
      const ontime = c.ontime;

      dPlanned += planned;
      dDone += done;
      dOnTime += ontime;

      const np = planTeamMap[`${doer} - ${task}`] || {};

      tasks.push({
        task,
        planned,
        done,
        ontime,
        notDonePct: misNegPct_(planned, done),
        notOnTimePct: misNegPct_(planned, ontime),
        nextPlannedTarget: np.target ?? null,
        nextPlanRemark: np.remark || ""
      });
    });

    rows.push({
      doer,
      misScore: Number(((misNegPct_(dPlanned, dDone) + misNegPct_(dPlanned, dOnTime)) / 2).toFixed(1)),
      tasks
    });
  }

  return {
    generatedAt: new Date().toISOString(),
    totals: built.totals,
    rows
  };
}

/**************************************
 * ORIGINAL CORE FUNCTION (UNCHANGED)
 **************************************/
function misBuildDoerTaskCounts_(period) {
  const sh = misGetSS_().getSheetByName(MIS.DATA_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const blocks = misDetectBlocks_(headers);

  let maxIdx = 0;
  blocks.forEach(b => {
    maxIdx = Math.max(maxIdx, b.planned, b.actual, b.doer, b.task);
  });

  const data = sh.getRange(2, 1, lastRow - 1, maxIdx + 1).getValues();

  const map = new Map();
  let tPlanned = 0, tDone = 0, tOnTime = 0;

  for (let row of data) {
    for (let b of blocks) {

      const doer = row[b.doer];
      if (!doer) continue;

      const planned = misParseDateFast_(row[b.planned]);
      if (!planned) continue;

      if (planned < period.start || planned >= period.end) continue;

      const task = row[b.task] || "(No Task)";

      if (!map.has(doer)) map.set(doer, new Map());
      const tm = map.get(doer);

      if (!tm.has(task)) tm.set(task, { planned: 0, done: 0, ontime: 0 });
      const c = tm.get(task);

      c.planned++; tPlanned++;

      const actual = misParseDateFast_(row[b.actual]);
      if (actual) {
        c.done++; tDone++;
        if (actual <= planned) {
          c.ontime++; tOnTime++;
        }
      }
    }
  }

  return { doerTaskMap: map, totals: { planned: tPlanned, done: tDone, ontime: tOnTime } };
}

/**************************************
 * HELPERS (UNCHANGED)
 **************************************/
function misNegPct_(p, a) {
  if (!p) return 0;
  if (a >= p) return 0;
  return Number((-((p - a) / p) * 100).toFixed(1));
}

function misGetSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function misDetectBlocks_(headers) {
  const blocks = [];
  for (let i = 0; i < headers.length; i++) {
    if (/planned/i.test(headers[i]) && /actual/i.test(headers[i + 1])) {
      blocks.push({
        planned: i,
        actual: i + 1,
        doer: i + 2,
        task: i + 3
      });
      i += 3;
    }
  }
  return blocks;
}

function misParseDateFast_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function misGetWeekPeriod_(key, now, offset) {
  const start = new Date(now);
  start.setDate(start.getDate() - start.getDay() + 1 + offset * 7);
  start.setHours(0, 0, 0, 0);

  return {
    key,
    start,
    end: new Date(start.getTime() + 7 * 86400000)
  };
}

function misGetMonthToDatePeriod_(now) {
  return {
    key: "MONTH_TO_DATE",
    start: new Date(now.getFullYear(), now.getMonth(), 1),
    end: new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1)
  };
}
