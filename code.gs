/**************************************
 *  FINAL FAST MIS DASHBOARD (CACHE)
 **************************************/

const MIS = {
  DATA_SHEET: "MIS Scorer",
  CACHE_SHEET: "MIS_CACHE"
};

/**************************************
 * WEB APP ENTRY
 **************************************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("MIS Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**************************************
 * MAIN API (FAST)
 **************************************/
function misGetDashboardJson(periodKey) {

  const sh = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(MIS.CACHE_SHEET);

  if (!sh) {
    return { rows: [], totals: {} };
  }

  const data = sh.getDataRange().getValues();

  if (data.length < 2) {
    return { rows: [], totals: {} };
  }

  const rows = [];

  let tPlanned = 0, tDone = 0, tOnTime = 0;

  for (let i = 1; i < data.length; i++) {

    const r = data[i];

    const doer = r[0];
    const task = r[1];
    const planned = r[2];
    const done = r[3];
    const ontime = r[4];

    tPlanned += planned;
    tDone += done;
    tOnTime += ontime;

    rows.push({
      doer,
      task,
      planned,
      done,
      ontime
    });
  }

  return {
    generatedAt: new Date().toISOString(),
    totals: {
      planned: tPlanned,
      done: tDone,
      ontime: tOnTime
    },
    rows
  };
}

/**************************************
 * 🔥 HEAVY COMPUTATION (RUN IN BACKGROUND)
 **************************************/
function misPrecomputeCache() {

  const sh = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(MIS.DATA_SHEET);

  if (!sh) throw new Error("Data sheet not found");

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 2) return;

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  const blocks = misDetectBlocks_(headers);

  const useBlocks = blocks.length
    ? blocks
    : [{ planned: 0, actual: 1, doer: 3, task: 4 }];

  let maxIdx = 0;
  for (let b of useBlocks) {
    maxIdx = Math.max(maxIdx, b.planned, b.actual, b.doer, b.task);
  }

  const colCount = maxIdx + 1;

  const data = sh.getRange(2, 1, lastRow - 1, colCount).getValues();

  const doerTaskMap = new Map();

  for (let i = 0; i < data.length; i++) {

    const row = data[i];

    for (let j = 0; j < useBlocks.length; j++) {

      const b = useBlocks[j];

      const doer = row[b.doer];
      if (!doer) continue;

      const task = row[b.task] || "(No Task)";

      let taskMap = doerTaskMap.get(doer);
      if (!taskMap) {
        taskMap = new Map();
        doerTaskMap.set(doer, taskMap);
      }

      let c = taskMap.get(task);
      if (!c) {
        c = { planned: 0, done: 0, ontime: 0 };
        taskMap.set(task, c);
      }

      c.planned++;

      if (row[b.actual]) {
        c.done++;
        c.ontime++;
      }
    }
  }

  // 🔥 WRITE CACHE
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let cacheSheet = ss.getSheetByName(MIS.CACHE_SHEET);
  if (!cacheSheet) cacheSheet = ss.insertSheet(MIS.CACHE_SHEET);

  cacheSheet.clear();

  cacheSheet.appendRow(["Doer", "Task", "Planned", "Done", "OnTime"]);

  doerTaskMap.forEach((taskMap, doer) => {
    taskMap.forEach((c, task) => {
      cacheSheet.appendRow([doer, task, c.planned, c.done, c.ontime]);
    });
  });

  Logger.log("✅ Cache Updated Successfully");
}

/**************************************
 * HELPERS
 **************************************/
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

