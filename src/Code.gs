// ===== EXPENSE TRACKER - VERSION 1.0 =====
// New in v28:
//   - SpreadsheetApp.flush() before each chart is inserted. The helper values
//     were written moments earlier in the same run and had not been committed
//     yet, so the donut was reading an empty range and rendering blank.
//   - Removed the unused kpiRow() helper.
// Your entry rows (columns A-E on each month tab) are only ever read, never
// modified, so existing data is untouched.

const T = {
  ink:       "#2B2118",
  inkSoft:   "#7A6A5C",
  cream:     "#FDFBF7",
  card:      "#FFFFFF",
  band:      "#F6F0E8",
  line:      "#EADFD2",
  plum:      "#5B4B6D",
  green:     "#2F8F5B",
  greenSoft: "#E4F3EA",
  red:       "#C4553D",
  redSoft:   "#FBEAE5",
  amber:     "#D89B3C",
  amberSoft: "#FBF0DC",
  teal:      "#2E7D7B",
  font:      "Nunito"
};

const MONEY = "$#,##0.00";
const MONEY0 = "$#,##0";
const DATE_FMT = "d mmm yyyy";   // 12 Aug 2026
const DATA_SHEET = "_Data";

// Dark palette for the Dashboard only; month tabs stay light for readability.
const D = {
  bg:       "#15171C",
  card:     "#1E2128",
  line:     "#2C313A",
  text:     "#F2F4F7",
  muted:    "#8A8F98",
  purple:   "#A78BFA",
  blue:     "#60A5FA",
  green:    "#34D399",
  pink:     "#F472B6",
  heatLow:  "#1E2128",
  heatHigh: "#7C5CE6",
  font:     "Nunito"
};

const DARK_PALETTE = [
  "#A78BFA", "#7C6BF0", "#60A5FA", "#FB923C", "#F87171",
  "#34D399", "#F472B6", "#38BDF8", "#FBBF24"
];

// Helper block columns (kept far right, visible so charts can read them)
const HELP_EXP = 26;  // Z  — month tab: expense by category
const HELP_INC = 29;  // AC — month tab: income by source
const DASH_CAT = 26;  // Z  — dashboard: selected-month categories
const DASH_TRD = 29;  // AC — dashboard: last-12-months trend
const DASH_SEL = 34;  // AH — dashboard: selected-month earned/spent/net
const DASH_CAL = 40;  // AN — dashboard: calendar date helpers

const EXPENSE_PALETTE = [
  "#C4553D", "#DE7B4B", "#E3A857", "#A85751",
  "#8C5A6E", "#D2694F", "#B8763C", "#9C4F52", "#E0B072"
];

const INCOME_PALETTE = [
  "#2F8F5B", "#2E7D7B", "#57A773", "#3E8E7E",
  "#6BB08A", "#1F6F5C", "#7FC098", "#48A9A6", "#95CBA6"
];

const CHIP_COLORS = [
  "#FBEAE5", "#FBF0DC", "#E9F1F7", "#F1EAF4",
  "#E4F3EA", "#FDF0E8", "#EFF0E6", "#F7EAEF", "#E8F2F1"
];

// ---------- Routing ----------

function doGet(e) {
  if (e && e.parameter && e.parameter.repair) return repairAllSheets();
  if (e && e.parameter && e.parameter.echo) {
    // Writes nothing — just reports exactly what arrived, so you can point the
    // Shortcut here temporarily to see what it is actually sending.
    const p = e.parameter;
    return ContentService.createTextOutput(
      "v28 ECHO\n" +
      "amount      = " + JSON.stringify(p.amount || p.am) + "\n" +
      "type        = " + JSON.stringify(p.type || p.ty) + "\n" +
      "category    = " + JSON.stringify(p.category || p.ct) + "\n" +
      "description = " + JSON.stringify(p.description || p.ds) + "\n" +
      "date        = " + JSON.stringify(p.date || p.dt) + "\n\n" +
      "all keys: " + JSON.stringify(Object.keys(p)));
  }
  if (e && e.parameter && e.parameter.debug) {
    return ContentService.createTextOutput(
      "v28 DEBUG — received parameters: " + JSON.stringify(e.parameter));
  }
  if (e && e.parameter && (e.parameter.amount || e.parameter.am)) {
    return handleEntry(e.parameter);
  }
  return ContentService.createTextOutput(
    "v28 — Expense Tracker API is running. No amount parameter was received.");
}

function doPost(e) {
  try {
    return handleEntry(JSON.parse(e.postData.contents));
  } catch (err) {
    return jsonOut({ status: "error", version: "v28", message: err.toString() });
  }
}

function handleEntry(data) {
  try {
    // Accept short aliases as well as the full names. Note the aliases are two
    // letters: Google reserves several single-letter query parameters for its
    // own routing, and using "c" made requests fail before reaching this code.
    const val = (long, short) => {
      const v = data[long] !== undefined && data[long] !== ""
        ? data[long] : data[short];
      return v === undefined ? "" : v;
    };

    const amountRaw = val("amount", "am");
    const typeRaw = val("type", "ty");
    const categoryRaw = val("category", "ct");
    const descRaw = val("description", "ds");
    const dateRaw = val("date", "dt");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    const tz = ss.getSpreadsheetTimeZone();

    const entryDate = parseEntryDate(dateRaw) || now;
    const monthName = Utilities.formatDate(entryDate, tz, "MMMM yyyy");

    let sheet = ss.getSheetByName(monthName);
    if (!sheet) {
      sheet = ss.insertSheet(monthName);
      buildMonthSheet(sheet);
    } else {
      ensureSummaryBlock(sheet, true);
    }

    const amt = parseFloat(amountRaw);
    const clean = isNaN(amt) ? 0 : amt;
    const target = firstFreeDataRow(sheet);

    sheet.getRange(target, 1, 1, 5).setValues([[
      entryDate,
      typeRaw,
      categoryRaw,
      descRaw,
      clean
    ]]);

    sortEntriesByDate(sheet);
    styleMonthRows(sheet);
    createCategoryCharts(sheet, monthName);
    orderMonthTabs(ss);
    rebuildDataSheet(ss);
    updateDashboard(ss);

    return jsonOut({
      status: "success", version: "v28",
      wrote: {
        month: monthName,
        date: Utilities.formatDate(entryDate, tz, "yyyy-MM-dd HH:mm"),
        backdated: !!parseEntryDate(dateRaw),
        type: typeRaw, category: categoryRaw,
        description: descRaw, amount: clean
      }
    });
  } catch (err) {
    return jsonOut({ status: "error", version: "v28", message: err.toString() });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Row helpers ----------

function lastDataRow(sheet) {
  const v = sheet.getRange("A:A").getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]).trim() !== "") return i + 1;
  }
  return 1;
}

function firstFreeDataRow(sheet) { return lastDataRow(sheet) + 1; }

function isMonthSheet(name) {
  return name !== "Dashboard" && name !== "Sheet1" && name !== DATA_SHEET;
}

// Totals read straight from the entry rows. Deliberately not read from the
// summary cells: those live in the same rows as the data, so deleting rows can
// wipe them out and leave #REF! behind.
function monthTotals(sheet) {
  const last = lastDataRow(sheet);
  if (last < 2) return { inc: 0, exp: 0 };

  const rows = sheet.getRange(2, 2, last - 1, 4).getValues(); // B..E
  let inc = 0, exp = 0;

  rows.forEach(r => {
    const type = String(r[0]).trim().toLowerCase();
    const amt = Number(r[3]) || 0;
    if (type === "income") inc += amt;
    else if (type === "expense") exp += amt;
  });

  return { inc: inc, exp: exp };
}

// ---------- In-sheet menu ----------

// Adds a "Tracker" menu when the spreadsheet opens.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💰 Tracker")
    .addItem("Refresh dashboard", "menuRefresh")
    .addItem("Delete selected entries", "menuDeleteSelected")
    .addSeparator()
    .addItem("Enable auto-refresh", "installAutoRefresh")
    .addItem("Disable auto-refresh", "removeAutoRefresh")
    .addToUi();
}

// Removes entries safely. Deleting whole rows by hand also removes the summary
// cells in G/H that sit on those same rows, which is what leaves #REF! behind.
function menuDeleteSelected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (!isMonthSheet(sheet.getName())) {
    ui.alert("Select rows on a month tab first.");
    return;
  }

  const sel = sheet.getActiveRange();
  const first = Math.max(sel.getRow(), 2);
  const last = sel.getLastRow();
  if (last < 2) {
    ui.alert("Select one or more entry rows first.");
    return;
  }

  const count = last - first + 1;
  const answer = ui.alert(
    "Delete entries",
    "Remove " + count + " entr" + (count === 1 ? "y" : "ies") + " from " +
    sheet.getName() + "?",
    ui.ButtonSet.YES_NO);

  if (answer !== ui.Button.YES) return;

  // Clear only columns A–E, then close the gap, leaving G/H untouched.
  sheet.getRange(first, 1, count, 5).clearContent();

  const lastRow = lastDataRow(sheet);
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues()
      .filter(r => String(r[0]).trim() !== "");
    sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    if (data.length) sheet.getRange(2, 1, data.length, 5).setValues(data);
  }

  refreshEverything();
  ss.toast("Removed " + count + " entr" + (count === 1 ? "y" : "ies") + ".",
    "Tracker", 3);
}

function menuRefresh() {
  refreshEverything();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Dashboard rebuilt.", "Tracker", 3);
}

// Shared rebuild used by the menu, the trigger and ?repair=1.
function refreshEverything() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheets().forEach(s => {
    if (!isMonthSheet(s.getName())) return;
    try { s.showColumns(1, s.getMaxColumns()); } catch (e) {}
    applyMonthChrome(s);
    ensureSummaryBlock(s, true);
    convertLegacyDates(s);
    sortEntriesByDate(s);
    styleMonthRows(s);
    createCategoryCharts(s, s.getName());
  });

  orderMonthTabs(ss);
  rebuildDataSheet(ss);
  updateDashboard(ss);
}

// Rebuilds automatically whenever the spreadsheet changes (rows deleted, edits
// made by hand). Installed on demand so it isn't running unless you want it.
function installAutoRefresh() {
  removeAutoRefresh();
  ScriptApp.newTrigger("autoRefresh")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Auto-refresh is on. Manual edits will rebuild the dashboard.", "Tracker", 5);
}

function removeAutoRefresh() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "autoRefresh") ScriptApp.deleteTrigger(t);
  });
}

function autoRefresh(e) {
  // Only react to structural changes; ignore formatting and our own writes.
  if (e && e.changeType &&
      ["EDIT", "INSERT_ROW", "REMOVE_ROW", "REMOVE_GRID", "OTHER"].indexOf(e.changeType) < 0) {
    return;
  }
  try { refreshEverything(); } catch (err) { /* never block the user's edit */ }
}

// ---------- Repair ----------

function repairAllSheets() {
  const names = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(s => isMonthSheet(s.getName()))
    .map(s => s.getName());

  refreshEverything();

  return ContentService.createTextOutput(
    "v28 REPAIR complete. Sheets rebuilt: " + (names.join(", ") || "none"));
}

// ---------- Consolidated data sheet (feeds dashboard formulas) ----------

function rebuildDataSheet(ss) {
  let ds = ss.getSheetByName(DATA_SHEET);
  if (!ds) ds = ss.insertSheet(DATA_SHEET);
  ds.clear();

  // Day is carried through so the calendar heatmap can total per-day spend.
  const out = [["Month", "Type", "Category", "Amount", "Day"]];

  ss.getSheets().forEach(s => {
    if (!isMonthSheet(s.getName())) return;
    const last = lastDataRow(s);
    if (last < 2) return;
    const vals = s.getRange(2, 1, last - 1, 5).getValues(); // A..E
    vals.forEach(v => {
      const d = v[0];
      const day = (d instanceof Date) ? d.getDate() : "";
      out.push([
        s.getName(), String(v[1]).trim(), String(v[2]).trim(),
        Number(v[4]) || 0, day
      ]);
    });
  });

  ds.getRange(1, 1, out.length, 5).setValues(out);
  ds.hideSheet(); // formulas read hidden sheets fine; only charts can't
}

// ---------- Date handling ----------

// Accepts "yyyy-MM-dd", "yyyy-MM-dd HH:mm", "dd/MM/yyyy" or "MM/dd/yyyy".
// Returns null when nothing usable was supplied, so the caller falls back to now.
function parseEntryDate(raw) {
  if (!raw) return null;
  const s = String(raw).trim();
  if (!s) return null;

  // ISO-ish: 2026-08-14 or 2026-08-14 19:30
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (m) {
    return new Date(+m[1], +m[2] - 1, +m[3], m[4] ? +m[4] : 12, m[5] ? +m[5] : 0);
  }

  // Slashed: assume day/month/year when the first part can't be a month
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const a = +m[1], b = +m[2], y = +m[3];
    const day = a > 12 ? a : b;
    const mon = a > 12 ? b : a;
    return new Date(y, mon - 1, day, 12, 0);
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// Keeps rows in date order, which matters once entries are back-dated.
function sortEntriesByDate(sheet) {
  const last = lastDataRow(sheet);
  if (last < 3) return; // nothing to reorder
  sheet.getRange(2, 1, last - 1, 5).sort({ column: 1, ascending: true });
}

// Puts month tabs in chronological order after the Dashboard.
function orderMonthTabs(ss) {
  const months = ss.getSheets()
    .filter(s => isMonthSheet(s.getName()))
    .sort((a, b) => monthKey(a.getName()) - monthKey(b.getName()));

  let pos = 2; // Dashboard sits at 1
  months.forEach(s => {
    ss.setActiveSheet(s);
    ss.moveActiveSheet(pos);
    pos++;
  });
}

// ---------- Month sheet ----------

function buildMonthSheet(sheet) {
  applyMonthChrome(sheet);
  ensureSummaryBlock(sheet, true);
}

function applyMonthChrome(sheet) {
  sheet.setHiddenGridlines(true);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setBackground(T.cream);

  sheet.getRange("A1:E1")
    .setValues([["DATE", "TYPE", "CATEGORY", "DESCRIPTION", "AMOUNT"]])
    .setBackground(T.plum).setFontColor("#FFFFFF")
    .setFontFamily(T.font).setFontSize(10).setFontWeight("bold")
    .setVerticalAlignment("middle").setHorizontalAlignment("left");
  sheet.getRange("E1").setHorizontalAlignment("right");

  sheet.setRowHeight(1, 40);
  sheet.setColumnWidth(1, 155);
  sheet.setColumnWidth(2, 105);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 240);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 30);
  sheet.setFrozenRows(1);
}

function ensureSummaryBlock(sheet, force) {
  // Always rebuilt rather than skipped when it "looks" present: deleting entry
  // rows can shift or wipe part of this block, leaving it half-valid.
  sheet.getRange("G1:I8").clearContent().clearFormat().setBackground(T.cream);

  sheet.getRange("G1").setValue("THIS MONTH")
    .setFontFamily(T.font).setFontSize(10).setFontWeight("bold")
    .setFontColor(T.inkSoft);

  const cards = [
    { row: 2, label: "TOTAL INCOME",  f: '=SUMIF(B:B,"Income",E:E)',  bg: T.greenSoft, fg: T.green },
    { row: 3, label: "TOTAL EXPENSE", f: '=SUMIF(B:B,"Expense",E:E)', bg: T.redSoft,   fg: T.red   },
    { row: 4, label: "NET",
      f: '=SUMIF(B:B,"Income",E:E)-SUMIF(B:B,"Expense",E:E)',
      bg: T.band, fg: T.ink }
  ];

  cards.forEach(c => {
    sheet.getRange(c.row, 7).setValue(c.label)
      .setFontFamily(T.font).setFontSize(10).setFontWeight("bold")
      .setFontColor(T.inkSoft).setBackground(c.bg).setVerticalAlignment("middle");

    sheet.getRange(c.row, 8).setFormula(c.f)
      .setNumberFormat(MONEY)
      .setFontFamily(T.font).setFontSize(14).setFontWeight("bold")
      .setFontColor(c.fg).setBackground(c.bg)
      .setHorizontalAlignment("right").setVerticalAlignment("middle");

    sheet.setRowHeight(c.row, 38);
  });

  sheet.getRange("G2:H4").setBorder(true, true, true, true, false, true,
    T.line, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 140);

  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
      .setFontColor(T.red).setBackground(T.redSoft)
      .setRanges([sheet.getRange("H4")]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setFontColor(T.green).setBackground(T.greenSoft)
      .setRanges([sheet.getRange("H4")]).build()
  ]);
}

// Older versions stored the date as text ("2026-08-14 19:29"). Convert any
// leftovers to real Date values so sorting and formatting behave.
function convertLegacyDates(sheet) {
  const last = lastDataRow(sheet);
  if (last < 2) return;

  const range = sheet.getRange(2, 1, last - 1, 1);
  const vals = range.getValues();
  let changed = false;

  for (let i = 0; i < vals.length; i++) {
    const v = vals[i][0];
    if (v instanceof Date) continue;
    const parsed = parseEntryDate(v);
    if (parsed) { vals[i][0] = parsed; changed = true; }
  }

  if (changed) range.setValues(vals);
}

function styleMonthRows(sheet) {
  convertLegacyDates(sheet);
  const last = lastDataRow(sheet);
  const maxRows = sheet.getMaxRows();

  if (maxRows > last) {
    sheet.getRange(last + 1, 1, maxRows - last, 5).clearFormat().setBackground(T.cream);
  }
  if (last < 2) return;

  const n = last - 1;
  const body = sheet.getRange(2, 1, n, 5);

  body.setFontFamily(T.font).setFontSize(11).setFontColor(T.ink)
      .setVerticalAlignment("middle").setBackground(T.card);

  for (let r = 2; r <= last; r++) sheet.setRowHeight(r, 34);

  sheet.getRange(2, 1, n, 1).setFontColor(T.inkSoft).setFontSize(10)
    .setNumberFormat(DATE_FMT).setHorizontalAlignment("left");
  sheet.getRange(2, 5, n, 1).setNumberFormat(MONEY).setFontSize(12)
    .setFontWeight("bold").setHorizontalAlignment("right");

  body.setBorder(null, null, null, null, null, true,
    T.line, SpreadsheetApp.BorderStyle.SOLID);

  const rows = sheet.getRange(2, 2, n, 4).getValues();
  const chips = {};
  let ci = 0;

  for (let i = 0; i < n; i++) {
    const r = i + 2;
    const type = String(rows[i][0]).trim();
    const cat = String(rows[i][1]).trim();

    const tc = sheet.getRange(r, 2);
    tc.setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center");
    if (type === "Income") tc.setFontColor(T.green).setBackground(T.greenSoft);
    else if (type === "Expense") tc.setFontColor(T.red).setBackground(T.redSoft);

    sheet.getRange(r, 5).setFontColor(type === "Income" ? T.green : T.ink);

    if (cat) {
      if (!(cat in chips)) { chips[cat] = CHIP_COLORS[ci % CHIP_COLORS.length]; ci++; }
      sheet.getRange(r, 3).setBackground(chips[cat])
        .setFontSize(10).setFontWeight("bold").setFontColor(T.ink);
    }

    sheet.getRange(r, 4).setFontColor(T.inkSoft).setFontSize(11);
  }
}

function createCategoryCharts(sheet, monthName) {
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  const last = lastDataRow(sheet);
  // Wipe helper blocks, including legacy positions used by older versions (J..P)
  sheet.getRange(1, 10, 300, 7).clearContent().clearFormat();
  sheet.getRange(1, HELP_EXP, 300, 2).clearContent();
  sheet.getRange(1, HELP_INC, 300, 2).clearContent();
  if (last < 2) return;

  const rows = sheet.getRange(2, 2, last - 1, 4).getValues();
  const spend = {}, earn = {};

  rows.forEach(r => {
    const type = String(r[0]).trim().toLowerCase();
    const cat = String(r[1]).trim();
    const amt = Number(r[3]) || 0;
    if (!cat) return;
    if (type === "expense") spend[cat] = (spend[cat] || 0) + amt;
    else if (type === "income") earn[cat] = (earn[cat] || 0) + amt;
  });

  const sRows = sortedEntries(spend);
  const eRows = sortedEntries(earn);

  if (sRows.length) {
    sheet.getRange(1, HELP_EXP, 1, 2).setValues([["Category", "Spent"]]);
    sheet.getRange(2, HELP_EXP, sRows.length, 2).setValues(sRows);
    SpreadsheetApp.flush();   // commit before the chart reads the range
    sheet.insertChart(donut(sheet,
      sheet.getRange(1, HELP_EXP, sRows.length + 1, 2),
      "Where the money went — " + monthName, EXPENSE_PALETTE, 6, 7));
  }

  if (eRows.length) {
    sheet.getRange(1, HELP_INC, 1, 2).setValues([["Source", "Earned"]]);
    sheet.getRange(2, HELP_INC, eRows.length, 2).setValues(eRows);
    SpreadsheetApp.flush();
    sheet.insertChart(donut(sheet,
      sheet.getRange(1, HELP_INC, eRows.length + 1, 2),
      "Where the money came from — " + monthName, INCOME_PALETTE, 15, 7));
  }
}

function donut(sheet, range, title, palette, posRow, posCol) {
  return sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(range)
    .setPosition(posRow, posCol, 0, 0)
    .setOption("title", title)
    .setOption("fontName", T.font)
    .setOption("titleTextStyle", { fontSize: 13, bold: true, color: T.ink, fontName: T.font })
    .setOption("colors", palette)
    .setOption("pieHole", 0.58)
    .setOption("pieSliceBorderColor", T.card)
    .setOption("legend", { position: "right",
      textStyle: { fontSize: 11, color: T.inkSoft, fontName: T.font } })
    .setOption("width", 460).setOption("height", 290)
    .setOption("backgroundColor", T.card)
    .build();
}

function sortedEntries(o) {
  return Object.keys(o).map(k => [k, o[k]]).sort((a, b) => b[1] - a[1]);
}

// ---------- Dashboard ----------

function updateDashboard(ss) {
  let d = ss.getSheetByName("Dashboard");
  if (!d) d = ss.insertSheet("Dashboard", 0);

  const prevPick = d.getRange("B7").getValue();

  d.clear();
  d.getRange(1, 1, d.getMaxRows(), d.getMaxColumns()).clearDataValidations();
  d.getCharts().forEach(c => d.removeChart(c));
  d.clearConditionalFormatRules();
  try { d.showColumns(1, d.getMaxColumns()); } catch (e) {}
  d.setHiddenGridlines(true);
  d.getRange(1, 1, d.getMaxRows(), d.getMaxColumns()).setBackground(D.bg);

  // ---- Gather ----
  const rows = [];
  const catSet = {};
  ss.getSheets().forEach(s => {
    if (!isMonthSheet(s.getName())) return;
    // Empty months are kept so they still show as zero in the picker and charts.
    const t = monthTotals(s);
    rows.push({ name: s.getName(), inc: t.inc, exp: t.exp, net: t.inc - t.exp });

    const last = lastDataRow(s);
    if (last < 2) return;
    s.getRange(2, 2, last - 1, 2).getValues().forEach(v => {
      if (String(v[0]).trim().toLowerCase() === "expense" && String(v[1]).trim()) {
        catSet[String(v[1]).trim()] = true;
      }
    });
  });
  rows.sort((a, b) => monthKey(a.name) - monthKey(b.name));

  const names = rows.map(m => m.name);
  const pick = names.indexOf(prevPick) >= 0
    ? prevPick : (names.length ? names[names.length - 1] : "");

  const Q = "'" + DATA_SHEET + "'!";

  // ---- Header ----
  d.getRange("A1:M1").merge().setValue("MY EXPENSE TRACKER")
    .setBackground(D.bg).setFontColor(D.text)
    .setFontSize(22).setFontWeight("bold").setFontFamily(D.font)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  d.setRowHeight(1, 52);

  // ---- Month picker ----
  d.getRange("A7").setValue("SHOWING")
    .setFontFamily(D.font).setFontSize(9).setFontWeight("bold")
    .setFontColor(D.muted).setVerticalAlignment("middle")
    .setHorizontalAlignment("right");
  d.setRowHeight(6, 24);

  const picker = d.getRange("B7:C7").merge();
  picker.clearDataValidations();
  picker.setNumberFormat("@");   // else "August 2026" is parsed as a date
  picker.setValue(pick)
    .setFontFamily(D.font).setFontSize(14).setFontWeight("bold")
    .setFontColor(D.text).setBackground(D.card)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  if (names.length) {
    try {
      picker.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(names, true).setAllowInvalid(true)
        .setHelpText("Pick a month").build());
    } catch (e) { picker.clearDataValidations(); }
  }
  d.setRowHeight(7, 38);

  // ---- Hero numbers ----
  const hero = [
    { col: 4, label: "SPENT",  f: '=SUMIFS(' + Q + 'D:D,' + Q + 'A:A,$B$7,' + Q + 'B:B,"Expense")', fg: D.pink },
    { col: 6, label: "EARNED", f: '=SUMIFS(' + Q + 'D:D,' + Q + 'A:A,$B$7,' + Q + 'B:B,"Income")',  fg: D.green },
    { col: 8, label: "LEFT",   f: "=F7-D7", fg: D.purple },
    { col: 10, label: "ENTRIES", f: '=COUNTIF(' + Q + 'A:A,$B$7)', fg: D.blue, plain: true }
  ];

  hero.forEach(h => {
    d.getRange(6, h.col, 1, 2).merge().setValue(h.label)
      .setFontFamily(D.font).setFontSize(9).setFontWeight("bold")
      .setFontColor(D.muted).setBackground(D.card)
      .setHorizontalAlignment("center").setVerticalAlignment("bottom");

    const v = d.getRange(7, h.col, 1, 2).merge().setFormula(h.f)
      .setFontFamily(D.font).setFontSize(20).setFontWeight("bold")
      .setFontColor(h.fg).setBackground(D.card)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    if (!h.plain) v.setNumberFormat(MONEY0);

    d.getRange(6, h.col, 2, 2).setBorder(
      true, true, true, true, false, false, D.line, SpreadsheetApp.BorderStyle.SOLID);
  });

  // ---- All-time strip ----
  const allIn = rows.reduce((a, m) => a + m.inc, 0);
  const allOut = rows.reduce((a, m) => a + m.exp, 0);

  d.getRange("A3:A4").merge().setValue("ALL TIME")
    .setFontFamily(D.font).setFontSize(10).setFontWeight("bold")
    .setFontColor(D.muted).setVerticalAlignment("middle");

  const allTime = [
    { col: 4, label: "SPENT",  value: allOut,        fg: D.pink },
    { col: 6, label: "EARNED", value: allIn,         fg: D.green },
    { col: 8, label: "NET",    value: allIn - allOut,
      fg: (allIn - allOut) < 0 ? D.pink : D.green },
    { col: 10, label: "MONTHS", value: rows.length,  fg: D.blue, plain: true }
  ];

  allTime.forEach(a => {
    d.getRange(3, a.col, 1, 2).merge().setValue(a.label)
      .setFontFamily(D.font).setFontSize(8).setFontWeight("bold")
      .setFontColor(D.muted).setBackground(D.card)
      .setHorizontalAlignment("center").setVerticalAlignment("bottom");

    const v = d.getRange(4, a.col, 1, 2).merge().setValue(a.value)
      .setFontFamily(D.font).setFontSize(15).setFontWeight("bold")
      .setFontColor(a.fg).setBackground(D.card)
      .setHorizontalAlignment("center").setVerticalAlignment("top");
    if (!a.plain) v.setNumberFormat(MONEY0);

    d.getRange(3, a.col, 2, 2).setBorder(
      true, true, true, true, false, false, D.line, SpreadsheetApp.BorderStyle.SOLID);
  });
  d.setRowHeight(3, 22);
  d.setRowHeight(4, 34);

  let rules = [];

  // ---- Category breakdown ----
  // Written as plain values, not formulas: Sheets charts will not render from
  // a formula-driven range created in the same run. Cost is that this section
  // refreshes on Refresh rather than the instant the dropdown changes.
  const pickTotals = {};
  const pickSheet = ss.getSheetByName(pick);
  if (pickSheet) {
    const pl = lastDataRow(pickSheet);
    if (pl >= 2) {
      pickSheet.getRange(2, 2, pl - 1, 4).getValues().forEach(v => {
        if (String(v[0]).trim().toLowerCase() !== "expense") return;
        const c = String(v[1]).trim();
        if (!c) return;
        pickTotals[c] = (pickTotals[c] || 0) + (Number(v[3]) || 0);
      });
    }
  }

  const catRows = Object.keys(pickTotals)
    .map(k => [k, pickTotals[k]])
    .sort((a, b) => b[1] - a[1]);

  const pickSpend = catRows.reduce((a, r) => a + r[1], 0);

  d.getRange(1, DASH_CAT, 300, 2).clearContent();

  const listTop = 12;
  d.getRange(listTop - 1, 1).setValue("WHERE IT WENT")
    .setFontFamily(D.font).setFontSize(10).setFontWeight("bold")
    .setFontColor(D.muted);
  d.setRowHeight(listTop - 1, 30);

  if (catRows.length) {
    d.getRange(1, DASH_CAT, 1, 2).setValues([["Category", "Spent"]]);
    d.getRange(2, DASH_CAT, catRows.length, 2).setValues(catRows);
    SpreadsheetApp.flush();   // commit before the donut reads the range

    for (let i = 0; i < catRows.length; i++) {
      const r = listTop + i;
      const name = catRows[i][0];
      const amt = catRows[i][1];

      d.getRange(r, 1, 1, 2).merge().setValue(name)
        .setFontFamily(D.font).setFontSize(12).setFontWeight("bold")
        .setFontColor(DARK_PALETTE[i % DARK_PALETTE.length]).setBackground(D.card)
        .setVerticalAlignment("middle").setHorizontalAlignment("left");

      d.getRange(r, 3).setValue(
        Object.keys(pickTotals).length
          ? countEntries(pickSheet, name) : 0)
        .setNumberFormat("0")
        .setFontFamily(D.font).setFontSize(10)
        .setFontColor(D.muted).setBackground(D.card)
        .setHorizontalAlignment("center").setVerticalAlignment("middle");

      d.getRange(r, 4).setValue(amt).setNumberFormat(MONEY0)
        .setFontFamily(D.font).setFontSize(12).setFontWeight("bold")
        .setFontColor(D.text).setBackground(D.card)
        .setHorizontalAlignment("right").setVerticalAlignment("middle");

      d.getRange(r, 5).setValue(pickSpend ? amt / pickSpend : 0)
        .setNumberFormat("0.0%")
        .setFontFamily(D.font).setFontSize(10)
        .setFontColor(D.muted).setBackground(D.card)
        .setHorizontalAlignment("right").setVerticalAlignment("middle");

      d.setRowHeight(r, 30);
    }

    d.getRange(listTop, 1, catRows.length, 5).setBorder(
      null, null, null, null, null, true, D.line, SpreadsheetApp.BorderStyle.SOLID);

    SpreadsheetApp.flush();
    d.insertChart(darkDonut(d,
      d.getRange(1, DASH_CAT, catRows.length + 1, 2),
      "Spending breakdown", listTop, 7));
  }

  // ---- Income vs expenses by month ----
  const recent = rows.slice(-12);
  d.getRange(1, DASH_TRD, 300, 4).clearContent();

  const trendTop = Math.max(listTop + catRows.length + 2, 25);
  d.getRange(trendTop - 1, 1).setValue("INCOME VS EXPENSES")
    .setFontFamily(D.font).setFontSize(10).setFontWeight("bold")
    .setFontColor(D.muted);
  d.setRowHeight(trendTop - 1, 30);

  if (recent.length) {
    d.getRange(1, DASH_TRD, 1, 4).setValues([["Month", "Income", "Expenses", "Net"]]);
    recent.forEach((m, i) => {
      d.getRange(i + 2, DASH_TRD, 1, 4)
        .setValues([[shortMonth(m.name), m.inc, m.exp, m.net]]);
    });
    SpreadsheetApp.flush();

    d.insertChart(darkBar(d, d.getRange(1, DASH_TRD, recent.length + 1, 3),
      "Income vs expenses", [D.green, D.blue], trendTop, 1, true));

    d.insertChart(darkBarTwo(d,
      d.getRange(1, DASH_TRD, recent.length + 1, 1),
      d.getRange(1, DASH_TRD + 3, recent.length + 1, 1),
      "Income left", [D.purple], trendTop, 7));
  }

  // ---- Calendar heatmap ----
  const calTop = trendTop + 17;
  d.getRange(calTop - 1, 1).setValue("CALENDAR")
    .setFontFamily(D.font).setFontSize(10).setFontWeight("bold")
    .setFontColor(D.muted);
  d.setRowHeight(calTop - 1, 30);

  // Hidden helpers: first day, weekday offset, days in month
  const H = DASH_CAL;
  d.getRange(1, H).setFormula('=IFERROR(DATEVALUE("1 "&$B$7),"")');
  d.getRange(2, H).setFormula('=IFERROR(WEEKDAY(' + colLetter(H) + '1,2),1)');
  d.getRange(3, H).setFormula('=IFERROR(DAY(EOMONTH(' + colLetter(H) + '1,0)),30)');
  const offRef = "$" + colLetter(H) + "$2";
  const dimRef = "$" + colLetter(H) + "$3";

  const dayNames = ["M", "T", "W", "T", "F", "S", "S"];
  d.getRange(calTop, 1, 1, 7).setValues([dayNames])
    .setFontFamily(D.font).setFontSize(9).setFontWeight("bold")
    .setFontColor(D.muted).setBackground(D.bg)
    .setHorizontalAlignment("center");
  d.setRowHeight(calTop, 24);

  const amountCells = [];
  const dayCells = [];

  for (let w = 0; w < 6; w++) {
    const dayRow = calTop + 1 + w * 2;
    const amtRow = dayRow + 1;

    for (let c = 1; c <= 7; c++) {
      const expr = "(" + (w * 7 + c) + "-" + offRef + "+1)";

      d.getRange(dayRow, c).setFormula(
        "=IF(OR(" + expr + "<1," + expr + ">" + dimRef + '),"",' + expr + ")")
        .setFontFamily(D.font).setFontSize(9)
        .setFontColor(D.muted).setBackground(D.card)
        .setHorizontalAlignment("center").setVerticalAlignment("bottom");

      const dayCell = colLetter(c) + dayRow;
      d.getRange(amtRow, c).setFormula(
        '=IF(' + dayCell + '="","",SUMIFS(' + Q + 'D:D,' + Q + 'A:A,$B$7,' +
        Q + 'E:E,' + dayCell + ',' + Q + 'B:B,"Expense"))')
        .setNumberFormat("0;;")   // blank when zero
        .setFontFamily(D.font).setFontSize(11).setFontWeight("bold")
        .setFontColor(D.text).setBackground(D.card)
        .setHorizontalAlignment("center").setVerticalAlignment("top");
    }

    d.setRowHeight(dayRow, 20);
    d.setRowHeight(amtRow, 26);
    amountCells.push(d.getRange(amtRow, 1, 1, 7));
    dayCells.push(d.getRange(dayRow, 1, 1, 7));
  }

  // Tint the day-number row too when that day had spend, so each day reads as
  // one block rather than two disconnected rows.
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=A" + (calTop + 2) + ">0")
      .setBackground(D.heatLow)
      .setRanges(dayCells).build()
  );

  // Heat: darker card when nothing spent, brighter as spend rises
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setGradientMinpointWithValue(D.heatLow, SpreadsheetApp.InterpolationType.NUMBER, "0")
      .setGradientMaxpointWithValue(D.heatHigh, SpreadsheetApp.InterpolationType.PERCENTILE, "90")
      .setRanges(amountCells).build()
  );

  d.setConditionalFormatRules(rules);

  // ---- Widths ----
  d.setColumnWidth(1, 130);
  d.setColumnWidth(2, 120);
  d.setColumnWidth(3, 60);
  d.setColumnWidth(4, 110);
  d.setColumnWidth(5, 70);
  d.setColumnWidth(6, 60);   // gutter: left block ends at 550, charts start there
  for (let c = 7; c <= 14; c++) d.setColumnWidth(c, 100);

  d.setFrozenRows(1);
}

// Number of expense entries in a month sheet for one category.
function countEntries(sheet, category) {
  if (!sheet) return 0;
  const last = lastDataRow(sheet);
  if (last < 2) return 0;
  let n = 0;
  sheet.getRange(2, 2, last - 1, 2).getValues().forEach(v => {
    if (String(v[0]).trim().toLowerCase() === "expense" &&
        String(v[1]).trim() === category) n++;
  });
  return n;
}

// ---------- Dark chart builders ----------

function darkDonut(sheet, range, title, posRow, posCol) {
  return sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(range)
    .setPosition(posRow, posCol, 0, 0)
    .setOption("title", title)
    .setOption("fontName", D.font)
    .setOption("titleTextStyle", { fontSize: 13, bold: true, color: D.text, fontName: D.font })
    .setOption("colors", DARK_PALETTE)
    .setOption("pieHole", 0.68)
    .setOption("pieSliceBorderColor", D.card)
    .setOption("pieSliceText", "none")
    .setOption("legend", { position: "right",
      textStyle: { fontSize: 11, color: D.muted, fontName: D.font } })
    .setOption("backgroundColor", D.card)
    .setOption("width", 540).setOption("height", 330)
    .build();
}

function darkBar(sheet, range, title, colors, posRow, posCol, legend) {
  return sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setPosition(posRow, posCol, 0, 0)
    .setOption("title", title)
    .setOption("fontName", D.font)
    .setOption("titleTextStyle", { fontSize: 13, bold: true, color: D.text, fontName: D.font })
    .setOption("colors", colors)
    .setOption("legend", legend
      ? { position: "top", textStyle: { fontSize: 11, color: D.muted, fontName: D.font } }
      : { position: "none" })
    .setOption("bar", { groupWidth: "45%" })
    .setOption("hAxis", { textStyle: { fontSize: 10, color: D.muted, fontName: D.font },
      gridlines: { color: "none" } })
    .setOption("vAxis", { textStyle: { fontSize: 10, color: D.muted, fontName: D.font },
      gridlines: { color: D.line }, baselineColor: D.line })
    .setOption("backgroundColor", D.card)
    .setOption("width", 540).setOption("height", 310)
    .build();
}

function darkBarTwo(sheet, labels, values, title, colors, posRow, posCol) {
  return sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(labels).addRange(values)
    .setPosition(posRow, posCol, 0, 0)
    .setOption("title", title)
    .setOption("fontName", D.font)
    .setOption("titleTextStyle", { fontSize: 13, bold: true, color: D.text, fontName: D.font })
    .setOption("colors", colors)
    .setOption("legend", { position: "none" })
    .setOption("bar", { groupWidth: "45%" })
    .setOption("hAxis", { textStyle: { fontSize: 10, color: D.muted, fontName: D.font },
      gridlines: { color: "none" } })
    .setOption("vAxis", { textStyle: { fontSize: 10, color: D.muted, fontName: D.font },
      gridlines: { color: D.line }, baselineColor: D.line })
    .setOption("backgroundColor", D.card)
    .setOption("width", 540).setOption("height", 310)
    .build();
}

// ---------- Small builders ----------

function shortMonth(name) {
  const p = String(name).trim().split(/\s+/);
  if (p.length < 2) return name;
  return p[0].substring(0, 3) + " " + p[p.length - 1].slice(-2);
}

function monthKey(name) {
  const p = String(name).trim().split(/\s+/);
  const names = ["january","february","march","april","may","june",
                 "july","august","september","october","november","december"];
  const mi = names.indexOf(String(p[0]).toLowerCase());
  const y = parseInt(p[p.length - 1], 10);
  if (isNaN(y) || mi < 0) return 0;
  return y * 12 + mi;
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}
