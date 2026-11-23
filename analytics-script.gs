/**
 * Mr Wakefield's Fount (A-level): anonymous usage counts.
 *
 * Paste this whole file into your Google Sheet: Extensions → Apps Script.
 * Then: Deploy → New deployment → ⚙️ "Select type" → Web app → Execute as "Me" →
 * Who has access "Anyone" → Deploy. Copy the web-app URL (ends in /exec).
 *
 * Updating the script later: paste the new version, save, then
 * Deploy → Manage deployments → ✏️ Edit → Version: "New version" → Deploy.
 * That keeps the same web address. Then run "setup" once from the editor
 * (it also sets the Summary charts to redraw every 10 minutes).
 *
 * The revision site only ever sends totals: games started, right/wrong per
 * mode and per card, the date, the year group, and two one-off flags
 * ("first time this device has sent anything", "first time today").
 * No names, no device IDs.
 */

const TABS = {
  "Daily": ["Date", "Year", "Questions answered", "Right", "% right", "Games started", "Active students", "New students", "Minutes on site"],
  "Games": ["Game", "Times started", "Last played"],
  "Modes": ["Mode", "Right", "Wrong", "Answers", "% right"],
  "Cards": ["Card #", "Topic", "Question", "Right", "Wrong", "Answers", "% right", "AO"],
  "Events": ["Event", "Detail", "Count", "Last seen"],
  "Hours": ["Hour", "Activity"],
};
const TEXT_COLS = ["Topic", "Question", "Last played", "AO", "Event", "Detail", "Last seen", "Hour"];
const MIN_ANSWERS = 10;   // a card needs this many answers before it counts as hardest/easiest

/** Makes the tabs (and updates their headings). Runs itself whenever data arrives. */
function tabs_() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName("Summary")) ss.insertSheet("Summary", 0);
  Object.entries(TABS).forEach(([name, head]) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
    sh.setFrozenRows(1);
  });
  const old = ss.getSheetByName("Card text");   // left over from the first version
  if (old) ss.deleteSheet(old);
  const first = ss.getSheetByName("Sheet1");
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
}

/** Run once from the editor: makes the tabs, sets the Summary to redraw every 10 minutes, and draws it now. */
function setup() {
  tabs_();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (["refreshSummary", "buildSummary"].includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("buildSummary").timeBased().everyMinutes(10).create();
  buildSummary();
}

function doGet() {
  return ContentService.createTextOutput("The Fount analytics endpoint is working.");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (!p || p.v !== 1) return done_();
    tabs_();
    const date = safeDate_(p.d), year = ["y12", "y13"].includes(p.y) ? p.y : "";
    const num = v => Math.max(0, Math.min(500, Math.floor(Number(v) || 0)));   // ignore silly numbers
    const u = p.u || {}, active = u.d === 1 ? 1 : 0, fresh = u.n === 1 ? 1 : 0;

    let right = 0, answered = 0;
    upsert_("Modes", 1, Object.entries(p.modes || {}).slice(0, 40).map(([m, rw]) => {
      const r = num(rw && rw[0]), w = num(rw && rw[1]); right += r; answered += r + w;
      return { key: [clean_(m, 40)], add: [r, w] };
    }), (row, a) => { row[1] += a[0]; row[2] += a[1]; row[3] = row[1] + row[2]; row[4] = pct_(row[1], row[3]); });

    let games = 0;
    upsert_("Games", 1, Object.entries(p.plays || {}).slice(0, 40).map(([g, n]) => { const k = num(n); games += k; return { key: [clean_(g, 40)], add: [k] }; }),
      (row, a) => { row[1] += a[0]; row[2] = date; });

    const secs = Math.max(0, Math.min(14400, Math.floor(Number(p.sec) || 0)));   // time on the site in this batch (max 4 hours)
    upsert_("Daily", 2, [{ key: [date, year], add: [answered, right, games, active, fresh, secs] }],
      (row, a) => { row[2] += a[0]; row[3] += a[1]; row[4] = pct_(row[3], row[2]); row[5] += a[2]; row[6] += a[3]; row[7] += a[4]; row[8] = Math.round((Number(row[8]) + a[5] / 60) * 10) / 10; });

    upsert_("Events", 2, (Array.isArray(p.ev) ? p.ev : []).slice(0, 60).map(x => ({ key: [clean_(x && x[0], 40), clean_(x && x[1], 60)], add: [num(x && x[2])] }))
      .filter(x => x.key[0]), (row, a) => { row[2] += a[0]; row[3] = date; });

    upsert_("Hours", 1, Object.entries(p.hrs || {}).filter(([h]) => /^\d\d$/.test(h) && Number(h) < 24).map(([h, n]) => ({ key: [h + ":00–" + h + ":59"], add: [num(n)] })),
      (row, a) => { row[1] += a[0]; });

    const cards = (Array.isArray(p.cards) ? p.cards : []).slice(0, 300);
    upsert_("Cards", 1, cards.map(c => ({ key: [Math.floor(Number(c[0]) || 0)], add: [num(c[1]), num(c[2])], topic: clean_(c[3], 40), q: clean_(c[4], 120), ao: clean_(c[5], 20) }))
      .filter(c => c.key[0] > 0),
      (row, a, c) => { row[1] = c.topic || row[1]; row[2] = c.q || row[2]; row[3] += a[0]; row[4] += a[1]; row[5] = row[3] + row[4]; row[6] = pct_(row[3], row[5]); row[7] = c.ao || row[7]; });

  } catch (err) {
    // bad data is ignored
  } finally {
    lock.releaseLock();
  }
  return done_();
}

/** The Summary tab: each section is a chart with its table underneath. Redrawn every 10 minutes (and from the menu). */
function buildSummary() { tabs_(); buildSummary_(); }
/** The first version ran this every hour; keeping the name means an old timer still works. */
function refreshSummary() { buildSummary(); }

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Fount stats").addItem("Refresh summary", "buildSummary").addToUi();
}

function buildSummary_() {
  const ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName("Summary"), tz = Session.getScriptTimeZone();
  const rows = name => { const s = ss.getSheetByName(name), n = s.getLastRow(); return n > 1 ? s.getRange(2, 1, n - 1, TABS[name].length).getValues() : []; };
  const daily = rows("Daily"), games = rows("Games"), modes = rows("Modes"), cards = rows("Cards"), events = rows("Events"), hours = rows("Hours");
  const fmt = (d, f) => Utilities.formatDate(d, tz, f);
  const dayKey = v => v instanceof Date ? fmt(v, "yyyy-MM-dd") : String(v);
  const today = fmt(new Date(), "yyyy-MM-dd");
  const sum = (arr, i, f) => arr.filter(f || (() => true)).reduce((t, r) => t + (Number(r[i]) || 0), 0);
  const last = n => { const out = []; for (let i = n - 1; i >= 0; i--) { const d = new Date(); d.setDate(d.getDate() - i); out.push(d); } return out; };
  const week = last(7).map(d => fmt(d, "yyyy-MM-dd"));

  // ---- the numbers ----
  const qs = sum(daily, 2), rt = sum(daily, 3);
  const byDay = {}; daily.forEach(r => { const k = dayKey(r[0]); byDay[k] = (byDay[k] || 0) + (Number(r[6]) || 0); });
  const busiest = Object.entries(byDay).sort((a, b) => b[1] - a[1])[0];
  const topGame = games.slice().sort((a, b) => b[1] - a[1])[0];
  const mins7 = sum(daily, 8, r => week.includes(dayKey(r[0]))), sd7 = sum(daily, 6, r => week.includes(dayKey(r[0])));
  const topHour = hours.slice().sort((a, b) => b[1] - a[1])[0];
  const days30 = last(30).map(d => { const k = fmt(d, "yyyy-MM-dd"), on = daily.filter(r => dayKey(r[0]) === k);
    return ["'" + fmt(d, "d MMM"), sum(on, 6), sum(on, 2, r => r[1] === "y12"), sum(on, 2, r => r[1] === "y13"), sum(on, 2, r => r[1] !== "y12" && r[1] !== "y13"), Math.round(sum(on, 8))]; });
  const byHour = hours.slice().sort((a, b) => String(a[0]).localeCompare(String(b[0]))).map(r => ["'" + String(r[0]).slice(0, 5), Number(r[1]) || 0]);
  const group = i => { const g = {}; cards.forEach(r => { const k = r[i] || "(unknown)"; g[k] = g[k] || [0, 0]; g[k][0] += Number(r[3]) || 0; g[k][1] += Number(r[5]) || 0; });
    return Object.entries(g).filter(([, v]) => v[1] > 0).map(([k, v]) => [k, pct_(v[0], v[1]), v[1]]).sort((a, b) => a[1] - b[1]); };
  const topics = group(1), aos = group(7);
  const gameRows = games.slice().sort((a, b) => b[1] - a[1]).slice(0, 15).map(r => [r[0], Number(r[1]) || 0]);
  const modeRows = modes.filter(r => Number(r[3]) > 0).sort((a, b) => a[4] - b[4]).map(r => [r[0], Number(r[4]) || 0, Number(r[3]) || 0]);
  const feat = events.slice().sort((a, b) => b[2] - a[2]).slice(0, 15).map(r => [r[1] ? r[0] + ": " + r[1] : r[0], Number(r[2]) || 0]);
  const enough = cards.filter(r => Number(r[5]) >= MIN_ANSWERS);
  const cardRow = r => ["#" + r[0], Number(r[6]) || 0, Number(r[5]) || 0, r[1], r[2]];
  const hardest = enough.slice().sort((a, b) => a[6] - b[6] || b[5] - a[5]).slice(0, 10).map(cardRow);
  const easiest = enough.slice().sort((a, b) => b[6] - a[6] || b[5] - a[5]).slice(0, 10).map(cardRow);

  // ---- draw it ----
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.clear();
  const W = 760, H = 300, CHART_ROWS = 16;
  let r = 1;
  const write = (vals, bold) => { const g = sh.getRange(r, 1, vals.length, vals[0].length); g.setValues(vals); if (bold) g.setFontWeight("bold"); const at = r; r += vals.length; return at; };
  const title = t => { sh.getRange(r, 1).setValue(t).setFontWeight("bold").setFontSize(13); r += 1; };
  // one section: title, chart(s), then the table the chart is drawn from
  const section = (name, head, body, charts, empty) => {
    title(name);
    if (!body.length) { write([[empty]]); r += 2; return; }
    const chartAt = r; r += CHART_ROWS * charts.length;
    const top = write([head], true); write(body);
    charts.forEach((c, k) => {
      let b = sh.newChart().setChartType(c.type).setPosition(chartAt + k * CHART_ROWS, 1, 0, 0)
        .setOption("title", c.title).setOption("width", W).setOption("height", H).setOption("legend", { position: c.legend || "none" });
      c.cols.forEach(col => { b = b.addRange(sh.getRange(top, col, body.length + 1, 1)); });
      if (c.stacked) b = b.setOption("isStacked", true);
      if (c.colors) b = b.setOption("colors", c.colors);
      if (c.max) b = b.setOption("hAxis", { minValue: 0, maxValue: 100 }).setOption("vAxis", { minValue: 0 });
      b = b.setNumHeaders(1);
      sh.insertChart(b.build());
    });
    r += 2;
  };
  const C = Charts.ChartType;

  sh.getRange(r, 1, 1, 2).setValues([["📊 Fount summary", "Updated " + fmt(new Date(), "d MMM yyyy, HH:mm")]]).setFontWeight("bold");
  sh.getRange(r, 1).setFontSize(16); r += 2;
  title("Headline numbers");
  write([
    ["Students (all time)", sum(daily, 7)],
    ["Active today", sum(daily, 6, x => dayKey(x[0]) === today)],
    ["Student-days, last 7 days", sd7],
    ["Questions answered", qs],
    ["% right overall", pct_(rt, qs) + "%"],
    ["Busiest day", busiest ? busiest[0] + " (" + busiest[1] + (busiest[1] === 1 ? " student)" : " students)") : "–"],
    ["Most played game", topGame ? topGame[0] + " (" + topGame[1] + ")" : "–"],
    ["Minutes on site, last 7 days", Math.round(mins7)],
    ["Average minutes per student-day", sd7 ? Math.round(mins7 / sd7 * 10) / 10 : "–"],
    ["Busiest hour", topHour ? String(topHour[0]) : "–"],
  ]);
  r += 2;
  section("📈 Daily use, last 30 days", ["Day", "Students", "Y12 questions", "Y13 questions", "Other questions", "Minutes on site"], days30, [
    { type: C.LINE, title: "Students per day", cols: [1, 2], colors: ["#1C7F57"] },
    { type: C.COLUMN, title: "Questions answered per day", cols: [1, 3, 4, 5], stacked: true, legend: "bottom", colors: ["#4F7CFF", "#E9954B", "#B0B0B0"] },
  ], "No use yet");
  section("🕒 Activity by hour", ["Hour", "Activity"], byHour, [{ type: C.COLUMN, title: "When students revise", cols: [1, 2], colors: ["#7C5CFF"] }], "Nothing yet");
  section("📚 Topics, weakest first", ["Topic", "% right", "Answers"], topics, [{ type: C.BAR, title: "% right by topic", cols: [1, 2], colors: ["#E9954B"], max: true }], "No answers yet");
  section("🎯 Assessment objectives, weakest first", ["AO", "% right", "Answers"], aos, [{ type: C.BAR, title: "% right by AO", cols: [1, 2], colors: ["#4F7CFF"], max: true }], "No answers yet");
  section("🕹️ Most played games", ["Game", "Times started"], gameRows, [{ type: C.BAR, title: "Games started", cols: [1, 2], colors: ["#1C7F57"] }], "No games yet");
  section("✅ % right by game, weakest first", ["Game", "% right", "Answers"], modeRows, [{ type: C.BAR, title: "% right by game", cols: [1, 2], colors: ["#C0392B"], max: true }], "No answers yet");
  section("🖱️ What gets used (top 15)", ["Feature", "Count"], feat, [{ type: C.BAR, title: "Feature use", cols: [1, 2], colors: ["#F2C14E"] }], "Nothing yet");
  section("😬 Hardest cards (at least " + MIN_ANSWERS + " answers)", ["Card", "% right", "Answers", "Topic", "Question"], hardest, [{ type: C.BAR, title: "Hardest cards: % right", cols: [1, 2], colors: ["#C0392B"], max: true }], "Not enough answers yet");
  section("😎 Easiest cards (at least " + MIN_ANSWERS + " answers)", ["Card", "% right", "Answers", "Topic", "Question"], easiest, [{ type: C.BAR, title: "Easiest cards: % right", cols: [1, 2], colors: ["#1C7F57"], max: true }], "Not enough answers yet");

  sh.setColumnWidth(1, 240); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 120); sh.setColumnWidth(4, 160); sh.setColumnWidth(5, 420);
}

/** Add to existing rows (matched on the first `keyCols` columns) or append new ones, in one write. */
function upsert_(tab, keyCols, items, apply) {
  if (!items.length) return;
  const sh = SpreadsheetApp.getActive().getSheetByName(tab), width = TABS[tab].length;
  const last = sh.getLastRow(), data = last > 1 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
  data.forEach(row => { for (let j = keyCols; j < width; j++) if (row[j] === "" && !TEXT_COLS.includes(TABS[tab][j])) row[j] = 0; });   // new columns on old rows
  const index = {};
  data.forEach((row, i) => { index[keyOf_(row.slice(0, keyCols))] = i; });
  items.forEach(it => {
    let i = index[keyOf_(it.key)];
    if (i === undefined) {
      const row = new Array(width).fill(0);
      it.key.forEach((v, j) => row[j] = v);
      for (let j = keyCols; j < width; j++) if (TEXT_COLS.includes(TABS[tab][j])) row[j] = "";
      data.push(row); i = index[keyOf_(it.key)] = data.length - 1;
    }
    apply(data[i], it.add, it);
  });
  if (data.length) sh.getRange(2, 1, data.length, width).setValues(data);
}

function keyOf_(vals) {
  return vals.map(v => v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(v)).join("|");
}
function safeDate_(d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d)) ? String(d) : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"); }
function pct_(r, n) { return n ? Math.round(r / n * 100) : 0; }
/** Plain text only: stops anything that looks like a formula being stored. */
function clean_(s, max) { s = String(s == null ? "" : s).replace(/[\r\n\t]+/g, " ").slice(0, max); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function done_() { return ContentService.createTextOutput("ok"); }

/** Run from the editor to clear all the data (keeps the tabs and headings). */
function wipeData() {
  const ss = SpreadsheetApp.getActive();
  ["Daily", "Games", "Modes", "Cards", "Events", "Hours"].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  });
  const summary = ss.getSheetByName("Summary");
  if (summary) summary.clear();
}
