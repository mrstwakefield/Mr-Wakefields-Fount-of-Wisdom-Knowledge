// Fount usage stats: paste ALL of this into Extensions > Apps Script, replacing what's there.
var RAW = "Raw data";
var SUM = "Summary";
var FIXED = ["Date", "Device", "Year", "Level", "Device type", "Questions", "Correct", "Accuracy %", "Last update"];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.v !== 1) return out_("bad");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d.day)) return out_("bad");
    if (!/^[a-z0-9]{4,20}$/.test(d.dev)) return out_("bad");

    var sh = sheet_(RAW);
    var width = Math.max(sh.getLastColumn(), 1);
    var head = sh.getRange(1, 1, 1, width).getValues()[0];
    if (!head[0]) {
      sh.getRange(1, 1, 1, FIXED.length).setValues([FIXED]).setFontWeight("bold");
      sh.setFrozenRows(1);
      sh.getRange("A:B").setNumberFormat("@");
      head = FIXED.slice();
    }

    var q = num_(d.q);
    var r = num_(d.r);
    var cells = {};
    cells["Date"] = d.day;
    cells["Device"] = d.dev;
    cells["Year"] = String(d.yr || "").slice(0, 8);
    cells["Level"] = num_(d.lvl);
    cells["Device type"] = String(d.plat || "").slice(0, 20);
    cells["Questions"] = q;
    cells["Correct"] = r;
    cells["Accuracy %"] = q ? Math.round(r / q * 100) : "";
    cells["Last update"] = new Date();
    addMap_(cells, "Game: ", d.games);
    addMap_(cells, "Topic Qs: ", d.tq);
    addMap_(cells, "Topic right: ", d.tr);
    addMap_(cells, "AO Qs: ", d.aq);
    addMap_(cells, "AO right: ", d.ar);
    cells["Cards"] = JSON.stringify(cleanCards_(d.c));

    var keys = Object.keys(cells);
    for (var k = 0; k < keys.length; k++) {
      if (head.indexOf(keys[k]) < 0) {
        head.push(keys[k]);
        sh.getRange(1, head.length).setValue(keys[k]).setFontWeight("bold");
      }
    }
    var row = [];
    for (var h = 0; h < head.length; h++) {
      row.push(cells.hasOwnProperty(head[h]) ? cells[head[h]] : "");
    }

    // find this device's row for today, searching the most recent 3000 rows
    var last = sh.getLastRow();
    var start = Math.max(2, last - 3000);
    var n = last - start + 1;
    var at = -1;
    if (n > 0) {
      var v = sh.getRange(start, 1, n, 2).getDisplayValues();
      for (var i = v.length - 1; i >= 0; i--) {
        if (v[i][0] === d.day && v[i][1] === d.dev) {
          at = start + i;
          break;
        }
      }
    }
    if (at > 0) {
      sh.getRange(at, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return out_("ok");
  } catch (err) {
    return out_("error");
  } finally {
    lock.releaseLock();
  }
}

function cleanCards_(c) {
  var out = {};
  if (!c) return out;
  Object.keys(c).slice(0, 2000).forEach(function (id) {
    if (/^\d{1,6}$/.test(id) && c[id] && c[id].length === 2) out[id] = [num_(c[id][0]), num_(c[id][1])];
  });
  return out;
}

function addMap_(cells, prefix, map) {
  if (!map) return;
  var keys = Object.keys(map);
  for (var i = 0; i < keys.length; i++) {
    cells[prefix + clean_(keys[i])] = num_(map[keys[i]]);
  }
}

var SITE = "https://mrstwakefield.github.io/Mr-Wakefields-Fount-of-Wisdom-Knowledge/index.html";
var TOPIC_NAMES = {"RM": "Research methods", "Issues and Debates part I": "Issues and debates", "Psychopathology": "Clinical psychology",
  "Forensic Psychology": "Forensic", "Social Influence": "Social influence"};

function refreshSummary() {
  var raw = sheet_(RAW);
  var sum = sheet_(SUM);
  var data = raw.getDataRange().getValues();
  sum.clear();
  if (data.length < 2) { sum.getRange(1, 1).setValue("No data yet."); return; }
  var head = data[0];
  var rows = data.slice(1).filter(function (x) { return x[0]; });
  var cQ = head.indexOf("Questions"), cC = head.indexOf("Correct"), cDev = head.indexOf("Device");
  var cCards = head.indexOf("Cards");
  var tz = Session.getScriptTimeZone();
  var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var weekAgo = Utilities.formatDate(new Date(Date.now() - 6 * 864e5), tz, "yyyy-MM-dd");
  var real = rows.filter(function (x) { return String(x[cDev]) !== "testrow1"; });

  // ---- headline numbers
  var devs = {}, perDay = {}, q = 0, c = 0, games = {}, active = 0, week = 0;
  real.forEach(function (x) {
    var dy = String(x[0]);
    devs[x[cDev]] = 1;
    perDay[dy] = perDay[dy] || { n: 0, q: 0 };
    perDay[dy].n++; perDay[dy].q += Number(x[cQ]) || 0;
    q += Number(x[cQ]) || 0; c += Number(x[cC]) || 0;
    if (dy === todayStr) active++;
    if (dy >= weekAgo && dy <= todayStr) week++;
  });
  head.forEach(function (name, i) {
    if (String(name).indexOf("Game: ") !== 0) return;
    var n = 0; real.forEach(function (x) { n += Number(x[i]) || 0; });
    if (n) games[String(name).slice(6)] = n;
  });
  var busiest = Object.keys(perDay).sort(function (a, b) { return perDay[b].n - perDay[a].n || perDay[b].q - perDay[a].q; })[0];
  var topGame = Object.keys(games).sort(function (a, b) { return games[b] - games[a]; })[0];

  var out = [];
  out.push(["Fount usage stats (anonymous)", "", "", "", "Updated " + new Date().toLocaleString("en-GB")]);
  out.push(["", "", "", "", ""]);
  out.push(["Headline numbers", "", "", "", ""]);
  out.push(["Students (all time)", Object.keys(devs).length, "", "", ""]);
  out.push(["Active today", active, "", "", ""]);
  out.push(["Student-days in the last 7 days", week, "", "", ""]);
  out.push(["Questions answered", q, "", "", ""]);
  out.push(["% right overall", q ? Math.round(c / q * 100) + "%" : "", "", "", ""]);
  out.push(["Busiest day", busiest ? busiest + " (" + perDay[busiest].n + " students)" : "", "", "", ""]);
  out.push(["Most played game", topGame ? topGame + " (" + games[topGame] + " times)" : "", "", "", ""]);

  // ---- hardest and easiest cards (answered at least 10 times)
  var cards = {};
  if (cCards >= 0) real.forEach(function (x) {
    var m = {}; try { m = JSON.parse(x[cCards] || "{}"); } catch (e) {}
    Object.keys(m).forEach(function (id) {
      var v = m[id];
      cards[id] = cards[id] || [0, 0];
      cards[id][0] += Number(v[0]) || 0;
      cards[id][1] += Number(v[1]) || 0;
    });
  });
  var list = [];
  Object.keys(cards).forEach(function (id) {
    var n = cards[id][0];
    if (n >= 10) list.push([id, n, Math.round(cards[id][1] / n * 100)]);
  });
  var text = cardText_();
  var cardRow = function (r) { var t = text[r[0]] || ["", ""]; return [Number(r[0]), TOPIC_NAMES[t[0]] || t[0], t[1], r[1], r[2] + "%"]; };
  out.push(["", "", "", "", ""]);
  out.push(["😬 Hardest 10 cards", "Topic", "Question", "Answers", "% right"]);
  list.slice().sort(function (a, b) { return a[2] - b[2] || b[1] - a[1]; }).slice(0, 10).forEach(function (r) { out.push(cardRow(r)); });
  if (!list.length) out.push(["No card has been answered 10 times yet.", "", "", "", ""]);
  out.push(["", "", "", "", ""]);
  out.push(["😎 Easiest 10 cards", "Topic", "Question", "Answers", "% right"]);
  list.slice().sort(function (a, b) { return b[2] - a[2] || b[1] - a[1]; }).slice(0, 10).forEach(function (r) { out.push(cardRow(r)); });
  if (!list.length) out.push(["No card has been answered 10 times yet.", "", "", "", ""]);

  // ---- topics and assessment objectives, weakest first
  out.push(["", "", "", "", ""]);
  out.push(["📚 Topics, weakest first", "Answers", "% right", "", ""]);
  pairs_(head, real, "Topic Qs: ", "Topic right: ").forEach(function (r) { out.push([r[0], r[1], r[2] + "%", "", ""]); });
  out.push(["", "", "", "", ""]);
  out.push(["🎯 Assessment objectives, weakest first", "Answers", "% right", "", ""]);
  pairs_(head, real, "AO Qs: ", "AO right: ").forEach(function (r) { out.push([r[0], r[1], r[2] + "%", "", ""]); });

  // ---- extras: daily use
  out.push(["", "", "", "", ""]);
  out.push(["Daily use (last 14 days)", "Students", "Questions", "", ""]);
  Object.keys(perDay).sort().reverse().slice(0, 14).forEach(function (dy) { out.push([dy, perDay[dy].n, perDay[dy].q, "", ""]); });

  sum.getRange(1, 1, out.length, 5).setValues(out);
  sum.getRange(1, 1).setFontWeight("bold").setFontSize(14);
  for (var i = 0; i < out.length; i++) {
    if (/^(Headline|😬|😎|📚|🎯|Daily use)/.test(out[i][0])) {
      sum.getRange(i + 1, 1, 1, 5).setFontWeight("bold").setBackground("#E3E8F2");
    }
  }
  sum.setColumnWidth(1, 260); sum.setColumnWidth(3, 420);
  sum.getRange(1, 3, out.length, 1).setWrap(true);
}

// answers and % right for a family of columns, weakest first
function pairs_(head, rows, qPrefix, rPrefix) {
  var res = [];
  head.forEach(function (name, i) {
    if (String(name).indexOf(qPrefix) !== 0) return;
    var key = String(name).slice(qPrefix.length), ri = head.indexOf(rPrefix + key), q = 0, c = 0;
    rows.forEach(function (x) { q += Number(x[i]) || 0; if (ri >= 0) c += Number(x[ri]) || 0; });
    if (q > 0) res.push([key, q, Math.round(c / q * 100)]);
  });
  return res.sort(function (a, b) { return a[2] - b[2]; });
}

// question text for each card, fetched from the public website once a day
function cardText_() {
  var cache = sheet_("Card text");
  var stamp = cache.getRange(1, 1).getValue();
  if (!(stamp instanceof Date) || Date.now() - stamp.getTime() > 864e5) {
    try {
      var html = UrlFetchApp.fetch(SITE, { muteHttpExceptions: true }).getContentText();
      var m = html.match(/const DATA = (\[[\s\S]*?\]);\n/);
      if (m) {
        var rows = JSON.parse(m[1]).map(function (d) { return [d.id, d.t, String(d.q).slice(0, 300)]; });
        cache.clear();
        cache.getRange(1, 1).setValue(new Date());
        cache.getRange(2, 1, rows.length, 3).setValues(rows);
      }
    } catch (e) {}
  }
  var map = {}, n = cache.getLastRow();
  if (n > 1) cache.getRange(2, 1, n - 1, 3).getValues().forEach(function (r) { map[String(r[0])] = [r[1], r[2]]; });
  return map;
}

function totals_(rows, cQ, cC) {
  var q = 0, c = 0;
  rows.forEach(function (x) { q += Number(x[cQ]) || 0; c += Number(x[cC]) || 0; });
  return { q: q, acc: q ? Math.round(c / q * 100) : "" };
}

function setup() {
  sheet_(RAW);
  sheet_(SUM);
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("refreshSummary").timeBased().everyHours(1).create();
  refreshSummary();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Fount stats").addItem("Refresh summary", "refreshSummary").addToUi();
}

function sheet_(name) {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function num_(v) {
  v = Number(v);
  return isFinite(v) && v >= 0 && v < 10000000 ? Math.round(v) : 0;
}

function clean_(s) {
  return String(s).replace(/[^\w &'-]/g, "").slice(0, 40);
}

function out_(s) {
  return ContentService.createTextOutput(s);
}
