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

function addMap_(cells, prefix, map) {
  if (!map) return;
  var keys = Object.keys(map);
  for (var i = 0; i < keys.length; i++) {
    cells[prefix + clean_(keys[i])] = num_(map[keys[i]]);
  }
}

function refreshSummary() {
  var raw = sheet_(RAW);
  var sum = sheet_(SUM);
  sum.clear();
  var data = raw.getDataRange().getValues();
  if (data.length < 2) {
    sum.getRange(1, 1).setValue("No data yet.");
    return;
  }
  var head = data[0];
  var rows = data.slice(1).filter(function (x) { return x[0]; });
  var cQ = head.indexOf("Questions");
  var cC = head.indexOf("Correct");
  var cY = head.indexOf("Year");

  var days = [];
  rows.forEach(function (x) { var dy = String(x[0]); if (days.indexOf(dy) < 0) days.push(dy); });
  days.sort().reverse();
  var last7 = days.slice(0, 7);
  var recent = rows.filter(function (x) { return last7.indexOf(String(x[0])) >= 0; });

  var out = [];
  out.push(["Fount usage stats (anonymous)", "", "", "Updated " + new Date().toLocaleString("en-GB")]);
  out.push(["", "", "", ""]);
  out.push(["Daily use", "Active devices", "Questions", "Accuracy %"]);
  days.slice(0, 14).forEach(function (dy) {
    var r = rows.filter(function (x) { return String(x[0]) === dy; });
    var t = totals_(r, cQ, cC);
    out.push([dy, r.length, t.q, t.acc]);
  });

  out.push(["", "", "", ""]);
  out.push(["Last 7 days by year", "Devices", "Questions", "Accuracy %"]);
  ["12", "13", ""].forEach(function (y) {
    var r = recent.filter(function (x) { return String(x[cY]) === y; });
    if (!r.length) return;
    var devs = [];
    r.forEach(function (x) { if (devs.indexOf(x[1]) < 0) devs.push(x[1]); });
    var t = totals_(r, cQ, cC);
    out.push([y ? "Year " + y : "Year not set", devs.length, t.q, t.acc]);
  });

  out.push(["", "", "", ""]);
  out.push(["Most played games (7 days)", "Times played", "", ""]);
  var games = [];
  head.forEach(function (name, i) {
    if (String(name).indexOf("Game: ") !== 0) return;
    var n = 0;
    recent.forEach(function (x) { n += Number(x[i]) || 0; });
    if (n > 0) games.push([String(name).slice(6), n, "", ""]);
  });
  games.sort(function (a, b) { return b[1] - a[1]; });
  out = out.concat(games);

  out.push(["", "", "", ""]);
  out.push(["Topics (7 days)", "Questions", "Accuracy %", ""]);
  var topics = [];
  head.forEach(function (name, i) {
    if (String(name).indexOf("Topic Qs: ") !== 0) return;
    var topic = String(name).slice(10);
    var ri = head.indexOf("Topic right: " + topic);
    var q = 0, c = 0;
    recent.forEach(function (x) { q += Number(x[i]) || 0; if (ri >= 0) c += Number(x[ri]) || 0; });
    if (q > 0) topics.push([topic, q, Math.round(c / q * 100), ""]);
  });
  topics.sort(function (a, b) { return b[1] - a[1]; });
  out = out.concat(topics);

  sum.getRange(1, 1, out.length, 4).setValues(out);
  sum.getRange(1, 1).setFontWeight("bold").setFontSize(14);
  for (var i = 0; i < out.length; i++) {
    if (/^(Daily use|Last 7 days|Most played|Topics)/.test(out[i][0])) {
      sum.getRange(i + 1, 1, 1, 4).setFontWeight("bold").setBackground("#E3E8F2");
    }
  }
  sum.autoResizeColumns(1, 4);
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
