# Anonymous usage stats → Google Sheet

The revision site can send **anonymous** daily usage to a Google Sheet you own. One row per device per day:

| Sent | Never sent |
|---|---|
| A random device code (e.g. `k3f9x2abc`) | Names, emails or the report-box name |
| Year group (12 / 13) and level | Answers or anything typed |
| Questions answered and correct that day | Anything from teacher mode or demo mode |
| Games played and topics answered (with accuracy) | |
| Phone/tablet or computer | |

Students can switch it off in **Saves → 📊 Anonymous usage stats**.

## Setup (about 5 minutes)

1. Create a new Google Sheet, e.g. **Fount usage stats**.
2. In the sheet: **Extensions → Apps Script**. Delete what's there and paste in the code below. Click **Save**.
3. In the toolbar's function menu, choose **setup** and click **Run**. Approve the permissions (it only needs access to this spreadsheet).
4. Click **Deploy → New deployment**. Choose type **Web app**.
   - Description: `Fount stats`
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click **Deploy** and copy the **Web app URL** (it ends in `/exec`).
6. Send the URL to Claude, who will paste it into the site (`ANALYTICS_URL` in `index.html`).

"Anyone" just means the site can send rows without students logging in. Nobody can read the sheet unless you share it.

The **Summary** tab refreshes every hour (or use the **Fount stats → Refresh summary** menu).

## Apps Script code

```javascript
const RAW = "Raw data", SUM = "Summary";
const FIXED = ["Date","Device","Year","Level","Device type","Questions","Correct","Accuracy %","Last update"];

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.v !== 1 || !/^\d{4}-\d{2}-\d{2}$/.test(d.day) || !/^[a-z0-9]{4,20}$/.test(d.dev)) return out_("bad");
    const sh = sheet_(RAW);
    let head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    if (!head[0]) { sh.getRange(1, 1, 1, FIXED.length).setValues([FIXED]).setFontWeight("bold"); sh.setFrozenRows(1); sh.getRange("A:B").setNumberFormat("@"); head = FIXED.slice(); }
    const q = num_(d.q), r = num_(d.r);
    const cells = { "Date": d.day, "Device": d.dev, "Year": String(d.yr || "").slice(0, 8), "Level": num_(d.lvl),
      "Device type": String(d.plat || "").slice(0, 20), "Questions": q, "Correct": r,
      "Accuracy %": q ? Math.round(r / q * 100) : "", "Last update": new Date() };
    for (const [k, v] of Object.entries(d.games || {})) cells["Game: " + clean_(k)] = num_(v);
    for (const [k, v] of Object.entries(d.tq || {})) cells["Topic Qs: " + clean_(k)] = num_(v);
    for (const [k, v] of Object.entries(d.tr || {})) cells["Topic right: " + clean_(k)] = num_(v);
    Object.keys(cells).forEach(k => { if (head.indexOf(k) < 0) { head.push(k); sh.getRange(1, head.length).setValue(k).setFontWeight("bold"); } });
    const row = head.map(h => (h in cells ? cells[h] : ""));
    const last = sh.getLastRow(), start = Math.max(2, last - 3000), n = last - start + 1;
    let at = -1;
    if (n > 0) { const v = sh.getRange(start, 1, n, 2).getDisplayValues(); for (let i = v.length - 1; i >= 0; i--) if (v[i][0] === d.day && v[i][1] === d.dev) { at = start + i; break; } }
    if (at > 0) sh.getRange(at, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
    return out_("ok");
  } catch (err) { return out_("error"); } finally { lock.releaseLock(); }
}

function refreshSummary() {
  const raw = sheet_(RAW), data = raw.getDataRange().getValues();
  const sum = sheet_(SUM); sum.clear();
  if (data.length < 2) { sum.getRange(1, 1).setValue("No data yet."); return; }
  const head = data[0], rows = data.slice(1).filter(x => x[0]);
  const col = name => head.indexOf(name);
  const days = [...new Set(rows.map(x => String(x[0])))].sort().reverse();
  const last7 = new Set(days.slice(0, 7)), recent = rows.filter(x => last7.has(String(x[0])));
  const out = [["Fount usage stats (anonymous)", "", "", "Updated " + new Date().toLocaleString("en-GB")], [""],
    ["Daily use", "Active devices", "Questions", "Accuracy %"]];
  days.slice(0, 14).forEach(dy => { const r = rows.filter(x => String(x[0]) === dy), q = r.reduce((a, x) => a + (+x[col("Questions")] || 0), 0), c = r.reduce((a, x) => a + (+x[col("Correct")] || 0), 0); out.push([dy, r.length, q, q ? Math.round(c / q * 100) : ""]); });
  out.push([""], ["Last 7 days by year", "Devices", "Questions", "Accuracy %"]);
  ["12", "13", ""].forEach(y => { const r = recent.filter(x => String(x[col("Year")]) === y); if (!r.length) return; const devs = new Set(r.map(x => x[1])).size, q = r.reduce((a, x) => a + (+x[col("Questions")] || 0), 0), c = r.reduce((a, x) => a + (+x[col("Correct")] || 0), 0); out.push([y ? "Year " + y : "Year not set", devs, q, q ? Math.round(c / q * 100) : ""]); });
  out.push([""], ["Most played games (7 days)", "Times played"]);
  head.map((h, i) => [h, i]).filter(([h]) => String(h).startsWith("Game: ")).map(([h, i]) => [String(h).slice(6), recent.reduce((a, x) => a + (+x[i] || 0), 0)]).filter(x => x[1] > 0).sort((a, b) => b[1] - a[1]).forEach(x => out.push(x));
  out.push([""], ["Topics (7 days)", "Questions", "Accuracy %"]);
  head.map((h, i) => [h, i]).filter(([h]) => String(h).startsWith("Topic Qs: ")).map(([h, i]) => { const t = String(h).slice(10), ri = col("Topic right: " + t), q = recent.reduce((a, x) => a + (+x[i] || 0), 0), c = ri >= 0 ? recent.reduce((a, x) => a + (+x[ri] || 0), 0) : 0; return [t, q, q ? Math.round(c / q * 100) : ""]; }).filter(x => x[1] > 0).sort((a, b) => b[1] - a[1]).forEach(x => out.push(x));
  const w = Math.max(...out.map(r => r.length));
  sum.getRange(1, 1, out.length, w).setValues(out.map(r => r.concat(Array(w - r.length).fill(""))));
  sum.getRange(1, 1).setFontWeight("bold").setFontSize(14);
  out.forEach((r, i) => { if (r.length > 1 && /^(Daily use|Last 7 days|Most played|Topics)/.test(r[0])) sum.getRange(i + 1, 1, 1, w).setFontWeight("bold").setBackground("#E3E8F2"); });
  sum.autoResizeColumns(1, w);
}

function setup() {
  sheet_(RAW); sheet_(SUM);
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("refreshSummary").timeBased().everyHours(1).create();
  refreshSummary();
}
function onOpen() { SpreadsheetApp.getUi().createMenu("Fount stats").addItem("Refresh summary", "refreshSummary").addToUi(); }
function sheet_(name) { const ss = SpreadsheetApp.getActive(); return ss.getSheetByName(name) || ss.insertSheet(name); }
function num_(v) { v = Number(v); return isFinite(v) && v >= 0 && v < 1e7 ? Math.round(v) : 0; }
function clean_(s) { return String(s).replace(/[^\w &'-]/g, "").slice(0, 40); }
function out_(s) { return ContentService.createTextOutput(s); }
```
