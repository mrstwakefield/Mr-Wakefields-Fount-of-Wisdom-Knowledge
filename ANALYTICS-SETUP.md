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
2. In the sheet: **Extensions → Apps Script**. Delete what's there and paste in the whole of `analytics-script.gs` (from this repository). Click **Save**.
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

The code is in **`analytics-script.gs`** in this repository. Open it, select everything (Ctrl+A), copy it, and paste it into the Apps Script editor in place of what's there.
