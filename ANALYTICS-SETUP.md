# Anonymous usage stats → Google Sheet

The revision site sends **anonymous totals** to a Google Sheet you own. It works the same way as the GCSE site.

| Sent | Never sent |
|---|---|
| Games started, and right/wrong per game | Names, emails or the report-box name |
| Right/wrong per card (card number, topic, question, AO) | Device codes or any ID |
| The date and year group (12 / 13) | Answers or anything typed |
| "First time this device has sent anything" / "first time today" (used to count students) | Anything from teacher mode or demo mode |

Students can switch it off in **Saves → 📊 Anonymous usage stats**.

## The tabs

- **Summary**: headline numbers, hardest and easiest 10 cards (at least 10 answers each), topics weakest first, AOs weakest first. Rebuilt every time data arrives.
- **Daily**: one row per date and year group.
- **Games**: how often each game is started.
- **Modes**: right/wrong per game.
- **Cards**: right/wrong per card.

## Updating the script (keeps the same web address)

1. Open the sheet → **Extensions → Apps Script**.
2. Select everything in the editor, delete it, and paste in the whole of `analytics-script.gs`. Click **Save**.
3. Choose **setup** in the function menu and click **Run**.
4. **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy.**

## Clearing the data

Run **wipeData** from the Apps Script editor. It keeps the tabs and headings.

## First-time setup

1. Create a Google Sheet, then **Extensions → Apps Script**. Paste in `analytics-script.gs` and save.
2. Run **setup** and approve the permissions.
3. **Deploy → New deployment → Web app**. Execute as **Me**, Who has access **Anyone**. Deploy and copy the URL (ends in `/exec`).
4. Put the URL in `ANALYTICS_URL` in `index.html`.

"Anyone" only means the site can send numbers without students logging in. Nobody can read the sheet unless you share it.
