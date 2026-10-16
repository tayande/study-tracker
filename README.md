# Study Tracker

A personal study tracker that runs free on Google. You open one link on your phone or laptop and get two pages:

- **Today:** your time slots for the day with that day's lessons. Tick each block **Done**, log the minutes you spent, and add notes. A week panel shows your progress and hours by activity.
- **Timetable:** your weekday and weekend schedule, the week's lesson plan with a **Log** button for each day, and your resource links.

Everything is saved in a Google Sheet in your own Google Drive. Once deployed, the link opens with no login.

## How it works

| File | What it does |
| --- | --- |
| `Code.gs` | Google Apps Script backend. Serves the page, reads your schedule and plan from the sheet, and saves your log. |
| `Index.html` | The tracker page (HTML, CSS and JavaScript). |

The Google Sheet holds your data in tabs: **Settings** (name, start date), **Schedule** (time slots), **Plan** (lessons by day number), **Resources** (links) and **Log** (what you did). To personalise the tracker, edit those tabs; no code changes are needed.

## Set it up (about 10 minutes, once)

1. **Sign in to Google** with the account that should own the tracker.
2. **Create a sheet:** go to [sheets.new](https://sheets.new) and name it `Study Tracker`.
3. **Open the script editor:** in the sheet, click **Extensions → Apps Script**. Rename the project (top left) to `Study Tracker`.
4. **Add the backend:** click `Code.gs`, delete what's there, paste the contents of `Code.gs` from this repo, and click **Save** (the disk icon).
5. **Add the page:** next to **Files**, click **+ → HTML**, name it exactly `Index` (Google adds `.html`). Delete what's there, paste the contents of `Index.html`, and click **Save**.
6. **Create the tabs:** at the top of the editor, choose `setup` in the function dropdown and click **Run**.
   - Click **Review permissions** and choose your account.
   - If you see "Google hasn't verified this app", click **Advanced → Go to Study Tracker (unsafe)**. It's your own script, so this is expected. Click **Allow**.
   - Go back to the sheet: you'll see the tabs **Start Here, Settings, Schedule, Plan, Resources, Log**.
7. **Personalise:** in the **Settings** tab, type your name and check the start date (`YYYY-MM-DD`). Adjust **Schedule** if your times differ.
8. **Deploy:** in the editor, click **Deploy → New deployment**. Click the gear next to "Select type" and choose **Web app**.
   - Description: `Study Tracker`
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**, then copy the **Web app URL**. That is your tracker link.
9. **Save the link:** open it, bookmark it, and on your phone use **Add to Home screen** so it opens like an app. Reload the sheet and a **Study Tracker** menu appears; **Show my tracker link** finds the link again any time.

## Everyday use

- Open your link. It lands on today. After each block, tick **Done** (it fills in the planned minutes) and change the minutes if you spent more or less.
- Slots after midnight belong to the day they started, so at 1 am the tracker still shows the previous day.
- Add next week's lessons as new rows in the **Plan** tab (Day 8, 9, ...). The Key column must match a Key in the **Schedule** tab, such as `aiefs`, `python`, `youtube`, `learning` or `research`.
- Changes to Settings, Schedule, Plan and Resources show up when you reload the page. No redeploy needed.

## Changing the code

Edit `Code.gs` or `Index.html` in the Apps Script editor and save. Then click **Deploy → Manage deployments**, click the pencil icon, set **Version** to **New version**, and click **Deploy**. Your link stays the same.

Keep this repo in step by pasting the same changes into the files here.

## Sharing with someone

Share your **Google Sheet** with them as **Viewer**, not your tracker link (anyone with your link can log in your tracker). They:

1. Open the sheet and choose **File → Make a copy**. The script is copied along with it.
2. In their copy, edit **Settings** and **Schedule**.
3. Open **Extensions → Apps Script** and follow steps 6 to 9 above.

Their tracker and data are completely separate from yours.

## Limits

- It uses Google's free Apps Script quotas, which are far above what a personal tracker needs.
- Pages served by Apps Script show a small Google banner at the top. That's normal.
- If the tracker can't save (for example, you're offline), it shows a red message and retries on your next edit.
