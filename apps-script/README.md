# Google Apps Script backend

Stores the schedule in a Google Sheet so the whole group shares one calendar. No server to run and no hosting bill — Apps Script runs the endpoint for you.

The script is standalone (not bound to the Sheet) and reaches the spreadsheet through `SpreadsheetApp.openById(SHEET_ID)`. `SHEET_ID` at the top of `Code.gs` points at the spreadsheet currently in use.

## Setup (once, ~5 minutes)

1. Create a new Google Sheet, or reuse one, and copy its id out of the URL into `SHEET_ID`.
2. Go to <https://script.google.com> and create a new project.
3. Paste the contents of `Code.gs` over the placeholder code, and save.
4. **Deploy > New deployment > Web app**:
   - *Execute as*: **Me**
   - *Who has access*: **Anyone**
5. Authorize when prompted, then copy the web app URL. It ends in `/exec`.
6. Open the calendar app, paste that URL into **Đồng bộ qua Google Sheet**, and press **Kết nối**.

After connecting, every change is written to the Sheet. Other people paste the same URL (or open a shared link, which already carries it) and see the same schedule. The app re-checks the Sheet every 30 seconds and whenever the tab regains focus.

## What lands in the Sheet

- `_state` — cell A1 holds the raw JSON state. This is the source of truth; don't edit it by hand.
- `Lịch` — a readable table (week, date range, person, notes) regenerated on every save. Safe to read, print, or chart; edits here are overwritten on the next save.

## Conflicts

Every save carries an `updatedAt` timestamp. The script rejects a write whose timestamp is older than what the Sheet already holds, so a stale tab cannot clobber newer data. Two people editing the same minute still resolve to last-write-wins.

## Access

*Who has access: Anyone* means anyone holding the `/exec` URL can read and write the schedule. The URL is unguessable, but treat it like a password: don't post it publicly. Note that a share link copied from the app contains this URL.

## Current deployment

A project is already deployed against the spreadsheet named in `SHEET_ID`. The `/exec` URL is not stored in this repo — it lives in the app's URL hash and `localStorage` once you connect, and in **Deploy > Manage deployments** in the Apps Script project.

## Updating the script

After editing `Code.gs`, redeploy with **Deploy > Manage deployments > edit > Version: New version**. Keeping the same deployment preserves the `/exec` URL, so nobody has to reconnect.
