# Lịch trực nhật

Single-file web app for a weekly chore rotation. One person per week, in order, with skip and swap support.

## Features

- Month calendar view: one row per week, with the assigned person in the left column.
- Enter the member list and the start date; the rotation is computed from there.
- **Bỏ tuần** (skip): that person gives up their turn and the rest of the rotation shifts one week earlier.
- **Đổi** (swap): click "Đổi" on two weeks to exchange the people assigned to them.
- The whole schedule is encoded in the URL hash, so sharing the link shares the schedule. State is also cached in `localStorage`.
- Optional shared storage with no server of your own: connect a Google Apps Script web app backed by a Google Sheet (see `apps-script/README.md`). Changes are then saved to the Sheet and picked up by everyone else's tab.
- Export/import the schedule as JSON.

## Run locally

Open `index.html` in a browser. No build step, no server, no dependencies.

Some browser extensions cannot touch `file://` pages; if you need the app on a real origin while developing, run `python3 -m http.server 8777` in this folder and open <http://localhost:8777/index.html>.

## Deploy

Any static host works, since the app is one file:

- **Netlify Drop** — drag the folder onto https://app.netlify.com/drop.
- **GitHub Pages** — push the folder to a repo, then enable Pages on the branch root.
- **Cloudflare Pages / Vercel** — point at the folder, no build command.

After deploying, set up the member list once and share the resulting URL (with the `#...` part) with the group.
