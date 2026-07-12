# 2 Out Storm 10U — Fully Online Storm Website

This is the complete online version for **2outstorm2035.com**.

## What runs online

- **GitHub Pages:** public 2 Out Storm 10U website
- **Google Sheets:** roster, schedule, announcements, videos, reports, form links, parent codes, and availability
- **Google Apps Script:** secure Coach Edit Board, coded parent availability, public website data feed, and media uploads
- **Google Drive:** uploaded player photos, gallery photos, thumbnails, and short Storm Channel video clips

You do not have to edit GitHub every time the roster, schedule, video library, or announcements change. After the one-time setup, use the online Coach Edit Board.

## Included

- Full storm-themed public website
- Storm Channel with YouTube, Google Drive, and uploaded clip support
- Meet the Storm roster and athlete profiles
- Storm Tracker schedule, results, and next-impact countdown
- Storm Development video library
- Storm Reports, shoutouts, Picture of the Week, and gallery
- The Shelter parent portal
- Private 6-digit player-code availability system
- Password-protected online Coach Edit Board
- Google Drive photo, thumbnail, and short-video uploads
- Existing Player Interest form already linked

## 1. Create the Google Sheet

1. Upload the separately included `2_Out_Storm_10U_Online_Website_Data.xlsx` to Google Drive.
2. Open it with Google Sheets.
3. Use the `START HERE` sheet as the checklist.

## 2. Install Apps Script

1. In the Sheet, choose **Extensions > Apps Script**.
2. Replace `Code.gs` with the included `apps-script/Code.gs`.
3. Add these HTML files with the exact names:
   - `Admin`
   - `Availability`
   - `StormStyles`
4. Copy the matching included HTML into each file.
5. In Apps Script Project Settings, enable **Show appsscript.json manifest file in editor**.
6. Replace the manifest with `apps-script/appsscript.json`.
7. Save and run `setupStormWorkbook` once.
8. Approve the requested Sheet and Drive permissions.
9. Refresh the Google Sheet.
10. Use **Storm Website > Set Coach Password**.
11. Use **Storm Website > Generate Missing Parent Codes** after player names are entered.

## 3. Deploy the online backend

1. Apps Script: **Deploy > New deployment**.
2. Type: **Web app**.
3. Execute as: **Me**.
4. Who has access: **Anyone**.
5. Deploy and copy the URL ending in `/exec`.

The same web-app URL powers:

- Public data
- Coach Edit Board
- Parent availability
- Photo and video uploads

## 4. Connect the public website

Open `assets/js/config.js` and paste the deployed `/exec` URL into:

```js
apiUrl: "PASTE_YOUR_APPS_SCRIPT_EXEC_URL_HERE"
```

Do this once before uploading the repository to GitHub.

## 5. Publish on GitHub Pages

1. Create a new public GitHub repository.
2. Upload the **contents** of this folder.
3. Open **Settings > Pages**.
4. Deploy from `main` and `/root`.
5. Test the temporary GitHub Pages address.
6. After testing, connect `2outstorm2035.com` in Pages settings.

## Coach Edit Board

Open the public site and choose **Coach Edit Board** in the footer or The Shelter. Sign in with the password set from the spreadsheet menu.

The editor controls:

- Players and athlete profiles
- Player photos and profile backgrounds
- Practices, games, tournaments, locations, scores, and results
- Storm Channel videos
- YouTube links or IDs
- Google Drive video links
- Short video uploads from your computer or phone
- Video thumbnails
- Announcements and Storm Warnings
- Picture of the Week
- Gallery photos
- Shoutouts and Storm Reports
- Form links and contact information
- Parent codes
- Availability results

## Video choices

### YouTube
Paste the YouTube URL or ID into `YouTubeID`.

### Upload a short clip
Save the video row first, then choose `VideoURL` in the upload box and upload the clip. It is stored in the Google Drive folder **2 Out Storm Website Videos**.

### Existing Google Drive video
Share the Drive video so **Anyone with the link** can view it, then paste the share link into `VideoURL`.

For clips over 45 MB, upload directly to Google Drive and paste the link rather than uploading through the editor.

## Parent availability

1. Enter the real roster in the Coach Edit Board.
2. Generate missing parent codes.
3. Give each family only its player’s 6-digit code.
4. Parents choose **The Shelter > Update Availability**.
5. Their updates are written to the private `Availability` sheet.
6. You can review them inside the Coach Edit Board.

`Parent Codes`, `Availability`, and `Admin Log` are never returned by the public website feed.

## Existing form link

The recovered Player Interest form is already connected:

`https://forms.gle/4PgA83dnnVBXMo2h8`

Paste the exact URLs for Private Parent Message, Accepted Player Profile, Submit Team Photo, and GameChanger into Website Settings when available.
