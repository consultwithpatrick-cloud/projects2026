# The Millionaire Engine

A self-contained personal-finance and money-utilization tracker, built as an installable Progressive Web App (PWA) backed by a Google Sheet. No server and no database fees: the phone talks to a Google Apps Script webhook, which reads and writes a Google Sheet.

Built by Patrick Panuncillon.

## Architecture

```
  PWA (this HTML)  ──HTTPS POST { secret, action, data }──►  Apps Script Web App (doPost)
   PIN-encrypted config on device                              validates the secret,
   (localStorage, AES-256-GCM)                                 reads/writes the Google Sheet
```

- **Frontend:** one HTML file plus Chart.js, installable to the home screen. Nothing client-specific is baked in; the webhook URL and secret are entered once at first run.
- **Backend:** `MillionaireEngine_AppsScript.gs`, a secret-gated webhook with actions for logging transactions, updating payables, and returning monthly summaries and category breakdowns.
- **Security:** the webhook URL and secret are encrypted on the device with a PIN-derived key (PBKDF2 → AES-256-GCM). The app auto-locks on idle and wipes after repeated failed PIN attempts.

## Files

| File | What it is |
|---|---|
| `millionaire-engine.html` | The PWA. The whole app in one file. |
| `manifest.json`, `sw.js` | PWA manifest and service worker (offline shell). |
| `icon-*.png`, `apple-touch-icon.png`, `*.svg` | App icons and marks. |
| `MillionaireEngine_AppsScript.gs` | The backend webhook (deploy to Google Apps Script). |
| `engine-demo/` | A standalone demo build with sample data, for a quick look without any setup. |

## Setup

1. Create a Google Sheet for your data.
2. Open the Apps Script editor, paste `MillionaireEngine_AppsScript.gs`, and set `SPREADSHEET_ID` to your sheet's ID.
3. Set a strong, unique `WEBHOOK_SECRET`. Prefer a Script Property named `WEBHOOK_SECRET` over the in-file constant.
4. Deploy as a Web App (execute as you, access to anyone), and copy the `/exec` URL.
5. Open the app, set a PIN, and paste in your `/exec` URL and secret. Add it to your home screen.

## Notes

- Placeholders in `MillionaireEngine_AppsScript.gs` (`SPREADSHEET_ID`, `WEBHOOK_SECRET`) must be filled in with your own values. Do not commit a real secret.
- Every instance is single-user: one person, one sheet, one webhook, one copy of the app.
