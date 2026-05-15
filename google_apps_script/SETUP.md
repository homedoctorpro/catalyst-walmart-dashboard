# Google Sheet Sync — Setup Guide

Connects the Retailers tab on the Walmart Sales Dashboard to a Google Sheet so edits sync both ways:

- Edits you make in the dashboard get written to the Sheet immediately.
- Edits you make in the Sheet appear in the dashboard when you load the tab or refocus the browser window.

## One-time setup (~5 minutes)

### 1. Create a new Google Sheet
Any name works (e.g. "Catalyst Retailers Sync"). Leave it empty — the script will populate it.

### 2. Open the Apps Script editor
In the Sheet, go to **Extensions → Apps Script**. A new tab opens with a code editor.

### 3. Paste the script
- In the Apps Script editor, delete the default `Code.gs` content.
- Open `retailers_sync.gs` (in this same folder), copy the **entire** file, paste it into the editor.
- Click the **💾 Save** icon (top toolbar). Name the project anything — e.g. "Retailers Sync".

### 4. Deploy as a Web App
- Click **Deploy → New deployment** (top-right blue button).
- Click the **⚙ gear** next to *Select type* → choose **Web app**.
- Fill in:
  - **Description:** anything (optional)
  - **Execute as:** *Me (your-email@…)*
  - **Who has access:** *Anyone*
- Click **Deploy**.

### 5. Authorize
- Google will prompt you to authorize the script. Pick your account.
- You may see a "Google hasn't verified this app" warning — this is normal for a private script you wrote.
- Click **Advanced → Go to {your project name} (unsafe) → Allow**.
- The script needs permission to read/write your Sheet (`SpreadsheetApp` scope).

### 6. Copy the Web App URL
The deployment dialog shows a URL like:
```
https://script.google.com/macros/s/AKfycbz……/exec
```
Copy it. This is the endpoint your dashboard will talk to.

### 7. Connect the dashboard
- Open the dashboard: https://homedoctorpro.github.io/catalyst-walmart-dashboard/
- Log in, click the **🤝 Retailers** tab.
- At the top of the tab, expand the **⚙ Sheet sync settings** bar.
- Paste the URL into the input, click **Connect**.
- The status dot should turn green and say "Synced".

### 8. Seed the Sheet with your current edits
- In the sync bar, click **⬆ Push all local edits to Sheet**.
- This copies your in-browser rep firm / status / next steps / U/S/W overrides into the Sheet as the starting state.
- From now on the Sheet is the source of truth.

You're done.

---

## Day-to-day use

- **Edit in the dashboard:** changes save to the Sheet immediately (status dot flashes yellow → green).
- **Edit in the Sheet:** changes appear in the dashboard when you reload the page or switch back to the dashboard tab.
- **Pull on demand:** the sync bar has a **↻ Pull from Sheet now** button if you want to force-refresh without reloading.
- **Multiple people:** anyone with the dashboard URL + password can edit, and changes route through your Sheet. Co-edits resolve last-write-wins (no conflict warnings).

## Editing in the Sheet

The Sheet has these columns:

| Column | Edit? | Notes |
|---|---|---|
| `retailer_id` | ❌ | Stable ID — don't change. |
| `retailer_name` | ❌ | Mirrors the dashboard name (gets overwritten on push). |
| `channel` | ❌ | Mirrors the dashboard channel (gets overwritten on push). |
| `rep_firm` | ✅ | Free text. The dashboard dropdown shows `Unassigned / PSE / Brian Schlager`; you can put any string here and the dashboard will display it, but only those three appear in the dropdown choices. Blank = Unassigned. |
| `status` | ✅ | One of: `in`, `pitched`, `target`, `declined` (or blank). |
| `next_steps` | ✅ | Free text. |
| `usw_override` | ✅ | Number — overrides the channel default U/S/W. Leave blank to use the channel default. |
| `updated_at` | ❌ | Auto-stamped on every write. |

Rows you delete in the Sheet effectively clear that retailer's overrides (the dashboard re-creates them as defaults).

Rows the dashboard hasn't touched won't appear in the Sheet yet — they get added the first time you edit them in the dashboard or run **Push all local edits**.

---

## Updating the script later

If you change `retailers_sync.gs`:

1. Open the Sheet → Extensions → Apps Script.
2. Edit the code, hit Save.
3. **Deploy → Manage deployments** → pencil-edit your existing deployment → **Version: New version** → **Deploy**.

This keeps the same URL working, no need to re-paste anywhere.

---

## Disconnecting

- **In the dashboard:** sync settings bar → **Disconnect**. The dashboard reverts to localStorage-only mode.
- **In Apps Script:** Deploy → Manage deployments → Archive your deployment. The URL becomes dead.

---

## Caveats

- **No realtime push.** Sheet edits arrive when you reload or refocus the tab. If two people touch the same cell within that window, last write wins silently.
- **~1–2 s per save.** Each edit is a round trip to Google. The status dot shows saving / saved / failed.
- **The script URL is in the dashboard's HTML source.** Anyone with the dashboard URL + password can read/write the Sheet. If that's not acceptable, redeploy the script with *Who has access: Anyone with Google account* and we'll wire a Google sign-in flow into the dashboard.
- **Quota.** Apps Script gives consumer Google accounts ~20,000 calls/day and 90 minutes of script runtime/day. This workflow uses a tiny fraction of that.
- **Offline edits.** If your laptop goes offline while editing, the change is saved in localStorage but the push fails (red status dot). On reconnect, click **⬆ Push all local edits to Sheet** to re-send.
