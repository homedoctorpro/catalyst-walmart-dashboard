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

### 8. Seed the Sheet with all 52 retailers
- In the sync bar, click **⬆ Push ALL retailers to Sheet (rebuild)**.
- This wipes the Sheet and writes all 52 retailers with store counts, channels, U/S/W, opportunity formulas, and any edits you've made.
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

| Col | Field | Edit? | Notes |
|---|---|---|---|
| A | `retailer_id` | ❌ | Stable ID — don't change. |
| B | `retailer_name` | ❌ | Mirrors the dashboard (overwritten on push). |
| C | `channel` | ❌ | Mirrors the dashboard (overwritten on push). |
| D | `us_stores` | ❌ | Store count (overwritten on push). |
| E | `default_usw` | ❌ | Channel default U/S/W (overwritten on push). |
| F | `usw_override` | ✅ | Number — overrides the channel default. Blank = use default. |
| G | `effective_usw` | ❌ | Formula: `=IF(F="",E,F)` — auto-updates when you edit F. |
| H | `annual_units` | ❌ | Formula: `=D*G*52` |
| I | `wholesale_opp` | ❌ | Formula: `=H*$10` |
| J | `retail_opp` | ❌ | Formula: `=H*$20` |
| K | `rep_firm` | ✅ | Free text. The dashboard dropdown shows `Unassigned / PSE / Brian Schlager / Internal / Jeff Day`; you can type any value here and it shows in the dashboard, but only those five appear in the dropdown choices. Blank = Unassigned. |
| L | `status` | ✅ | One of: `in`, `pitched`, `target`, `non-target`, `declined` (or blank). |
| M | `next_steps` | ✅ | Free text. |
| N | `updated_at` | ❌ | Auto-stamped on every write. |
| O | `next_review` | ✅ | Date (`yyyy-mm-dd`). The dashboard shows it red once the date has passed. Sheets made before columns O–P existed get the headers added in place on the next read or write; no data is cleared. |
| P | `priority` | ✅ | Whole number 1–10 (1 = highest), or blank. Anything else is ignored on pull. The dashboard sorts each channel by priority, then store count. |
| Q | `sf_account_id` | ❌ | Linked Salesforce Account (written by the Salesforce sync). |
| R | `sf_account_name` | ❌ | Account name from Salesforce. |
| S | `contacts_json` | ❌ | Up to 5 Account contacts (name, title, email, phone); shown on hover in the dashboard. |
| T | `sf_synced_at` | ❌ | Last successful Salesforce sync for the row. |

A `TOTAL` row at the bottom sums stores, units, wholesale $, and retail $.

The opportunity columns (G–J) are live Sheet formulas. When you change `usw_override` in column F, the dollar columns recompute automatically — and the dashboard reads the override on next pull.

Rows you delete in the Sheet effectively clear that retailer's overrides (the dashboard re-creates them as defaults from RT_RETAILERS). To get the retailer back as a fresh row, click **Push ALL retailers to Sheet** again.

---

## Salesforce two-way sync

Links each retailer row to a Salesforce Account (org `energex1`, Enterprise Edition) and keeps
status, rep firm, next steps, next review and priority in step both ways. It also pulls the
Account's contacts into the dashboard: the first contact shows under the retailer name, and
hovering it lists everyone with email and phone.

**What it writes in Salesforce:** only the six custom fields below. It never touches `Name`,
`Active_Door_Count__c` or any `GUMU__` field, because those come from the ERP connector. It never
creates Accounts.

### 1. Account custom fields (Setup → Object Manager → Account → Fields & Relationships → New)

| Label | API name | Type |
|---|---|---|
| Dashboard Retailer ID | `Dashboard_Retailer_ID__c` | Text(80), **Unique** + **External ID** |
| Retailer Status | `Retailer_Status__c` | Picklist, not restricted: Currently In, Pitched, Target, Non-Target, Declined |
| Retailer Rep Firm | `Retailer_Rep_Firm__c` | Picklist, not restricted: Unassigned, PSE, Brian Schlager, Internal, Jeff Day |
| Retailer Next Steps | `Retailer_Next_Steps__c` | Text Area (Long), 32,768 |
| Retailer Next Review | `Retailer_Next_Review__c` | Date |
| Retailer Priority | `Retailer_Priority__c` | Number(2, 0) |

Leave the picklists unrestricted so a new rep firm added on the dashboard doesn't bounce.
Field-level security must be **Edit** for the integration user in step 2.

### 2. External Client App (Setup → External Client App Manager → New)

- Enable OAuth; callback URL `https://login.salesforce.com/services/oauth2/success` (unused by this flow).
- Scopes: **Manage user data via APIs (api)**.
- Enable **Client Credentials Flow**; under Policies set **Run As** to a user with edit access to Accounts.
- Copy the **Consumer Key** and **Consumer Secret** (Salesforce emails a verification code to view them).

### 3. Apps Script

1. Paste the latest `retailers_sync.gs` and redeploy a new version (see *Updating the script later*).
2. **Project Settings → Script Properties**: add `SF_CLIENT_ID` (consumer key) and `SF_CLIENT_SECRET`.
3. From the editor's function dropdown, run in order:
   - `testSalesforceConnection`: confirms the login works.
   - `suggestSalesforceLinks`: writes an `SF_Link` tab listing candidate Accounts per retailer.
     Tick **confirm** on the right Account for each one; paste an Account Id for any with no match.
     Aggregate rows (`other-grocery`, `indie-via-distributors`) are skipped.
   - `applySalesforceLinks`: stamps the retailer ID onto those Accounts and runs the first sync.
   - `installSalesforceTrigger`: syncs every 15 minutes from then on.

### How conflicts resolve

A hidden `SF_Sync` tab keeps the last synced value of each field. Whichever side changed since
then wins. If both changed the same field, the later edit wins (Sheet `updated_at` vs the Account's
`LastModifiedDate`). On the first sync for a newly linked retailer, the Sheet wins for any field
that has a value there, and Salesforce fills in the blanks. Dashboard saves also push to Salesforce
immediately; if that push fails, the next scheduled run retries it.

To unlink a retailer, clear `Dashboard_Retailer_ID__c` on its Account. The next sync removes the
link and contacts from the Sheet.

**Contacts are personal data.** They travel through the same web app URL as everything else, so
anyone with the dashboard password can see them.

---

## Updating the script later

If `retailers_sync.gs` is updated in the repo (e.g. a new column was added):

1. Copy the latest `retailers_sync.gs` content from the repo.
2. Open your Sheet → Extensions → Apps Script.
3. Select-all in the editor and paste over the existing code.
4. Hit **💾 Save**.
5. **Deploy → Manage deployments** → pencil-edit your existing deployment → **Version: New version** → **Deploy**. (Apps Script pins the web app URL to a specific version, so you have to redeploy for changes to take effect.)
6. Back in the dashboard, click **⬆ Push ALL retailers to Sheet** so the Sheet picks up the new schema.

The web app URL stays the same — no need to re-paste it.

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
