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
| K | `rep_firm` | ✅ | Free text. The dashboard dropdown shows `Unassigned / Brian Schlager / Internal / Jeff Day / PSE / StoR`; you can type any value here and it shows in the dashboard, but only those six appear in the dropdown choices. Blank = Unassigned. |
| L | `status` | ✅ | One of: `in`, `pitched`, `target`, `non-target`, `declined` (or blank). |
| M | `next_steps` | ✅ | Free text. |
| N | `updated_at` | ❌ | Auto-stamped on every write. |
| O | `next_review` | ✅ | Date (`yyyy-mm-dd`). The dashboard shows it red once the date has passed. Sheets made before columns O–P existed get the headers added in place on the next read or write; no data is cleared. |
| P | `priority` | ✅ | Whole number 1–10 (1 = highest), or blank. Anything else is ignored on pull. The dashboard sorts each channel by priority, then store count. |
| W | `needs_distributor` | ✅ | `1` when the retailer buys through a distributor, blank otherwise. Dashboard-only; not sent to Salesforce. |
| X | `distributor_name` | ✅ | Free text — which distributor. Dashboard-only. |
| AA | `reset_date` | ✅ | Date (`yyyy-mm-dd`) of the retailer's shelf reset / modular. Hidden in the dashboard until you tick "Show reset date". |
| Z | `key_contact` | ✅ | Email (or name) of the Account contact marked as our key contact in the dashboard. Dashboard-only; Salesforce contacts themselves are never written. |
| Y | `deadline` | ✅ | Date (`yyyy-mm-dd`) the deal has to land by; shows red in the dashboard once it passes. Used as the Opportunity's CloseDate when set. |
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
| Retailer Rep Firm | `Retailer_Rep_Firm__c` | Picklist, not restricted: Unassigned, Brian Schlager, Internal, Jeff Day, PSE, StoR |
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
   - `installAllTriggers`: turns on the 15-minute sync **and** the instant push
     when someone types into the Sheet.

### Auto-sync

Three paths keep the two sides together, all on by default once `installAllTriggers` has run:

| Change made in | Reaches the other side |
|---|---|
| Dashboard | Immediately (the dashboard pushes on every save) |
| The Sheet | Within seconds (`sfOnSheetEdit`, an installable onEdit trigger) |
| Salesforce | Within 15 minutes (`syncSalesforce`), or on the next dashboard pull |

The dashboard pulls when you open the Retailers tab, when you come back to the
browser window, and every 60 seconds while that tab is open.

### What reaches Salesforce

Only retailers whose status is **Target** or **Pitched** put field values into Salesforce.
Every other row (Currently In, Non-Target, Declined, blank) stays in the dashboard and the
Sheet: no push, no pull, and the Account keeps its link plus whatever it already held. Account
names and contacts still refresh for every linked row, so the dashboard hover card stays current.

When a retailer drops out of Target/Pitched, the sync stops writing and logs one
`(left synced set)` row on SF_Log. It never blanks what it already wrote, so the Account can
hold a stale status. Run `clearUnsyncedFields` by hand to blank those five fields on every
Account whose retailer is no longer Target/Pitched; it logs each value it clears.

### Opportunities

A retailer reaching **Target** or **Pitched** gets one Catalyst cat-litter Opportunity on its
Account, created by the next sync:

| Field | Value |
|---|---|
| Name | `Catalyst Cat Litter - {Retailer} {close year}` |
| StageName | Target → Qualification, Pitched → Proposal |
| CloseDate | the retailer's Deadline, else its Next Review, else 90 days out |
| Amount | the dashboard's annual wholesale sizing (blank when the row has no stores) |
| ForecastCategoryName | Omitted, so a top-down estimate never lands in a forecast total |
| Type | New Business |
| Product_Category__c | Pet Litter |
| Product_Label__c | Catalyst |
| Sales_Channel__c | Mass and Club → Big Box, Pet Specialty → Pet Specialty, Grocery → Grocery; blank otherwise |
| Problem_Statement__c | the row's Next Steps at creation |

Never more than one per retailer. If the Account already carries an Opportunity tagged Pet
Litter, open or closed, the sync adopts that one instead of creating another. The Opportunity's
Id and stage come back into the Sheet (columns U–V) and show as a badge on the dashboard row,
linking straight to the record.

Stage moves in one direction only: the sync nudges Qualification ↔ Proposal to match the
retailer's status, and never touches an Opportunity someone has advanced to Negotiation,
Documentation or a closed stage. Dropping a retailer off the target list leaves its Opportunity
alone for a human to close. Amount is set once at creation and never rewritten.

Every sync also sweeps the whole org for Opportunities tagged `Pet Litter` (any value in
`SF_OPP_CATEGORIES`) or labeled `Catalyst`. Ones on a linked Account attach to their retailer row.
Ones on an Account no retailer is linked to land on an **SF_Opps** tab and show in a panel under
the dashboard's retailer table, so a deal opened outside this list is still visible. Link that
Account (stamp `Dashboard_Retailer_ID__c`) to pull it onto a row.

`previewOpportunities` lists what the next sync would create, without writing.

### Overwrite warnings

Before any push, the script reads the Account and compares each field against the
last synced value. Anything it is about to replace that somebody changed in
Salesforce gets a row on the **SF_Log** tab, newest first: when, retailer, account,
field, the value kept, the value overwritten, which side won and what triggered it.
Scheduled syncs log the same way, including both-sides-changed conflicts and
first-link overwrites. The tab keeps the most recent 500 rows.

Add a Script Property `SF_ALERT_EMAIL` to also get an email each time one happens.
Leave it unset for no email.

`previewSalesforceSync` is a dry run: it logs every field the next sync would write,
and flags conflicts, without sending anything to Salesforce. Worth running before a
big round of edits.

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

## Talking to the pipeline through Claude (MCP)

The same web app speaks MCP, so anyone on the team can work the list in conversation instead of
opening the dashboard. No install: in Claude, **Settings → Connectors → Add custom connector**,
paste **`https://catalyst-retailers-mcp.fly.dev/mcp`**, and four tools appear.

That address is a thin proxy (`mcp_proxy/`, running on Fly) in front of the Apps Script
deployment. Claude's connector check refuses the 302 redirect every Apps Script web app answers
with, so the `/exec` URL can't be used directly even though it speaks MCP correctly. The proxy
holds no logic — it forwards each JSON-RPC message and returns the answer. When the Apps Script
deployment URL changes, update the proxy's secret and the connector URL stays put:

    flyctl secrets set APPS_SCRIPT_URL="<new /exec url>" -a catalyst-retailers-mcp

| Tool | What it does |
|---|---|
| `find_retailers` | Search by name, status, rep firm, channel, or only overdue deadlines |
| `get_retailer` | One retailer in full, including Salesforce contacts and its opportunity stage |
| `update_retailer` | Set status, priority, rep firm, next steps, deadline, next review, reset date, distributor, key contact, store count |
| `pipeline_summary` | Counts by status and rep firm, deadlines passed, reviews coming up |

Writes go through the same upsert path as a dashboard edit, so the Salesforce sync, the
Target/Pitched gate and the SF_Log overwrite warnings all behave identically. Only one retailer
changes per call, and a name that matches several rows changes nothing and asks which one.

Store counts are the exception worth knowing: the dashboard owns them, so a count set here is
overwritten the next time someone runs **Push ALL retailers to Sheet**. For a permanent change,
edit `RT_RETAILERS` in `dashboard_template.html`.

**Access.** The connector inherits whatever the web app deployment allows. Deployed as "Anyone",
the URL is the only gate — treat it like a password and share it directly, not in a public channel.

---

## Deploying without the copy-paste (clasp)

Pasting the file into the editor and cutting a new version by hand is the slow way, and
creating a *new* deployment instead of a version changes the `/exec` URL, which then has to be
re-pointed everywhere. `clasp`, Google's Apps Script CLI, does both steps in one command and
always updates the same deployment:

    ./deploy_apps_script.sh "what changed"

### One-time setup

1. **Install and log in** (the login opens a browser once):

       npm install -g @google/clasp
       clasp login

2. **Turn on the Apps Script API** for the Google account that owns the Sheet:
   <https://script.google.com/home/usersettings> → *Google Apps Script API* → **On**.

3. **Collect two ids.**
   - *scriptId*: in the Apps Script editor, Project Settings → IDs → Script ID.
   - *deploymentId*: Deploy → Manage deployments → the active web app → the id under its name
     (it looks like `AKfycb…`, longer than the URL fragment).

4. **Write `google_apps_script/.clasp.json`** (gitignored, since the ids are account-specific):

       {
         "scriptId": "<script id>",
         "deploymentId": "<deployment id>",
         "rootDir": "src"
       }

After that, every change to `retailers_sync.gs` ships with one command, the `/exec` URL stays
put, and the MCP proxy needs no secret update. The script copies the `.gs` files into `src/`
next to `appsscript.json` before pushing, so the repo keeps its flat layout.

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
