# gsheet-snowflake-connector

Google Sheets Apps Script to run Snowflake queries directly from a spreadsheet. Supports saved queries, named sheet tabs, and daily auto-refresh via the Snowflake SQL API v2 with Programmatic Access Tokens (PAT).

---

## Prerequisites

- A Google account with access to Google Sheets
- A Snowflake account with permission to generate a Programmatic Access Token (PAT)

---

## Installation

### 1. Open the Apps Script editor

1. Open (or create) the Google Sheet where you want to use this connector.
2. In the top menu, go to **Extensions → Apps Script**.

### 2. Add the script files

You need to create four files in the Apps Script editor.

**`code.gs`**

Paste the contents of [code.gs](code.gs) into the editor. If the default project already has a `Code.gs` file, replace its contents entirely.

**`sidebar.html`**

1. Click the **+** next to "Files" and choose **HTML**.
2. Name it `sidebar` (no extension).
3. Paste the contents of [sidebar.html](sidebar.html).

**`config.html`**

1. Add another HTML file named `config`.
2. Paste the contents of [config.html](config.html).

**`scheduler.html`**

1. Add another HTML file named `scheduler`.
2. Paste the contents of [scheduler.html](scheduler.html).

### 3. Set up the onOpen trigger

The menu **❄️ Snowflake** is built by `buildSnowflakeMenu()`. You need to register it as an `onOpen` trigger so it appears every time the spreadsheet is opened.

**Option A — via the Apps Script trigger UI (recommended)**

1. In the Apps Script editor, click the clock icon (**Triggers**) in the left sidebar.
2. Click **+ Add Trigger** (bottom right).
3. Set:
   - Function to run: `buildSnowflakeMenu`
   - Event source: `From spreadsheet`
   - Event type: `On open`
4. Click **Save** and authorize when prompted.

**Option B — run `installSnowflakeTrigger` once**

1. Open [code.gs](code.gs) and update the `ssId` constant in `installSnowflakeTrigger` to your spreadsheet's ID (found in the sheet URL: `https://docs.google.com/spreadsheets/d/<ID>/edit`).
2. In the editor, select `installSnowflakeTrigger` from the function dropdown and click **Run**.
3. Authorize when prompted.

### 4. Authorize the script

The first time you run a function (or when the trigger fires), Google will ask you to review and approve the permissions the script needs (network access, spreadsheet access, properties storage). Click **Review permissions → Allow**.

### 5. Reload the spreadsheet

Close and reopen the sheet (or refresh the page). The **❄️ Snowflake** menu should now appear in the top menu bar.

---

## Snowflake setup

### 1. Create a network policy

Google Apps Script makes requests from Google's IP ranges, which Snowflake may block by default. Create a network policy that allows all IPs and assign it to your user:

```sql
CREATE NETWORK POLICY GOOGLE_SHEETS_POLICY
  ALLOWED_IP_LIST = ('0.0.0.0/0')
  COMMENT = 'Allow all IPs for Google Sheets Apps Script connector';

ALTER USER <your_username> SET NETWORK_POLICY = GOOGLE_SHEETS_POLICY;
```

Replace `<your_username>` with your Snowflake username. This requires the `SECURITYADMIN` role or higher.

### 2. Get a Programmatic Access Token (PAT)

1. Log into your Snowflake account.
2. Click your username in the bottom-left corner.
3. Select **Programmatic access tokens**.
4. Click **+ Generate** and copy the token — you will not be able to see it again.

---

## Configuration

1. In Google Sheets, click **❄️ Snowflake → ⚙️ Configure**.
2. Fill in the fields:

| Field | Required | Description |
|---|---|---|
| Account Identifier | Yes | From your Snowflake URL: `https://<account>.snowflakecomputing.com` |
| Programmatic Access Token | Yes | The PAT you generated above |
| Warehouse | No | Default warehouse for queries (e.g. `ANALYTICS`) |
| Database | No | Default database (e.g. `MY_DATABASE`) |
| Schema | No | Default schema (e.g. `PUBLIC`) |
| Role | No | Default role (e.g. `ANALYTICS_ROLE`) |
| Slack Webhook URL | No | Incoming Webhook URL to receive failure alerts (see [Slack notifications](#slack-notifications)) |

3. Click **💾 Save Configuration**.

---

## Usage

### Run a query

1. Click **❄️ Snowflake → 🔍 Run Query** to open the sidebar.
2. Enter a query name (this becomes the sheet tab name) and write your SQL.
3. Click **Run** — results are written to a new tab named after your query.

### Save a query for later

In the sidebar, toggle **Save query** before running. Saved queries appear in the sidebar and can be re-run or deleted.

### Enable daily auto-refresh

When saving a query, check **Auto-refresh daily**. The script will automatically run all auto-refresh queries at ~6 AM every day and write fresh results to their respective tabs.

### Refresh all queries manually

Click **❄️ Snowflake → 🔄 Refresh All** to re-run every saved query that has auto-refresh enabled.

---

## Slack notifications

When a query fails during auto-refresh, the connector will post a Slack message listing every query that failed and how many attempts were made.

Failed queries are retried up to **3 times** with a **3-second pause** between attempts. The Slack alert fires only after all retries are exhausted.

### Set up a Slack Incoming Webhook

1. Go to [api.slack.com/apps](https://api.slack.com/apps) and click **Create New App → From scratch**.
2. Name the app (e.g. `Snowflake Alerts`) and select your workspace.
3. In the left sidebar, go to **Incoming Webhooks** and toggle it **On**.
4. Click **Add New Webhook to Workspace**, choose the channel where alerts should appear, and click **Allow**.
5. Copy the webhook URL — it looks like `https://hooks.slack.com/services/T.../B.../...`
6. Open **❄️ Snowflake → ⚙️ Configure** in your spreadsheet, paste the URL into the **Slack Webhook URL** field, and save.

If the field is left blank, Slack notifications are disabled and the behavior is unchanged.
