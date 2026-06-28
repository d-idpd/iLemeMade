# iLemeMade — Apps Script Setup

This folder contains the Google Apps Script backend that turns your Google Sheet into a REST API.

## One-Time Setup

### Step 1 — Open the Script Editor

1. Open the spreadsheet: https://docs.google.com/spreadsheets/d/1oRLuRspyOfbo-_x9uZJzSGhn_DKa7CrDnyZOlg1PKL8
2. Click **Extensions → Apps Script**
3. Delete any placeholder code in the editor

### Step 2 — Paste the Code

1. Open `Code.gs` from this folder
2. Copy the entire file contents
3. Paste it into the Apps Script editor
4. Click **Save** (floppy disk icon or Ctrl+S)
5. Name the project: `iLemeMade API`

### Step 3 — Run First-Time Setup

1. In the function dropdown (top toolbar), select **`firstTimeSetup`**
2. Click **Run**
3. Grant permissions when prompted (this is your own script on your own sheet — it's safe)
4. Check the spreadsheet — you should see all 6 tabs created with headers and sample data

### Step 4 — Deploy as a Web App

1. Click **Deploy → New deployment**
2. Click the gear icon next to "Type" and choose **Web app**
3. Fill in the fields:
   - **Description:** `iLemeMade API v1`
   - **Execute as:** `Me`
   - **Who has access:** `Anyone`
4. Click **Deploy**
5. Copy the **Web app URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfy.../exec
   ```
6. Paste that URL into `js/config.js` as `SHEETS_API_URL`

### Step 5 — Test the API

Open this URL in your browser (replace with your actual URL):

```
https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?action=ping
```

You should see:
```json
{ "status": "ok", "message": "iLemeMade API is running!" }
```

Test the dashboard:
```
?action=getDashboard
```

Test products list:
```
?action=getProducts
```

## API Reference

| action | Description |
|--------|-------------|
| `ping` | Health check |
| `getProducts` | All product templates |
| `getInventory` | All inventory items (optional: `?status=Available` or `?productId=HK-001`) |
| `getItem&itemId=HK-001-001` | Single item with costs, labor, and sale data |
| `getSales` | All sales (optional: `?channel=Etsy`) |
| `getDashboard` | Calculated metrics: revenue, profit, savings |
| `getCosts&itemId=HK-001-001` | Cost breakdown for one item |
| `getLabor&itemId=HK-001-001` | Labor record for one item |
| `setup` | Create sheet tabs and headers (safe to re-run) |
| `seedData` | Add sample records (skips if data already exists) |

## Re-deploying After Changes

When you edit `Code.gs`, you must create a new deployment version:

1. Click **Deploy → Manage deployments**
2. Click the pencil icon
3. Change **Version** to **New version**
4. Click **Deploy**

The URL stays the same — no need to update `config.js`.

## Statuses for Inventory Items

| Status | Meaning |
|--------|---------|
| `Available` | In stock, ready to sell |
| `Sold` | Already sold |
| `Reserved` | Held for a specific buyer |
| `Donated` | Given away |
| `Damaged` | Not for sale |

## Sales Channels

Use consistent values in the Channel column:
- `In Person`
- `Etsy`
- `Instagram`
- `Gift`
