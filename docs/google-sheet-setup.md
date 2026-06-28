# Google Sheets Setup Guide

This guide walks through setting up the iLemeMade data backend in Google Sheets with a Google Apps Script API layer.

---

## Sheet Structure

The spreadsheet has **6 tabs**. Here is the full schema:

### Products

| Column | Type | Example | Notes |
|--------|------|---------|-------|
| ProductID | Text | `HK-001` | Unique code per product line |
| ProductName | Text | `Hello Kitty Resin` | Display name |
| Category | Text | `Resin Art` | Used for filtering and color coding |
| Description | Text | `Hello Kitty in clear resin` | Short description |
| DefaultMaterialCost | Number | `3.00` | Typical material cost in dollars |
| DefaultLaborMinutes | Number | `60` | Typical time to make one item |
| DefaultSalePrice | Number | `18.00` | Suggested retail price |

**Product ID format:** `XX-NNN` where `XX` is the category code.

| Code | Product |
|------|---------|
| BC | BearCat Keychain |
| HK | Hello Kitty Resin |
| CB | Chicken Banana |
| DR | 3D Printed Dragon |
| RL | Resin Lemon |
| SH | Trail Shoe Art |
| RF | Recycled Filament Casting |
| BR | Handmade Bracelet |
| RK | Hand Painted Rock |

---

### Inventory

| Column | Type | Example | Notes |
|--------|------|---------|-------|
| ItemID | Text | `HK-001-003` | Unique ID for each physical piece |
| ProductID | Text | `HK-001` | Links to Products tab |
| Status | Text | `Available` | See status table below |
| Color | Text | `Rainbow` | Color description |
| DateCreated | Date | `2025-04-01` | When this item was made |
| PhotoURL | Text | `https://...` | Link to photo (Google Drive or Imgur) |
| Story | Text | `Rainbow swirl — my favorite!` | Kid-written story about this piece |

**Item ID format:** `ProductID-NNN` (e.g., `HK-001-003` is the 3rd Hello Kitty item)

**Status values:**
- `Available` — In stock, ready to sell
- `Sold` — Already sold
- `Reserved` — Held for a specific buyer
- `Donated` — Given away
- `Damaged` — Not for sale

---

### Costs

| Column | Type | Example | Notes |
|--------|------|---------|-------|
| ItemID | Text | `HK-001-003` | Links to Inventory |
| ResinCost | Number | `3.00` | Cost of resin used |
| FilamentCost | Number | `0.00` | Cost of 3D filament |
| PaintCost | Number | `0.50` | Paint and pigment cost |
| LEDCost | Number | `0.00` | LED components |
| HardwareCost | Number | `0.25` | Chains, hooks, rings |
| PackagingCost | Number | `0.50` | Bag, tag, box |
| OtherCost | Number | `0.00` | Anything else |

**MaterialCost = ResinCost + FilamentCost + PaintCost + LEDCost + HardwareCost + PackagingCost + OtherCost**

---

### Labor

| Column | Type | Example | Notes |
|--------|------|---------|-------|
| ItemID | Text | `HK-001-003` | Links to Inventory |
| MinutesWorked | Number | `75` | Total minutes spent making this item |
| HourlyRate | Number | `5.00` | Your pay rate per hour |

**LaborCost = (MinutesWorked / 60) × HourlyRate**

---

### Sales

| Column | Type | Example | Notes |
|--------|------|---------|-------|
| Date | Date | `2025-04-01` | Date of sale |
| ItemID | Text | `HK-001-003` | Item that was sold |
| SalePrice | Number | `18.00` | How much it sold for |
| Channel | Text | `Etsy` | Where it was sold |
| Notes | Text | `School fair` | Any notes |

**Channel options:** In Person, Etsy, Instagram, Gift

---

### Dashboard

This tab is written automatically by the Apps Script every time `getDashboard` is called.
You don't need to edit it by hand.

| Metric | Description |
|--------|-------------|
| revenue | Total of all sale prices |
| expenses | Total costs (materials + labor) for sold items |
| profit | revenue − expenses |
| inventoryCount | Number of Available items |
| itemsSold | Number of unique items with a sale record |
| bestSeller | Product name with highest revenue |
| savingsBalance | 50% of profit |
| reinvestmentBalance | 25% of profit |
| spendingBalance | 25% of profit |

---

## Cost Formulas

```
MaterialCost = ResinCost + FilamentCost + PaintCost + LEDCost + HardwareCost + PackagingCost + OtherCost

LaborCost = (MinutesWorked / 60) × HourlyRate

TotalCost = MaterialCost + LaborCost

Profit = SalePrice − TotalCost
```

---

## Savings Rules

Every dollar of profit gets split three ways:

| Bucket | Percentage | Purpose |
|--------|-----------|---------|
| Savings | 50% | Long-term savings |
| Business Reinvestment | 25% | Buy more supplies |
| Spending Money | 25% | Fun money! |

---

## Photo Tips

To add photos:
1. Take a photo of the item
2. Upload to Google Drive (or any public image host)
3. Right-click → Get link → Change to "Anyone with the link can view"
4. Paste the link in the `PhotoURL` column

For Google Drive direct-embed URLs, convert:
```
https://drive.google.com/file/d/FILE_ID/view
```
to:
```
https://drive.google.com/uc?id=FILE_ID
```

---

## Adding a New Item (Workflow)

When you make something new:

1. **Products tab** — if it's a new product type, add a row
2. **Inventory tab** — add a row with the new ItemID, status = `Available`
3. **Costs tab** — add a row with what it cost to make
4. **Labor tab** — add a row with how long it took

When it sells:

5. **Sales tab** — add a row with the date, price, and channel
6. The item status updates automatically via the API

---

## Troubleshooting

**API returns `{ "error": "Sheet not found: Products" }`**
→ Run `firstTimeSetup()` in the Script Editor to create all tabs.

**Data is stale / not updating**
→ After editing `Code.gs`, create a new deployment version (Deploy → Manage deployments → pencil icon → New version).

**Permission error when running the script**
→ Click "Review permissions" and log in with the Google account that owns the spreadsheet.

**`SHEETS_API_URL` is empty**
→ You haven't deployed the web app yet. See `apps-script/README.md` Step 4.
