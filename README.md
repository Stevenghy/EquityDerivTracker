# Equity Derivatives Cross Tracker
## Bloomberg-Integrated Excel VBA System for IDB Broker Teams

---

## Architecture

```
┌─────────────┐     ┌──────────────┐     ┌─────────────┐
│  Bloomberg   │────▶│  ScanInput   │────▶│   DataLog   │
│  Terminal    │     │  (raw data)  │     │  (deduped)  │
└─────────────┘     └──────────────┘     └──────┬──────┘
                                                │
┌─────────────┐                          ┌──────▼──────┐
│  PrevClose  │─── premium calc ────────▶│  Dashboard  │
│  (closing   │                          │  (charts)   │
│   levels)   │                          └─────────────┘
└─────────────┘
```

## Sheets

| Sheet | Purpose |
|-------|---------|
| **ScanInput** | Raw Bloomberg scan data lands here. Has macro buttons for export/scan. |
| **PrevClose** | Previous day closing levels per product. Used for premium calculation. |
| **DataLog** | Master deduplicated trade log. IDB flagging, premium %, manual override. |
| **Dashboard** | Intraday scatter chart (premium % over time), summary stats, product pie chart. |

## VBA Modules

| Module | Purpose |
|--------|---------|
| `Module_Config` | All constants — column mappings, sheet names, thresholds, market hours |
| `Module_Utils` | Helper functions — dedup key, prev close lookup, premium calc, CSV export |
| `Module_ScanExport` | Main export: ScanInput → DataLog with dedup + IDB flagging |
| `Module_Dashboard` | Chart generation — intraday premium scatter, summary table, product pie |
| `Module_AutoScan` | Unattended scanning — scheduled Bloomberg refreshes during market hours |
| `Module_Setup` | One-time initialization — creates all sheets, headers, buttons |
| `ThisWorkbook` | Auto_Open handler for unattended mode via Task Scheduler |

## Setup Instructions

### Step 1: Create the Workbook
1. Open Excel → Save as `.xlsm` (Macro-Enabled Workbook)
2. Press `Alt+F11` to open VBA editor
3. Import each `.bas` file: File → Import File
4. For `ThisWorkbook.bas`: copy the code into the existing `ThisWorkbook` object

### Step 2: Initialize
1. Run `InitializeWorkbook` (press `Alt+F8`, select it)
2. This creates all 4 sheets with headers, formatting, and control buttons

### Step 3: Configure Products
1. Go to `PrevClose` sheet
2. Update product names to match your Bloomberg tickers
3. Fill in yesterday's closing levels (or use Bloomberg BDP formulas)

### Step 4: Configure Bloomberg Scan
1. Go to `ScanInput` sheet
2. Set up your Bloomberg BSRCH/BDS formulas to populate columns A-G
3. Adapt `RunBloombergScan()` in Module_AutoScan to your specific BB setup

### Step 5: Daily Use
1. Bloomberg scan populates ScanInput
2. Click **"Export to Log"** → dedupes + calculates premium + flags IDB trades
3. Click **"Refresh Charts"** → updates Dashboard with today's data
4. Review DataLog → use **Manual Override** column to correct IDB flags
5. Click **"Export CSV"** for external analysis

## IDB Detection Logic

```
Premium % = Cross Level / Previous Close

Display as: 100.XX%

If premium rounds cleanly to 2dp → "LIKELY IDB"    (green)
If premium has >2dp noise       → "LIKELY CLIENT"   (red)
If no close data available      → "NO CLOSE DATA"   (yellow)
```

**Example:**
- Cross Level: 22,150.00 | Prev Close: 22,080.00
- Premium: 100.32% → LIKELY IDB ✅
- Premium: 100.3173% → LIKELY CLIENT ❌

## Unattended Mode (Days Off)

### Option A: Task Scheduler (Recommended)
1. Edit `TaskScheduler_Setup.bat` with your Excel and workbook paths
2. Run as Administrator
3. Set `AUTO_SCAN_ON_OPEN = True` in ThisWorkbook
4. Excel opens at 08:55, scans every 30 min, closes at 17:30

### Option B: Manual
- Click **"Start Auto Scan"** before leaving
- It scans every 30 min during market hours
- Auto-exports CSV at market close

## Control Buttons (ScanInput sheet)

| Button | Action |
|--------|--------|
| **Export to Log** | Deduplicate + export scan data to DataLog |
| **Clear Scan** | Clear ScanInput data after export |
| **Start Auto Scan** | Begin scheduled Bloomberg scanning |
| **Stop Auto Scan** | Cancel scheduled scans |
| **Refresh Charts** | Rebuild Dashboard charts for today |
| **Update Prev Close** | Pull fresh closing levels from Bloomberg |
| **Export CSV** | Export DataLog to timestamped CSV file |

## Dashboard Visualizations

1. **Intraday Premium Scatter** — Premium % (100.xx) plotted over time, one series per product. Shows price drift throughout the day.
2. **Today's Summary** — Total crosses, IDB count, client count, total notional.
3. **Product Pie Chart** — Trade distribution across products.

## Customization

- **Add products:** Edit `PrevClose` sheet + update sample array in Module_Setup
- **Change scan interval:** Edit `AUTO_SCAN_INTERVAL_MIN` in Module_Config
- **Change market hours:** Edit `MARKET_OPEN_HOUR` / `MARKET_CLOSE_HOUR`
- **IDB threshold:** Edit `IDB_DP_THRESHOLD` (default: 2 decimal places)
- **Bloomberg tickers:** Adapt formulas in ScanInput and `UpdatePrevClose()`
