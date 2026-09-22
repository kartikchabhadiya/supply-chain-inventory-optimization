# 📦 Supply Chain & Inventory Optimization Tool

An Excel-based inventory planning system built on the Kaggle **"Supply Chain Logistics Problem"** dataset (9,215 order-level shipment records across 772 SKUs and 7 plants). The workbook calculates Safety Stock, Reorder Point (ROP), Economic Order Quantity (EOQ) and Inventory Turnover for every SKU, flags items needing restock, ranks SKUs with an ABC (Pareto) analysis, and lets a manager stress-test the plan against lead-time delays and demand spikes with an interactive What-If engine — all fully formula-driven, zero hardcoded results.

📊 **[Download / view the workbook](.supply-chain-inventory-optimization.xlsx)**
  **[Download / view the workbook](.supply-chain-inventory-optimization (2).xlsx)**
  **[Dashboard](images/dashboard_screenshot.png)**
---

## 🔍 What it does

| Module | Description |
|---|---|
| **Reorder Point Alerts** | Live `CRITICAL / WARNING / HEALTHY` status per SKU based on Stock On Hand vs. Safety Stock and ROP |
| **ABC (Pareto) Analysis** | SKUs auto-ranked by inventory value with cumulative % and A/B/C classification — re-sorts itself if inputs change |
| **EOQ & Safety Stock Engine** | Classic inventory-theory formulas (EOQ, Safety Stock, ROP) computed per SKU from real order-level demand and lead-time data |
| **What-If Scenario Engine** | Dropdown-driven Lead-Time Delay and Demand Spike controls, a Base-vs-Scenario KPI comparison, and a two-factor sensitivity grid across every delay × spike combination |
| **Executive Dashboard** | KPI cards, ABC value-split chart, stock-health chart, and a top-10-SKUs-by-value table, all live-linked to the model |

## 📐 Key metrics, at a glance

- **772 SKUs** across **7 plants**, built from **9,215 raw order lines**
- **136 Critical / 280 Warning / 356 Healthy** SKUs under baseline conditions
- Class A SKUs (~24% of items) drive **~80%** of total inventory value
- Lead times in the dataset are short (0–4 days) — the sensitivity grid shows even a 2–5 day shipping delay pushes most SKUs into reorder territory, meaning the network runs with very little slack

## 🗂️ Workbook structure

- **Read Me** — navigation, assumptions, and a color legend
- **Assumptions** — editable global inputs (order cost, ABC thresholds); change these and the whole model recalculates
- **Raw_Orders / Raw_WhCosts / Raw_WhCapacities** — cleaned source data (untouched)
- **Inventory_Calculations** — the core engine: one row per SKU, every formula (Safety Stock, ROP, EOQ, Turnover, live status, and a parallel What-If recalculation)
- **ABC_Analysis** — formula-driven Pareto ranking and classification
- **What-If Scenario** — scenario controls, KPI deltas, and two-factor sensitivity grids
- **Dashboard** — the executive summary view

## 🧮 Methodology & formulas

Built entirely with native Excel functions — no VBA, no add-ins:

- `SUMPRODUCT` for a fully transparent, portable equivalent of Excel's What-If Data Tables (each sensitivity-grid cell is an independent formula, not a dependent table)
- `INDEX` / `MATCH` for all cross-sheet lookups (avoids volatile `VLOOKUP`/`OFFSET`)
- `IFERROR` guards on every lookup and ratio to prevent propagated errors
- Standard inventory-theory formulas:
  - `Safety Stock = MAX(0, (Max Daily Sales × Max Lead Time) − (Avg Daily Sales × Avg Lead Time))`
  - `Reorder Point = (Avg Daily Sales × Avg Lead Time) + Safety Stock`
  - `EOQ = √(2 × Annual Demand × Order Cost ÷ Holding Cost per Unit)`
  - `Average Inventory = Safety Stock + EOQ / 2` (used for Turnover, so the ratio reflects the optimized policy rather than a single stock snapshot)
- Conditional formatting drives the live Critical/Warning/Healthy alerts
- Data-validation dropdowns power the What-If scenario controls

**Verification:** every one of the workbook's **28,719 formulas** recalculates cleanly with zero errors (checked with a full LibreOffice recalculation pass).

## 📝 Assumptions & data-modeling notes

The source dataset is a single-day order backlog (shipment/logistics data, not a sales or warehouse ledger), so several inputs are modeled as clearly-labeled proxies rather than presented as raw fact:

- **Demand proxy** — each order line is treated as one demand observation (no daily POS history exists in the source data)
- **Lead time** — taken directly from the dataset's transit-time field (a genuine measured value)
- **Holding cost** — taken directly from the warehouse cost source sheet (a genuine value)
- **Order cost per PO** — not in the source data; set as an editable assumption (default $75, a typical mid-size-manufacturer figure)
- **COGS** — modeled as Annual Demand × Holding Cost/unit, since no unit sale price exists in the source data
- **Stock On Hand** — not present in the source data at all; simulated once per SKU (a random multiple of that SKU's Reorder Point) purely so the alert system has a realistic mix of statuses to demonstrate. These cells are shaded and marked as exactly where a real WMS/ERP on-hand extract should be pasted in — every downstream formula recalculates automatically

All of this is documented in-workbook on the **Read Me** tab, including a full color legend for editable vs. calculated cells.

## 🛠️ Skills demonstrated

Excel formula design (`INDEX`/`MATCH`, `SUMPRODUCT`, `IFERROR`), inventory theory (EOQ, Safety Stock, ROP, ABC/Pareto analysis), scenario/sensitivity modeling, conditional formatting, data validation, dashboard design, and clear documentation of assumptions and data limitations.

## 📚 Source data

Kaggle ["Supply Chain Logistics Problem"](https://www.kaggle.com/datasets/laurinbrechter/supply-chain-logistics-problem) dataset (OrderList, FreightRates, WhCosts, WhCapacities, ProductsPerPlant, VmiCustomers, PlantPorts).

---