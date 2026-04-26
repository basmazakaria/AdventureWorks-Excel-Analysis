# 📊 Adventure Works Sales Analysis
> End-to-end Business Intelligence project built entirely in **Microsoft Excel** using the AdventureWorks database.

---

## 🗂️ Project Overview

This project demonstrates a full BI pipeline — from raw relational data to a clean, interactive dashboard — using only Excel's built-in tools: **Power Query**, **Data Modeling**, and **Excel Charts**.

---

## 🛠️ Tools & Technologies

| Tool | Purpose |
|------|---------|
| Microsoft Excel | Main platform |
| Power Query | Data extraction & transformation (ETL) |
| Excel Data Model | Star schema relationships |
| SQL Server (AdventureWorks) | Data source |

---

## ⚙️ ETL — Power Query Steps

- Connected to **SQL Server** (AdventureWorks database)
- Merged `OrderHeader` and `OrderDetail` tables
- Extracted a custom `Date_Key` column (`yyyymmdd` format) for relationship mapping
- Cleaned and shaped data for the fact table: `Fact_Orders`

---

## 🗃️ Data Model — Star Schema

```
Dim_Territory ──┐
Dim_ShipMethod ─┤
                ├──► Fact_Orders
Dim_Product ────┤
Dim_Date ───────┘
```

**Fact Table:** `Fact_Orders`
- SalesOrderID, OrderDate, CustomerID, TerritoryID, ShipMethodID
- SubTotal, TaxAmt, Freight, TotalDue
- OrderQty, ProductID, LineTotal, Date_Key

**Dimension Tables:**
- `Dim_Territory` → TerritoryID, Territory
- `Dim_ShipMethod` → ShipMethodID, Shipmethod
- `Dim_Product` → ProductID, Products, Color, Subcategory, Category
- `Dim_Date` → DateKey, Date, Year, Quarter, Month

---

## 📈 Dashboard Highlights

| Metric | Value |
|--------|-------|
| 💰 Total Sales | $2,926,970,124 |
| 👥 Num. Customers | 19,119 |
| 📦 Num. Orders | 31,465 |
| 🔢 Total Qty | 274,914 |
| 🚚 Total Freight | $78,690,398 |

### Key Visuals
- 📅 **Total Sales Per Month** — Line chart showing monthly trends
- 🌍 **Total Sales Per Territory** — Bar chart across all regions
- 🏷️ **Num. of Orders Per Category** — Horizontal bar chart
- 🥧 **Total Sales Per Category** — Pie chart breakdown

### Filters / Slicers
- Ship Method
- Product Category
- Product Subcategory

---

## 🔍 Key Insights

- 📈 **August** was the peak month with **$391M** in sales
- 🌎 **Southwest** was the top territory with **$696M** in total revenue
- 🚲 **Bikes** dominate with **41%** of total sales
- 🧩 **Components** had the lowest order count (2,650) — a potential growth opportunity
- 👟 **Accessories** had the highest number of orders (19,524)

---

## 📁 Repository Structure

```
📦 AdventureWorks-Excel-Analysis
 ┣ 📊 AdventureWorks_Dashboard.xlsx   # Main Excel file (Power Query + Model + Dashboard)
 ┣ 📸 screenshots/
 ┃  ┣ dashboard.png
 ┃  ┣ data_model.png
 ┃  └ power_query.png
 ┗ 📄 README.md
```

---

## 🚀 How to Use

1. Clone or download the repository
2. Open `AdventureWorks_Dashboard.xlsx` in Microsoft Excel
3. If prompted, update the data source connection to your local SQL Server
4. Refresh Power Query (`Data → Refresh All`)
5. Explore the dashboard using the slicers

---

## 🙋 Author

Feel free to connect on [LinkedIn](#) or open an issue if you have any questions!
