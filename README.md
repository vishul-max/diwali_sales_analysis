# 🪔 Diwali Sales Analysis — EDA & Power BI Dashboard

An end-to-end exploratory data analysis (EDA) of Diwali-season retail sales in India, covering data profiling, cleaning, Power BI modelling (Power Query + DAX) and an interactive dashboard.

> **Headline result:** ₹106.25M revenue across 11,239 sales records. Women aged 26–45 drive most of the revenue, and **Food, Clothing, Electronics and Footwear** together generate **76.8%** of it.

---

## 📑 Table of Contents
1. [Project Overview](#-project-overview)
2. [Repository Structure](#-repository-structure)
3. [Dataset](#-dataset)
4. [Data Quality & Cleaning](#-data-quality--cleaning)
5. [Power BI Implementation](#-power-bi-implementation)
6. [Key Findings](#-key-findings)
7. [Recommendations](#-recommendations)
8. [Limitations & Assumptions](#-limitations--assumptions)
9. [How to Run](#-how-to-run)
10. [Tools Used](#-tools-used)

---

## 🎯 Project Overview

**Business problem:** Which customers, regions and product categories generate the most revenue during Diwali, and where should marketing and inventory effort go?

**Objectives**
- Profile and clean the raw sales data.
- Understand sales by gender, age group, marital status, occupation, state/zone and product category.
- Build a Power BI dashboard that answers these questions interactively.
- Turn the findings into actionable recommendations.

---

## 📁 Repository Structure

```
diwali-sales-eda-powerbi/
├── README.md
├── data/
│   ├── raw/
│   │   └── diwali_data.xlsx              # Raw data + Excel pivots/dashboard (Data, Pivot_Tables, Analysis, Dashboard)
│   └── processed/
│       └── Diwali_processed_Data.xlsx    # Same records, currency columns formatted as text (₹ 23,952)
├── powerbi/
│   └── Diwali_Sales_Dashboard.pbix       # Power BI report (add your .pbix here)
├── docs/
│   ├── Diwali_EDA_Report.docx            # Full structured EDA report
│   └── Diwali_Project_Presentation.pptx  # Presentation deck
└── assets/
    └── charts/                           # Chart images used in this README
```

---

## 📊 Dataset

| Property | Value |
|---|---|
| Records (rows) | 11,239 |
| Columns | 15 (14 usable, `Status` is empty) |
| Grain | One row = one sales record (customer × product line) |
| Currency | Indian Rupee (₹) |
| Time coverage | **No date column** — time-series analysis is not possible |
| Geography | 16 states in 5 zones |

### Data Dictionary

| Column | Type | Description |
|---|---|---|
| `User_ID` | Integer | Customer identifier (⚠️ not unique or consistent — see below) |
| `Cust_name` | Text | Customer first name |
| `Product_ID` | Text | Product identifier (2,350 distinct) |
| `Gender` | Text | `F` / `M` |
| `Age Group` | Text | 0-17, 18-25, 26-35, 36-45, 46-50, 51-55, 55+ |
| `Age` | Integer | Age in years (12–92) |
| `Marital_Status` | Integer | `0` = Unmarried, `1` = Married *(assumed coding)* |
| `State` | Text | 16 Indian states |
| `Zone` | Text | Central, Southern, Western, Northern, Eastern |
| `Occupation` | Text | 15 occupations |
| `Product_Category` | Text | 18 categories |
| `Orders` | Integer | Units/quantity in the record (1–4) |
| `Amount` | Decimal | Total sale value in ₹ (₹188 – ₹23,952) |
| `Status` | — | **Empty in all rows — dropped** |
| `Avg_order_value` | Decimal | Derived: `Amount / Orders` |

---

## 🧹 Data Quality & Cleaning

| # | Issue found | Evidence | Action |
|---|---|---|---|
| 1 | `Status` column completely empty | 11,239 / 11,239 null | Removed |
| 2 | Exact duplicate rows | 8 rows (₹70,304, 0.07% of revenue) | Removed in Power Query |
| 3 | `Age Group` inconsistent with `Age` | 22 rows for `User_ID 1001941`: Age 26 labelled "36-45" | Rebuilt `Age Group` from `Age` with DAX/Power Query |
| 4 | `User_ID` is not a reliable customer key | 3,752 distinct IDs over 11,239 rows; 502 IDs appear with more than one gender | Treated each row as a sales record; **do not count customers with `DISTINCT(User_ID)`** |
| 5 | Currency stored as text in the processed file | `"₹ 23,952"` | Strip `₹`, `,` and spaces → Decimal number |
| 6 | Misleading KPIs in the original Excel `Analysis` sheet | "Total Customer" = 11,239 (row count); "Average Sales" = ₹55.7M (a *sum* of `Avg_order_value`) | Replaced by correct measures below |

> No other missing values or negative/zero amounts were found.

### Power Query (M) snippet — processed file

```m
let
    Source   = Excel.Workbook(File.Contents("Diwali_processed_Data.xlsx"), null, true),
    Sheet1   = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet1, [PromoteAllScalars=true]),
    NoStatus = Table.RemoveColumns(Promoted, {"Status"}),
    Clean    = Table.TransformColumns(NoStatus, {
                 {"Amount",          each Number.From(Text.Select(_, {"0".."9","."})), type number},
                 {"Avg_order_value", each Number.From(Text.Select(_, {"0".."9","."})), type number}}),
    NoDups   = Table.Distinct(Clean)
in
    NoDups
```

---

## 📈 Power BI Implementation

### DAX Measures

```dax
Total Revenue          = SUM ( Diwali[Amount] )
Total Records          = COUNTROWS ( Diwali )
Total Units            = SUM ( Diwali[Orders] )
Avg Amount per Record  = DIVIDE ( [Total Revenue], [Total Records] )
Avg Value per Unit     = DIVIDE ( [Total Revenue], [Total Units] )
Revenue Share %        = DIVIDE ( [Total Revenue], CALCULATE ( [Total Revenue], ALL ( Diwali ) ) )
Female Revenue %       = DIVIDE ( CALCULATE ( [Total Revenue], Diwali[Gender] = "F" ), [Total Revenue] )
State Rank             = RANKX ( ALL ( Diwali[State] ), [Total Revenue] )

-- Calculated columns
Marital Label = IF ( Diwali[Marital_Status] = 1, "Married", "Unmarried" )
Age Group Fix =
SWITCH ( TRUE (),
    Diwali[Age] <= 17, "0-17",
    Diwali[Age] <= 25, "18-25",
    Diwali[Age] <= 35, "26-35",
    Diwali[Age] <= 45, "36-45",
    Diwali[Age] <= 50, "46-50",
    Diwali[Age] <= 55, "51-55",
    "55+" )
```

### Dashboard Pages

| Page | Visuals |
|---|---|
| **1. Executive Overview** | KPI cards (Revenue, Records, Units, Avg Amount) · Revenue by category · Top states · Gender donut |
| **2. Customer Demographics** | Age group × gender stacked column · Marital status × gender · Revenue by occupation |
| **3. Product & Geography** | Category treemap · State map/bar · Zone × Category matrix |
| **Slicers (all pages)** | Gender · Age Group · Zone · Product Category |

---

## 🔍 Key Findings

| KPI | Value |
|---|---|
| Total revenue | **₹106.25M** |
| Sales records | **11,239** |
| Units sold (`Orders`) | **27,981** |
| Mean / median amount per record | **₹9,454 / ₹8,109** |
| Average value per unit | **₹3,797** |

1. **Women generate 70.0% of revenue** (₹74.3M) — but spend per record is almost identical (₹9,491 vs ₹9,367 for men). The gap comes from *how many* women buy, not *how much* they spend.
2. **Ages 26–45 generate 60.9% of revenue**; ages 26–35 alone contribute 40.1%. Under-18s contribute 2.5%.
3. **Unmarried buyers generate 58.5%**; unmarried women are the single largest segment (41.2%).
4. **Uttar Pradesh, Maharashtra and Karnataka contribute 44.5%** of revenue; the top 5 states contribute 63.1%. The Central zone leads with 39.2%, Eastern trails with 6.6%.
5. **Food is the #1 category (31.9%)** and has a much higher ticket than Clothing (₹13,628 vs ₹6,213 per record) despite fewer records. Top 4 categories = 76.8%.
6. **Occupation:** IT Sector, Healthcare, Aviation and Banking together account for 48.2% of revenue.
7. **Amount is not correlated with age (r = 0.03) or units (r = −0.01)** — segments differ mainly in *volume*, not basket size.

### Sample visuals
| | |
|---|---|
| ![Revenue by category](assets/charts/category.png) | ![Revenue by state](assets/charts/state.png) |
| ![Age & gender](assets/charts/age_gender.png) | ![Amount distribution](assets/charts/amount_hist.png) |

---

## 💡 Recommendations

*These are hypotheses supported by the data, not proven causal effects.*

1. **Target women aged 26–45** with festive campaigns — the core revenue segment.
2. **Prioritise Food and gifting bundles**; upsell high-ticket categories (Footwear, Furniture, Auto) whose per-record value exceeds ₹14,000.
3. **Concentrate regional spend on UP, Maharashtra, Karnataka and Delhi**; test growth campaigns in Eastern/Northern states, which have lower revenue *and* lower ticket size.
4. **Explore corporate/occupational gifting** with IT, Healthcare, Aviation and Banking professionals.
5. **Grow the under-18 / 55+ segments** through family bundles and senior-friendly offers (currently 2.5% and 3.8%).

---

## ⚠️ Limitations & Assumptions

- No date/time column → no trend or seasonality analysis.
- `Marital_Status` coding (1 = Married) and `Orders` meaning (units per record) are **assumed**; they are not documented in the source.
- `User_ID` is inconsistent, so customer-level (repeat-buyer) analysis is unreliable.
- Delhi is classified under the **Central** zone in this dataset; the classification is kept as provided.
- Headline figures use all 11,239 records so they reconcile with the original Excel dashboard. Removing the 8 duplicates gives 11,231 records and ₹106.18M revenue.

---

## ▶️ How to Run

1. Clone the repository.
   ```bash
   git clone https://github.com/<your-username>/diwali-sales-eda-powerbi.git
   ```
2. Open `powerbi/Diwali_Sales_Dashboard.pbix` in **Power BI Desktop**.
3. If prompted, update the data source path: *Home → Transform data → Data source settings*.
4. Click **Refresh** to reload the data.

---

## 🛠 Tools Used

- **Power BI Desktop** — Power Query, data modelling, DAX, visuals
- **Microsoft Excel** — initial pivots and dashboard
- **Python (pandas, matplotlib)** — profiling and validation of figures

---

## 👤 Author

Vishul Kaushik 
