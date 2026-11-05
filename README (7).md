<div align="center">

# 📊 SUPERSTORE SALES ANALYSIS DASHBOARD

### ✦ Sales Performance · Profitability Analysis · Interactive Excel Reporting ✦

**An Academic Project in Microsoft Excel**

<br>

![Project](https://img.shields.io/badge/PROJECT-SUPERSTORE%20DASHBOARD-1d4ed8?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Status](https://img.shields.io/badge/STATUS-FINAL%20PROJECT-16a34a?style=for-the-badge)
![Focus](https://img.shields.io/badge/FOCUS-SALES%20ANALYTICS-7c3aed?style=for-the-badge)

<br>

**🎓 B.TECH — Final Year** &nbsp; • &nbsp; **📚 Data Analysis Project**  
**👨‍🏫 Faculty Guide: Prof. Girish Gondaliya**

<br>

---

**📘 Documentation** · **🧹 Data Cleaning** · **🧮 Excel Formulas** · **📈 PivotCharts** · **🎛️ Interactive Filters**

</div>

> [!IMPORTANT]
> **Evaluator quick access:** Open `Final_ExcelDashboard_FINAL_Designed.xlsx`, select the **5.Dashboard** worksheet, and use the **Region slicer** and **Order Date timeline** to explore the charts. The workbook should be submitted alongside this README.

## 🧭 Table of Contents

- [✨ Project Overview](#-project-overview)
- [🎯 Project Objectives](#-project-objectives)
- [🗂️ Workbook Architecture](#️-workbook-architecture)
- [🔍 Worksheet-by-Worksheet Breakdown](#-worksheet-by-worksheet-breakdown)
- [🧠 Excel Formula Toolkit](#-excel-formula-toolkit)
- [📊 Dashboard Visualizations](#-dashboard-visualizations)
- [🧪 Practical Demonstrations](#-practical-demonstrations)
- [🚀 How to Review](#-how-to-review)
- [✅ Submission Checklist](#-submission-checklist)
- [🏆 Learning Outcomes](#-learning-outcomes)
- [👩‍🎓 Academic Details](#-academic-details)

---

## ✨ Project Overview

**Superstore Sales Analysis Dashboard** is a Microsoft Excel project that turns sales transaction data into clear business insights using organized worksheets, lookup formulas, PivotTables, charts, conditional formatting and interactive filters.

The dashboard helps compare **regional sales**, **category-wise profit**, **customer segment sales**, and **yearly sales trends** for the **2014–2017** period.

### 📌 Project Snapshot

| 📍 Metric | 📊 Project Scope |
|:--|:--|
| 🗂️ Total worksheets | **6** |
| 🧾 Source transactions | **9,994** |
| 🗓️ Analysis period | **2014–2017** |
| 📊 PivotTables | **4** |
| 📈 Charts | **4** |
| 🎛️ Interactive controls | **Region slicer and Order Date timeline** |
| 🧩 Main capabilities | Data preparation, lookup formulas, PivotTables, visual analysis |
| 📁 Workbook | `Final_ExcelDashboard_FINAL_Designed.xlsx` |

### 🔄 Analysis Workflow

```text
  ORIGINAL DATA           PREPARATION & ANALYSIS          DASHBOARD
       │                           │                         │
       ▼                           ▼                         ▼
  Raw Sales Data  ──►  Clean Data + Excel Table  ──►  PivotTables
                                            │                 │
                                            └──► Formulas     ▼
                                                    Charts + Slicer + Timeline
```

---

## 🎯 Project Objectives

- 🧹 **Organize and prepare data** for reliable analysis.
- 🔎 **Apply lookup and logical formulas** to retrieve and classify information.
- 📊 **Summarize transactions** using Excel PivotTables.
- 📈 **Visualize key measures** with readable charts.
- 🎛️ **Explore results interactively** with a Region slicer and Order Date timeline.
- 📝 **Document the workbook** so another user can understand and review it.

---

## 🗂️ Workbook Architecture

| # | 📑 Worksheet | 🎯 Purpose |
|:--:|:--|:--|
| 01 | **1.Raw Data** | Original Superstore transaction data |
| 02 | **2.Clean Data** | Prepared data used for analysis; includes the `tblSuperstore` table |
| 03 | **3.Pivot_Tables** | Summary tables that support the charts |
| 04 | **4.Advanced_Formulas** | Demonstrations of lookup and conditional formulas |
| 05 | **5.Dashboard** | Interactive sales and profit dashboard |
| 06 | **6.Documentation** | Workbook notes and usage guidance |

---

## 🔍 Worksheet-by-Worksheet Breakdown

### 01 · 🗃️ Raw Data — Original Dataset

Contains the source Superstore records, including order information, customers, products, locations, sales and profit. This worksheet is kept separately so the original data can be compared with the prepared dataset.

### 02 · 🧹 Clean Data — Analysis-Ready Table

Contains the prepared transaction data for calculations and summaries. The structured Excel table is named **`tblSuperstore`**.

**Review focus:** Check field headers, data types, missing values and duplicate handling before interpreting results. The existence of a clean-data sheet does not by itself prove that every possible data-quality issue was removed.

### 03 · 📋 Pivot_Tables — Summary Analysis

The workbook contains **four PivotTables** supporting the dashboard's four main comparisons: region, category, customer segment and year.

**Key idea:** PivotTables group detailed transactions into readable summaries without manually totaling each group.

### 04 · 🧮 Advanced_Formulas — Formula Demonstrations

This sheet includes examples of **`VLOOKUP`**, **`INDEX` + `MATCH`**, and **`IF`**.

| 🔎 Demonstration | 💡 Purpose |
|:--|:--|
| Product lookup | Retrieve product details from an identifier |
| Customer lookup | Match a customer ID to the corresponding customer name |
| Profit classification | Display a status based on a profit value |

**Example of an actual customer lookup pattern used in the workbook:**

```excel
=INDEX('2.Clean Data'!$G$2:$G$9995,MATCH(E2,'2.Clean Data'!$F$2:$F$9995,0))
```

**Formula explanation:** `MATCH` finds the customer ID position, and `INDEX` returns the customer name at that position.

### 05 · 📊 Dashboard — Sales Performance Overview

The dashboard contains a main title, analysis subtitle, KPI display areas, four charts and interactive filtering controls.

**KPI headings:** Total Sales, Total Profit, Profit Margin and Data Records.

> [!NOTE]
> The dashboard notes specify that the displayed KPI figures represent the **full dataset**, while the connected charts respond to slicer and timeline selections. Do not describe these KPI values as dynamically filter-linked without separately verifying that behavior.

### 06 · 📘 Documentation — Workbook Instructions

Provides supporting notes so the workbook can be reviewed, refreshed and presented in a consistent order.

---

## 🧠 Excel Formula Toolkit

| 🧩 Function / Feature | ⚙️ Application |
|:--|:--|
| `VLOOKUP` | Retrieve information using a matching key |
| `INDEX` | Return a value from a selected position |
| `MATCH` | Find the position of an exact match |
| `IF` | Apply a logical condition and return a status |
| PivotTables | Summarize large datasets by category |
| PivotCharts | Visualize PivotTable results |
| Conditional formatting | Highlight values or statuses based on rules |
| Slicer | Filter connected summaries by region |
| Timeline | Filter connected summaries by order date |

### 💻 Selected Formula Patterns

```excel
=VLOOKUP(lookup_value,table_array,column_index,FALSE)
=INDEX(return_range,MATCH(lookup_value,lookup_range,0))
=IF(profit_cell>0,"Profit",IF(profit_cell<0,"Loss","Break Even"))
```

*These are explanatory patterns. The exact workbook formulas should be reviewed directly in the Excel Formula Bar; not every sample pattern above is a verbatim cell formula.*

---

## 📊 Dashboard Visualizations

| 📈 Chart | 📌 Business Question |
|:--|:--|
| **01 · Regional Sales Performance** | Which region contributes the most sales? |
| **02 · Category Profit Analysis** | Which product category contributes the most profit? |
| **03 · Customer Segment Sales** | How do sales differ across customer segments? |
| **04 · Yearly Sales Trend** | How have sales changed across 2014–2017? |

The **Region slicer** and **Order Date timeline** are intended to make the connected PivotCharts interactive.

---

## 🧪 Practical Demonstrations

| 🎯 Demo | 🖱️ Action | 👀 Expected Observation |
|:--|:--|:--|
| Region analysis | Select one region in the slicer | Connected charts show that region's results |
| Date analysis | Choose a year or date range in the timeline | Connected charts reflect the selected period |
| Reset filters | Clear slicer and timeline filters | Charts return to the full available dataset |
| Customer lookup | Change the lookup customer ID | Matching customer name is returned |
| Profit status | Inspect an `IF` formula cell | Status follows the formula's profit condition |
| Refresh analysis | Select **Data → Refresh All** | Pivot-based summaries refresh from their source |

> [!TIP]
> For a viva demonstration, first show the unfiltered dashboard, then filter by region and date, clear the filters, and finally explain one lookup formula in the Formula Bar.

---

## 🚀 How to Review

1. 📂 Open `Final_ExcelDashboard_FINAL_Designed.xlsx` in **Microsoft Excel Desktop**.
2. 🏠 Start with **5.Dashboard** to see the project overview.
3. 📑 Review **1.Raw Data** and **2.Clean Data**.
4. 📊 Open **3.Pivot_Tables** and inspect the summaries.
5. 🧮 Open **4.Advanced_Formulas** and select formula cells to inspect the Formula Bar.
6. 🎛️ Return to the dashboard and test the Region slicer and Order Date timeline.
7. 🔄 Use **Data → Refresh All** if the PivotTable summaries need updating.
8. 📘 Review **6.Documentation** for workbook guidance.

**💻 Compatibility:** Microsoft Excel Desktop is recommended for the best support for PivotTables, PivotCharts, slicers and timelines. Other spreadsheet editors may not preserve all interactive features.

---

## ✅ Submission Checklist

- [x] Six worksheets documented
- [x] Dataset scope and analysis period included
- [x] Four PivotTables and four charts documented
- [x] Main lookup and logical formulas explained
- [x] Region slicer and Order Date timeline described
- [x] Workbook usage and viva steps included
- [ ] Confirm the final workbook opens correctly in Excel Desktop
- [ ] Recheck all chart filters and refresh behavior
- [ ] Add or confirm student details before submission
- [ ] Submit the **Excel workbook + README.md** together

---

## 🏆 Learning Outcomes

Through this project, I learned how to organize transaction data, prepare a structured Excel table, use lookup and logical formulas, build PivotTable summaries and present sales performance through charts.

I also practiced using slicers and timelines to explore business data and explaining my analysis in a clear, step-by-step presentation.

---

## 👩‍🎓 Academic Details

| 🏷️ Field | 📋 Information |
|:--|:--|
| **Student Name** | Sumit Kaswala |
| **Academic Program** | Bachelor of Technology (B.Tech) – Final Year |
| **Course Focus** | Data Analysis |
| **Project Title** | Superstore Sales Analysis Dashboard |
| **Faculty Guide** | Prof. Girish Gondaliya |
| **Tools** | Microsoft Excel · Markdown |
| **Deliverables** | Excel Workbook · README Documentation |

---

<div align="center">

### 🌟 Thank You for Reviewing My Project!

**Superstore Sales Analysis Dashboard**  
*Microsoft Excel · Sales Analysis · Interactive Reporting*

[⬆️ Back to Top](#-superstore-sales-analysis-dashboard)

</div>
