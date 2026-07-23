# Regional & Salesperson Revenue Analysis 2023

## Project Overview 
Interactive Excel dashboard analyzing regional and salesperson sales performance to support data-driven revenue and territory decisions.

---

## Table of Contents
1. [Business Problem](#business-problem)
2. [Project Objective](#project-objective)
3. [Dataset](#dataset)
4. [Tools Used](#tools-used)
5. [Workflow](#workflow)
6. [Data Analysis (Advanced Formulas)](#data-analysis-advanced-formulas)
7. [Data Analysis (Pivot Tables)](#data-analysis-pivot-tables)
8. [Dashboard](#dashboard)
9. [Key Insights](#key-insights)
10. [Recommendations](#recommendations)
11. [Repository Structure](#repository-structure)
12. [Limitations & Next Steps](#limitations--next-steps)

---
 
## Business Problem
Sales leadership had no consolidated view of **which regions and salespeople were driving revenue** and which were underperforming. Decisions on territory support, incentive design, and resource allocation were being made without a clear, data-backed picture of performance trends.

---

## Project Objectives 
Build an Excel dashboard that lets stakeholders:
- Track monthly sales trends over time
- Compare performance across regions
- Identify top and bottom-performing salespeople
- Filter dynamically by region, product category, and customer segment

---

## ## Dataset Information
- **Source:** Synthetic dataset generated for portfolio purposes. Does not represent real individuals or companies.
- **Size:** 501 rows (500 orders + header), 10 columns
- **Columns:** `Order_ID`, `Order_Date`, `Region`, `Sales_person`, `Product_Category`, `Product`, `Unit_Price($)`, `Quantity`, `Sales_Amount($)`, `Customer_Segment`
- **Type:** Sales transaction data
> **Privacy note:** The full workbook is not shared publicly. Sample screenshots are included below to demonstrate the analysis workflow.

**Data set from row 1 to 35**
Screenshot:
![1_to_35_](Data/screenshot/dataset.png)

---

## Tools Used
- Excel
- Excel Formulas & Functions (XLOOKUP, INDEX/MATCH, SUMIFS, COUNTIFS)
- PivotTables & Pivot Charts
- Slicers (interactive filtering)

---

## Workflow
This project follows a standard analytics lifecycle:
 
```
Raw Data → Data Validation → Formula-Based Analysis → Pivot Table Summarization → Dashboard & Visualization → Insights → Recommendations
```
---


---



---

## Data Analysis (Advanced Formulas)
The dataset was used to answer specific business questions using advanced Excel formulas:
 
| Business Question | Technique |
|---|---|
| Find product & quantity for a given Order ID | `XLOOKUP` + `TEXTJOIN` |
| Find salesperson for a given Order ID | `INDEX` + `MATCH` |
| Total sales for Electronics in West region, Q2 2023 | `SUMIFS` with date range |
| Orders by Enterprise customers with quantity > 5 | `COUNTIFS` |
| Total & average sales amount | `SUM`, `AVERAGE` |
| Total number of orders | `COUNTA` |
| Best-selling product | `INDEX` + `MATCH` + `MAX` + `SUMIF` |
| Top-performing region | `INDEX` + `MATCH` + `MAX` + `SUMIF` |
 
**Example — Total Sales for Electronics in West Region, Q2 2023:**
```excel
=SUMIFS(data1!I2:I501,data1!E2:E501,"Electronics",data1!C2:C501,"West",data1!B2:B501,">="&DATE(2023,4,1),data1!B2:B501,"<="&DATE(2023,6,30))
```
`SUMIFS` Found rows where Category = Electronics, Region = West, Date between 1-Apr-2023 and 30-Jun-2023, then sums corresponding I column values.

---

**Advanced Excel Formula Analysis**
Screenshot:
![kpi_Calculation](Analysis/screenshots/Formulas.png)

---

### Data Analysis using Pivot Tables

#### Monthly Sales Trend

This Pivot Table provides a clear and structured view of sales performance over time by summarizing total sales on a monthly basis. 

Analysis Screenshot:
![Analysis](Analysis/screenshots/table_1.png)

---

#### Sales by Region

This Pivot Table presents a regional breakdown of sales, helping to evaluate performance across different geographic areas.

Analysis Screenshot:
![Analysis](Analysis/screenshots/table_2.png)

---

#### Indivisual Sales

This Pivot Table provides a detailed analysis of sales performance at the individual salesperson level. It helps evaluate each salesperson’s contribution to overall revenue and supports performance-based decision-making.

Analysis Screenshot:
![Analysis](Analysis/screenshots/table_3.png)

---

### Insights Generated

- Sales show a consistent upward trend from January to August, indicating steady business growth.
- The North region contributes the highest share of total sales, making it the top-performing region, suggesting opportunities to balance growth across other regions.
- One salesperson leads with the highest sales, while others have room to improve and boost their performance.

---

## 📊 Dashboard Creation: An excel dashboard was created to summarize insights visually.

The dashboard includes:

- KPIs

- Pivot Charts

- Slicers

---

### Key Performance Indicators (KPIs): The following KPIs were displayed at the top of the dashboard to summarize overall performance:

- Total Orders

- Total Sales

- Average SalesAmount per Order

- Product with the highest Total SalesAmount 

- Region with the highest total SalesAmount

---

### Pivot Charts

Pivot Charts provide a visual representation of Pivot Table data, making it easier to analyze patterns, compare performance, and identify trends. They help transform complex data into clear, interactive insights for better decision-making. the following pivot charts are used

- Monthly Sales Trend

- Sales by Region

- Indivisual Sales

---

### Slicers

Slicers are interactive filters that make analyzing Pivot Table data easy and intuitive. They allow you to quickly filter by categories like Region, Product Category, or Cuustomer Segment, helping you focus on specific insights without altering the underlying data. the following slicers are used

- Region

- Product Category

- Customer Segment

---

Dashboard Screenshot:
<img width="1280" height="769" alt="dashboard" src="https://github.com/user-attachments/assets/657f8da7-8089-46b8-a4ed-ef5949ea0a0e" />

---

```text
START ─────────────────────────────────────────────

HR Analytics Dashboard Project/
│
├── Data/
│   └── screenshot/
│       └── dataset.png
│
├── Analysis/
│   └── screenshots/
│       ├── Formulas.png
│       ├── table_1.png
│       ├── table_2.png
│       └── table_3.png
│
├── Dashboard/
│   └── screenshot/
│       └── dashboard.png
│
└── README.md

END ─────────────────────────────────────────────
```

#### Calculate Total Sales for Electronics in the West Region during Q2 2023 (April-June)

Calculated the Total Sales for Electronics in the West Region during Q2 2023 (April-June) using : 

**Formula:**
```excel
=SUMIFS(data1!I2:I501,data1!E2:E501,"Electronics",data1!C2:C501,"West",data1!B2:B501,">="&DATE(2023,4,1),data1!B2:B501,"<="&DATE(2023,6,30))
```
SUMIFS Found rows where Category = Electronics, Region = West, Date between 1-Apr-2023 and 30-Jun-2023, then sums corresponding I column values.
