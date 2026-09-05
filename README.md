# Retail Sales Analytics Dashboard using Excel Power Pivot & DAX

## 📊 Project Overview

The **Retail Sales Analytics Dashboard** is an interactive Business Intelligence project developed using **Microsoft Excel, Power Pivot, DAX, PivotTables, PivotCharts, and Slicers**.

The project focuses on transforming raw retail transaction data into meaningful business insights through a structured **Star Schema data model** and analytical DAX measures.

The dashboard provides a centralized view of important business metrics such as **Revenue, Profit, Profit Margin, Orders, Revenue per Customer, Brand Performance, Channel Performance, Store Performance, Category Profitability, Shipping Performance, and Monthly Revenue Trends**.

The solution is designed to help business users quickly identify performance trends, compare business segments, and make data-driven decisions.

---

# 🎯 Business Problem

Retail organizations generate large amounts of transactional data across customers, products, stores, and multiple sales channels. Without an effective analytical system, extracting useful insights from this data can be time-consuming.

This project addresses the following business requirements:

* Monitor overall sales and revenue performance.
* Track total profit and profitability.
* Measure profit margins.
* Analyze customer purchasing behavior.
* Compare performance across sales channels.
* Identify high-performing brands.
* Identify high-performing stores and cities.
* Analyze category-level profitability.
* Evaluate shipping-mode performance.
* Track monthly revenue trends.
* Provide an interactive and user-friendly reporting solution.

---

# 📁 Dataset Information

The project uses a retail sales dataset containing approximately:

* **20,000 Sales Transactions**
* **2,500 Customers**
* Multiple Product Brands
* Multiple Product Categories
* Multiple Store Locations
* Multiple Sales Channels

The dataset is organized into separate fact and dimension tables to support efficient analytical processing.

---

# 🗂️ Data Tables

## 1. Sales_Fact

The `Sales_Fact` table is the central **transaction/fact table** containing individual sales records.

### Key Fields

* `OrderID`
* `OrderDateKey`
* `CustomerID`
* `ProductID`
* `StoreID`
* `Quantity`
* `UnitPrice`
* `DiscountPct`
* `GrossSales`
* `NetRevenue`
* `COGS`
* `Profit`

This table provides the numerical measures required for revenue, profit, order, and sales analysis.

---

## 2. Customers

The `Customers` table contains customer-level information.

### Key Fields

* `CustomerID`
* `CustomerName`
* `Gender`
* `Age`
* `Segment`
* `City`
* `State`
* `PreferredChannel`
* `IsActive`

This table supports customer segmentation and customer-level analysis.

---

## 3. Products

The `Products` table contains product-related information.

### Key Fields

* `ProductID`
* `ProductName`
* `Brand`
* `Category`

It enables analysis of revenue and profitability by product, brand, and category.

---

## 4. Stores

The `Stores` table contains information about stores and their associated channels.

### Key Fields

* `StoreID`
* `StoreName`
* `City`
* `Channel`

This table is used for geographical and store/channel performance analysis.

---

## 5. Date

The `Date` table is the dedicated **Date Dimension** used for time-based analysis.

### Key Fields

* `Date`
* `Month`
* `Quarter`
* `Year`

It enables monthly, quarterly, and yearly trend analysis.

---

# ⭐ Data Model

A **Star Schema** was implemented using Excel Power Pivot.

### Fact Table

```text
Sales_Fact
```

### Dimension Tables

```text
Customers
Products
Stores
Date
```

### Model Structure

```text
                    Customers
                        │
                        │
                        ▼
Products ─────────► Sales_Fact ◄───────── Stores
                        │
                        │
                        ▼
                       Date
```

The `Sales_Fact` table acts as the central transaction table, while the dimension tables provide descriptive attributes for filtering and analysis.

### Benefits of the Data Model

* Structured analytical architecture
* Efficient filtering
* Improved DAX calculations
* Reduced data duplication
* Better scalability
* Easier maintenance
* Simplified reporting
* Multi-dimensional analysis

---

# 🧮 DAX Measures

DAX (**Data Analysis Expressions**) was used to create reusable business calculations.

## Revenue Measures

### Total Revenue

```DAX
Total Revenue =
SUM(Sales_Fact[NetRevenue])
```

Calculates the total revenue generated from all sales transactions.

### Online Purchase Revenue

```DAX
Online Purchase =
CALCULATE(
    [Total Revenue],
    Customers[PreferredChannel]="Online"
)
```

Calculates revenue associated with customers whose preferred channel is Online.

### Retail Store Revenue

```DAX
Retail Stores Revenue =
CALCULATE(
    [Total Revenue],
    Stores[Channel]="Retail"
)
```

Calculates revenue generated through retail stores.

---

## Order Measures

### Total Orders

```DAX
Total Orders =
DISTINCTCOUNT(Sales_Fact[OrderID])
```

Counts unique orders in the sales dataset.

### Average Order Value

```DAX
Average Order Value =
DIVIDE(
    [Total Revenue],
    [Total Orders]
)
```

Measures the average revenue generated per order.

---

## Customer Measures

### Total Customers

```DAX
Total Customers =
DISTINCTCOUNT(Customers[CustomerID])
```

Calculates the number of unique customers.

### Revenue Per Customer

```DAX
Revenue Per Customer =
DIVIDE(
    [Total Revenue],
    [Total Customers]
)
```

Measures the average revenue generated per customer.

---

## Profit Measures

### Total Profit

```DAX
Total Profit =
SUM(Sales_Fact[Profit])
```

Calculates total profit generated from sales.

### Profit Margin

```DAX
Profit Margin =
DIVIDE(
    [Total Profit],
    [Total Revenue]
)
```

Measures profitability as a percentage of revenue.

---

## Brand Analysis

### Pulse Brand Revenue

```DAX
Pulse Brand Revenue =
CALCULATE(
    [Total Revenue],
    Products[Brand]="Pulse"
)
```

Calculates revenue specifically generated by the Pulse brand.

---

# 📌 Dashboard KPIs

The dashboard provides a set of high-level KPIs for quick business monitoring.

| KPI                         | Business Purpose                               |
| --------------------------- | ---------------------------------------------- |
| **Total Revenue**           | Measures overall sales revenue                 |
| **Total Profit**            | Measures total business profitability          |
| **Profit Margin**           | Measures profitability relative to revenue     |
| **Total Orders**            | Measures order volume                          |
| **Revenue Per Customer**    | Measures average customer revenue contribution |
| **Online Purchase Revenue** | Measures online channel contribution           |
| **Retail Revenue**          | Measures retail channel contribution           |
| **Brand Revenue**           | Measures revenue generated by brands           |

These KPI cards provide management with an immediate overview of business performance.

---

# 📈 Dashboard Visualizations

## 1. Revenue by Brands

The **Revenue by Brands** visualization compares revenue generated by individual brands.

It helps identify:

* Leading brands
* Revenue contribution
* Brand-level performance
* Potential growth opportunities

---

## 2. Channel-wise Revenue Analysis

This visualization compares revenue performance across different sales channels, including:

* Retail
* Online Hub
* B2B Desk

The analysis helps understand the contribution of different channels and supports channel-specific business strategies.

---

## 3. Top Performing Stores by City

This visualization identifies the stores and cities contributing the highest revenue.

It helps businesses understand:

* Regional performance
* Store productivity
* High-performing markets
* Potential expansion opportunities

---

## 4. Profit by Category

The **Profit by Category** visualization compares profitability across product categories.

It helps identify:

* Most profitable categories
* Lower-performing categories
* Category contribution to total profit
* Opportunities for product strategy optimization

---

## 5. Profit Margin by Shipping Mode

This chart compares profit margins across different shipping methods.

It helps evaluate the relationship between logistics choices and profitability.

---

## 6. Monthly Revenue Trend

The monthly revenue trend tracks revenue performance over time.

It helps identify:

* Growth patterns
* Revenue fluctuations
* Seasonal trends
* High-performing periods
* Low-performing periods

This information can support sales forecasting, inventory planning, and promotional strategies.

---

# 🎛️ Interactive Dashboard Features

The dashboard is designed as an interactive analytical reporting system.

## Slicers

Users can dynamically filter the dashboard using:

* **Channel**
* **Month**
* **Brand**

When a slicer selection is made, the connected PivotTables, PivotCharts, and KPI calculations update according to the selected context.

This allows users to perform both high-level and detailed analysis without manually modifying the underlying dataset.

---

# 💡 Key Business Insights

The dashboard provides several useful business insights:

* Approximately **20,000 transactions** are included in the analysis.
* The business generated approximately **₹15.96 million in revenue** based on the underlying dataset values.
* Overall **profit margin is approximately 22.15%**.
* **Electronics** is one of the strongest profit-generating categories.
* The **Retail channel** contributes significant revenue compared with individual online channels.
* Revenue performance varies across cities and stores.
* Certain locations contribute substantially more revenue than others.
* Monthly revenue analysis highlights fluctuations and high-sales periods.
* Brand-level analysis helps identify high-performing brands and potential growth opportunities.

---

# 🛠️ Tools & Technologies

The project was developed using:

* **Microsoft Excel**
* **Power Query**
* **Power Pivot**
* **DAX**
* **PivotTables**
* **PivotCharts**
* **Slicers**
* **Data Modeling**
* **Star Schema**

---

# 🧠 Skills Demonstrated

This project demonstrates practical experience in:

* Data Cleaning
* Data Transformation
* Data Modeling
* Star Schema Design
* Power Pivot
* DAX Calculations
* KPI Development
* Business Intelligence
* Sales Analytics
* Customer Analysis
* Profitability Analysis
* Data Visualization
* Dashboard Design
* Interactive Reporting
* Business Insight Generation
* Excel Reporting

---

# 📊 Business Value

The dashboard delivers business value in four major areas:

### 1. Performance Tracking

Provides a centralized view of revenue, profit, orders, and other important KPIs.

### 2. Better Decision-Making

Converts raw transactional data into actionable business insights.

### 3. Revenue & Profit Analysis

Helps identify high-performing brands, categories, channels, stores, and locations.

### 4. Strategic Planning

Supports sales planning, marketing decisions, inventory planning, channel strategy, and overall business growth.

---

# 📸 Dashboard Preview

Add your dashboard screenshot here:


![Retail Sales Analytics Dashboard](images/TASK13.png)

```

The screenshot should showcase the complete dashboard, including KPI cards, charts, slicers, and overall dashboard layout.

---

# 🎥 Demo Video

A project demonstration video can be added here:

```markdown
[▶ Watch Retail Sales Analytics Dashboard Demo](YOUR_VIDEO_LINK)
```

The demonstration should cover:

1. Dashboard introduction
2. KPI overview
3. Brand revenue analysis
4. Channel analysis
5. Store/city performance
6. Category profitability
7. Shipping-mode analysis
8. Monthly revenue trend
9. Interactive slicers
10. Key business insights

---

# 📂 Workbook Structure

| Worksheet       | Purpose                                 |
| --------------- | --------------------------------------- |
| **Customers**   | Customer information                    |
| **Products**    | Product, brand and category information |
| **Stores**      | Store, city and channel information     |
| **Sales_Fact**  | Central transaction/fact table          |
| **Date**        | Date dimension for time analysis        |
| **PIVOT Table** | Analytical PivotTables                  |
| **DASHBOARD**   | Final interactive dashboard             |

---

# ▶️ How to Use

### Step 1

Download or open the Excel workbook.

### Step 2

Open the `.xlsm` file in Microsoft Excel.

### Step 3

Enable macros/content if Excel displays a security notification.

### Step 4

Navigate to the **DASHBOARD** worksheet.

### Step 5

Review the KPI cards and visualizations.

### Step 6

Use the available slicers to filter the analysis by:

* Channel
* Month
* Brand

### Step 7

Observe how the dashboard responds to different selections.

### Step 8

Clear the filters to return to the overall business view.

---

# 🚀 Future Improvements

The project can be further enhanced by implementing:

* Automated data refresh
* Sales forecasting
* Revenue prediction
* Customer segmentation
* Customer Lifetime Value analysis
* Product-level recommendations
* Return-rate analysis
* Inventory analytics
* Advanced geographic analysis
* Customer retention analysis
* AI-based business recommendations
* Automated management reporting

---

# 🏆 Project Outcome

The project successfully demonstrates the complete journey from **raw retail transaction data to an interactive Business Intelligence dashboard**.

The combination of **Power Query, Power Pivot, DAX, Star Schema modeling, PivotTables, PivotCharts, and Slicers** enables the dashboard to transform complex transactional information into clear and actionable business insights.

The final solution provides decision-makers with a consolidated view of **sales performance, profitability, customer behavior, brand performance, channel performance, store performance, geographic trends, shipping profitability, and monthly revenue patterns**.

---


## ⭐ Project Summary

> **Raw Data → Power Query → Data Model → Power Pivot → DAX → PivotTables → Dashboard → Business Insights**

This project demonstrates how **Microsoft Excel can be used as a powerful Business Intelligence and Data Analytics platform** for transforming transactional retail data into an interactive decision-support solution.
