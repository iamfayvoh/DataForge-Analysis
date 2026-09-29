# 📊 E-Commerce Sales & Profitability Analysis (Excel Data Project)

![Excel](https://img.shields.io/badge/Tool-Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Status](https://img.shields.io/badge/Project_Status-Completed-success?style=for-the-badge)

## 📌 Project Overview
This project provides an end-to-end sales performance and profitability audit for an e-commerce platform. Using Microsoft Excel, raw transaction data was audited for quality, transformed into analytical tables, evaluated through pivot tables, and visualized via an executive dashboard. 

The primary objective is to **identify profit leaks**, uncover performance drivers across product categories and payment methods, and offer data-backed strategic recommendations to eliminate revenue friction.

---

## 🖼️ Executive Dashboard Preview

<!-- INSERT YOUR DASHBOARD IMAGE BELOW -->
![E-Commerce Sales & Profitability Dashboard](YOUR_DASHBOARD_IMAGE_URL_HERE)
*Figure 1.1: Interactive E-Commerce Sales & Profitability Dashboard built in Microsoft Excel.*

---

## 🔑 Key Performance Indicators (KPIs)

* **Total Revenue:** ₹437.8K (₹437,771)
* **Net Profit:** ₹37.0K (₹36,963)
* **Average Profit Margin:** 8.44%
* **Total Units Sold:** 5,615 Units
* **Total Transactions:** 1,500 Orders
* **Average Order Value (AOV):** ₹875.54

---

## 📁 Workbook Architecture & Methodology

The analysis was performed across a structured multi-sheet Excel workbook (`Data_Forge_Sales_Analysis_Assignment.xlsx`):

| Sheet Name | Description & Purpose |
| :--- | :--- |
| `Dashboard` | Executive summary board with dynamic slicers and visual KPIs. |
| `Summary` | Detailed Profit Leak matrix highlighting negative profit sub-categories and payment channel failures. |
| `Calculations` | Central repository for baseline aggregation metrics (SUMIFS, COUNTIFS, Rank formulas, AOV). |
| `PT_Charts` | Pivot table source configurations powering dashboard charts and visual layouts. |
| `Pivot Tables` | Analytical pivot views evaluating sub-category margins, transaction volume, and payment modes. |
| `Details_cleaned_work` | Working sheet for data transformation, field formatting, and calculations. |
| `Details_cleaned` | Raw master dataset (1,500 clean orders). |
| `Data_Quality_Report` | Audit trail documenting dataset verification and structural setup. |

---

## 🚨 Key Insights & Profit Leak Analysis

<!-- INSERT YOUR PROFIT LEAKS TABLE IMAGE HERE -->
![Profit Leaks Analysis](YOUR_PROFIT_LEAKS_IMAGE_URL_HERE)
*Figure 1.2: Profit Leak Breakdown by Sub-Category and Payment Method.*

### 1. Loss-Making Sub-Categories (Direct Margin Leaks)
Out of 17 sub-categories, 5 consistently operate at a net loss, dragging overall profit margins down:

* **Furnishings:** -₹806 Net Profit (-5.98% Margin) | 34 Loss Orders
* **Electronic Games:** -₹644 Net Profit (-1.64% Margin) | 43 Loss Orders
* **Kurti:** -₹401 Net Profit (-11.93% Margin) | 23 Loss Orders
* **Skirt:** -₹315 Net Profit (-16.19% Margin) | 29 Loss Orders
* **Leggings:** -₹130 Net Profit (-6.17% Margin) | 23 Loss Orders

### 2. Payment Channel Vulnerabilities
While all payment methods are net profitable overall, transaction-level losses reveal significant operational leaks:

* **Cash on Delivery (COD):** Generates the largest absolute loss pool (**-₹15,483** across 244 loss orders).
* **UPI:** Lowest overall profit margin (**4.79%**) with the highest failure/loss rate (**37.16%** of orders losing money).

### 3. Top Profit Drivers
* **Printers:** #1 Profit Generator (**₹8,606** | 14.52% Margin).
* **Bookcases:** #2 Profit Generator (**₹6,516** | 11.46% Margin).
* **Electronics Category:** Highest overall profit contribution (**₹13.2K**) on low relative volume (1,154 units sold).

---

## 💡 Strategic Business Recommendations

| Objective | Focus Area | Action Plan | Impact |
| :--- | :--- | :--- | :--- |
| **Plug Leaks** | Loss Sub-Categories | Audit supplier costs and increase prices for Furnishings, Electronic Games, and apparel items (Kurti, Skirt, Leggings). Discontinue if unit economics remain negative. | Immediately stops negative profit drag across inventory. |
| **Plug Leaks** | High-Volume Item Losses | Implement minimum price floors and restrict promo code stacking on Sarees, Phones, and Bookcases. | Recovers up to ~₹15.5K in order-level losses. |
| **Plug Leaks** | Payment Friction | Introduce a nominal COD surcharge or minimum threshold to offset logistics and return costs. | Mitigates shipping and handling risk on COD transactions. |
| **Drive Growth** | High-Margin Categories | Expand inventory and ad spend on Printers, Bookcases, and high-margin Electronics. | Maximizes Return on Ad Spend (ROAS). |
| **Drive Growth** | Payment Optimization | Offer checkout discounts for Credit Card payments (highest net margin at 14.51%). | Shifts customer checkout mix away from high-loss channels. |

---

## 💻 Technical Excel Formulas Used

* **Multi-Criteria Counting (COUNTIFS):**
  ```excel
  =COUNTIFS(Details_cleaned!G:G, "COD", Details_cleaned!C:C, "<0")
