# Nova Gear Online Sales Analysis (2021–2025)

An end-to-end Excel analytics project covering data cleaning, exploratory analysis, and interactive dashboard design for a multi-country e-commerce business.

<img width="1213" height="675" alt="Nova Dashboard" src="https://github.com/user-attachments/assets/e46f73a1-4c55-4a02-9d34-7a9236cbe805" />



---

## Table of Contents

- [Dashboard Preview](#dashboard-preview)
- [Project Overview](#project-overview)
- [Business Problem](#business-problem)
- [Project Objectives](#project-objectives)
- [Dataset Information](#dataset-information)
- [Tools & Technologies Used](#tools--technologies-used)
- [Data Cleaning Process](#data-cleaning-process)
- [New Calculated Columns](#new-calculated-columns)
- [Data Analysis Process](#data-analysis-process)
- [Pivot Table Analysis](#pivot-table-analysis)
- [Dashboard Development](#dashboard-development)
- [Key Insights](#key-insights)
- [Business Recommendations](#business-recommendations)
- [Challenges Encountered](#challenges-encountered)
- [Conclusion](#conclusion)
- [Author Information](#author-information)

---

## Dashboard Preview

The image above shows the final one-page interactive dashboard built in Excel. It combines KPI cards, trend and comparison charts, and three slicers (Year, Country, and Product Category) that filter every visual on the page simultaneously, giving a sales director a complete performance snapshot in under five minutes.

---

## Project Overview

Nova Gear is an online consumer-electronics retailer operating across eight countries - USA, Canada, Germany, France, Australia, China, Nigeria, and Kenya. This project analyzes five years of order history (2021–2025) exported from the company's web platform to uncover how the business has evolved, how customers behave across different markets, and where revenue is being lost.

The raw export arrived in a messy, unaudited state. This project walks through the full analytics lifecycle: profiling and cleaning the data, building calculated fields, running PivotTable-driven analysis, and packaging the findings into a single interactive dashboard.

---

## Business Problem

Management suspected the business had shifted materially since launch but had no quantified view of how. Three core questions needed answers:

1. **What changed?** Where is the business growing or shrinking, and since when?
2. **Who are our customers?** Do the eight markets behave like one global audience, or eight distinct ones?
3. **Where are we losing money?** Cancellations and refunds were suspected to be eroding revenue, but the scale was unknown.

---

## Project Objectives

- Audit and clean a 10,000 row raw sales export, documenting every data quality issue found and the action taken.
- Engineer calculated fields (Revenue, Year, Month) to support time-based and financial analysis.
- Use PivotTables to answer the three management questions with at least eight numbers-backed findings.
- Design a single-page interactive dashboard with KPI cards, charts, and slicers suitable for an executive audience.
- Translate findings into concrete, actionable business recommendations.

---

## Dataset Information

- **Source file:** `Nova_Gear_Online_Sales.xlsx` - a flat export from the company's web platform.
- **Original size:** 10,000 records
- **Final cleaned size:** 9,760 records
- **Date range:** 2021 - 2025
- **Currency:** All amounts in USD
- **Raw Dataset**
- <img width="1587" height="717" alt="Nova_raw_data" src="https://github.com/user-attachments/assets/130a82f4-48e8-42e6-809b-d9a98f112901" />
- **Clean Dataset**
- <img width="1902" height="716" alt="Nova_clean_data" src="https://github.com/user-attachments/assets/1dd2e117-2c5b-4d1c-986a-e8965c9151c5" />



**Columns:**
Order ID, Order Date, Customer Name, Country, City, Product Category, Product Name, Unit Price, Quantity, Revenue, Device Type, Payment Method, Order Status, Month Name, Year, Month.

---

## Tools & Technologies Used

- **Microsoft Excel** - data cleaning, formulas, PivotTables, PivotCharts, slicers, and dashboard design.

---

## Data Cleaning Process

Before any analysis began, a comprehensive data quality audit was performed on a duplicated copy of the raw sheet, keeping the original export untouched as a source of truth.

### 1. Duplicate Records
The dataset initially contained 10,000 records. Using Excel's *Remove Duplicates* feature, approximately **150 duplicate records** were identified and removed, reducing the dataset to 9,850 records.

### 2. Date Standardization
The `Order Date` column contained a mix of text-formatted dates and inconsistent formats (e.g., "22 December 2025") that Excel did not recognize natively. The column was standardized using *Text to Columns* with the DMY (Day-Month-Year) date format, converting all entries into valid Excel date values suitable for time-based analysis.

### 3. Country Standardization
The `Country` field contained multiple representations of the same country (e.g., "USA," "U.S.A," "United States") along with leading/trailing spaces. The `=TRIM()` function was applied to remove excess whitespace, and country names were manually standardized into one consistent naming convention per country.

### 4. Product Category Standardization
The `Product Category` field mixed uppercase, lowercase, and inconsistently spaced text. `=TRIM()` removed extra spaces, and `=PROPER()` standardized capitalization, producing clean, consistent category labels for reporting and PivotTable grouping.

### 5. Payment Method Standardization
The `Payment Method` field had inconsistent naming (e.g., "Cards" vs. "Credit Card" used interchangeably) and blank entries. *Find & Replace* was used to unify "Cards" into "Credit Card," correcting approximately **3,882 records**. The remaining **71 blank records** were filled with "Unknown" rather than removed, preserving transaction volume for revenue analysis.

### 6. Device Type Standardization
The `Device Type` field contained blank entries. Rather than discarding these rows, blanks were replaced with "Unknown" so the transactions remained available for revenue and order analysis while clearly flagging missing classification.

### 7. Quantity Validation
The `Quantity` column contained negative, zero, and blank values that would have distorted revenue calculations regardless of order status (including Completed and Refunded orders). All rows with negative, zero, or blank quantities were removed approximately **90 records** leaving only valid sales transactions.

### 8. Unit Price Validation
Some `Unit Price` values were stored as text rather than numbers, which would have broken revenue formulas. These were converted into proper numeric format to ensure accurate mathematical operations.

**Cleaning Summary:**

| Stage | Record Count |
|---|---|
| Original dataset | 10,000 |
| After duplicate removal (~150 removed) | 9,850 |
| After invalid quantity removal (90 removed) | 9,760 |
| **Final clean dataset** | **9,760** |

---

## New Calculated Columns

Three analytical fields were engineered to support deeper business analysis:

| Column | Formula | Purpose |
|---|---|---|
| **Revenue** | `= Unit Price × Quantity` | Enables revenue-based analysis across products, countries, and time periods |
| **Year** | `=YEAR(Order Date)` | Supports yearly trend analysis |
| **Month** | `=TEXT(Order Date, "MMMM")` | Supports monthly trend analysis |

---

## Data Analysis Process

With a clean dataset in place, analysis was structured around the three questions posed by management:

- **What changed?** Revenue and order volume were aggregated by Year to trace the company's growth trajectory from 2021 through 2025.
- **Who are our customers?** Revenue was broken down by Country, City, and Device Type to determine whether customer behavior was uniform or market-specific.
- **Where are we losing money?** Order Status was used to isolate cancelled and refunded orders, quantifying lost revenue in dollar terms and by market.

Each finding was written as a single sentence containing a number, backed directly by a PivotTable or formula ensuring every insight was verifiable rather than anecdotal.

---

## Pivot Table Analysis

The following PivotTables were built to drive the findings and power the dashboard:

| PivotTable | Row Field | Value Field |
|---|---|---|
| Orders by Month | Month | Sum of Orders |
| Orders by Year | Year | Sum of Total Orders |
| Revenue by Country | Country | Sum of Revenue (USD) |
| Revenue by City | City | Sum of Revenue (USD) |
| Revenue by Device Type | Device Type | Sum of Revenue (USD) |
| Revenue by Payment Method | Payment Method | Sum of Revenue (USD) |
| Lost Revenue by Order Status | Order Status | Sum of Revenue (USD) |
| KPI Summary | Indicator | Value |
<img width="1840" height="702" alt="Nova_pivot_analysis" src="https://github.com/user-attachments/assets/6e4bef58-91e5-4d76-9725-66fd4365d499" />
<img width="1352" height="677" alt="Nova_pivot_analysis2" src="https://github.com/user-attachments/assets/a15c6f0c-165e-4674-8c60-1d3a7482aa67" />
<img width="1326" height="672" alt="Nova_pivot_analysis_chart1" src="https://github.com/user-attachments/assets/c388dcac-55b5-4755-a3c8-21c566f3963b" />
<img width="1213" height="615" alt="nova_pivot_analysis_chart2" src="https://github.com/user-attachments/assets/409713c0-f215-4590-8813-8f65ab46e277" />
<img width="1395" height="687" alt="Nova_pivot_analysis_chart3" src="https://github.com/user-attachments/assets/c54315b4-d2e2-474f-86bb-ae40643a986b" />












All PivotTables were connected to three shared slicers - **Year**, **Country**, and **Product Category** allowing every table and chart to filter simultaneously from a single control point.

---

## Dashboard Development

The dashboard was designed as a single page for a sales director audience with roughly five minutes to review it. Every element was chosen to earn its space.

**KPI Cards:**
- Total Revenue - $5,171,302
- Total Orders - 9,760
- Total Quantity - 16,884
- Average Order Value - $530
- Lost Revenue - $366,327

**Charts:**
- Revenue Trend by Year (line chart) - tracks growth from 2021–2025
- Revenue by Country (bar chart) - compares market performance
- Revenue by Product Category (horizontal bar chart) - shows category concentration
- Revenue by Device Type (donut chart) - highlights mobile vs. desktop vs. tablet share
- Lost Revenue by Country (clustered bar chart) - breaks cancellations and refunds down by market
<img width="1213" height="675" alt="Nova Dashboard" src="https://github.com/user-attachments/assets/e46f73a1-4c55-4a02-9d34-7a9236cbe805" />


**Slicers:**
- Year
- Country
- Product Category

These slicers are linked across every PivotTable and chart on the page, so a single click updates the entire dashboard view.

---

## Key Insights

**Growth & Performance**
1. Revenue more than doubled - from $687,748 in 2021 to $1,493,811 in 2025, approximately **117% growth**.
2. Order volume grew significantly - from 1,291 orders in 2021 to 2,782 in 2025, approximately **116% growth**.
3. Growth was consistent - revenue increased year-over-year throughout the entire five-year period with no declines.

**Customer & Market Insights**

4. The USA is Nova Gear's largest market, generating approximately **$1.33 million** in revenue  25.7% of total sales.
5. China is the second-largest market, generating approximately **$1.01 million** in revenue.
6. Computing and Phones & Tablets dominate sales, together generating approximately **$4.21 million**  over **81%** of total revenue.
7. Mobile commerce drives purchases - mobile devices generated approximately **$2.88 million**, or **55.6%** of total sales.
8. Customer behavior varies by market - device preferences and revenue contribution differ meaningfully by country, indicating Nova Gear operates across multiple distinct customer markets rather than one homogeneous global audience.

**Revenue Leakage Insights**

9. Total revenue losses reached approximately **$366,327** from cancelled and refunded orders between 2021 and 2025.
10. Revenue leakage represents **7.1%** of total sales.
11. Losses are concentrated - the USA and China together account for nearly **40%** of all lost revenue, making them the top priority for recovery efforts.

---

## Business Recommendations

**1. Prioritize Revenue Recovery in the USA and China**
The USA and China accounted for approximately $80,180 and $65,448 in lost revenue respectively nearly 40% of all cancellations and refunds combined. A detailed review of cancelled/refunded orders in these markets is recommended, investigating product quality, delivery delays, fulfillment errors, payment processing, and customer service. Recovering revenue here requires no new customer acquisition spend.

**2. Strengthen the Mobile Commerce Experience**
Mobile devices drive 55.6% of total sales ($2.88M). Continued investment in mobile site performance, checkout flow, mobile payment integration, and browsing experience can improve conversion in the company's most important channel.

**3. Protect and Expand High-Performing Product Categories**
Computing and Phones & Tablets generate over 81% of total revenue. Management should prioritize inventory availability, marketing spend, and assortment in these categories, while monitoring the concentration risk of relying so heavily on two categories.

**4. Develop Market-Specific Strategies**
Revenue performance and device preferences vary meaningfully across countries. Rather than a single global strategy, Nova Gear should localize promotions, payment options, product recommendations, and advertising by market.

**5. Monitor Revenue Leakage as a Core Business KPI**
With $366,327 (7.1% of revenue) lost to cancellations and refunds, Nova Gear should establish ongoing tracking of refund rate, cancellation rate, and lost revenue by country and category to catch operational issues before losses grow.

**6. Diversify Revenue Sources**
With Computing and Phones & Tablets dominating over 81% of revenue, categories like Audio, Smart Home, Wearables, and Accessories represent untapped growth potential. Bundling, cross-selling, promotions, and loyalty incentives could reduce dependency on the top two categories and improve long-term resilience.

**Strategic Priority Summary:**
- Reduce cancellations and refunds in the USA and China
- Continue optimizing the mobile purchasing experience
- Protect high-performing product categories
- Implement localized market strategies
- Monitor revenue leakage through ongoing KPI tracking
- Diversify revenue across additional product categories

---

## Challenges Encountered

This dataset presented several layered data quality issues rather than a single obvious problem duplicate records, text-formatted dates, inconsistent country and category naming, mixed payment method labels, and invalid quantities all had to be identified and resolved without accidentally discarding valid transactions. Deciding how to treat blank and missing values (e.g., filling with "Unknown" instead of deleting rows) required careful judgment to preserve data integrity while keeping the analysis accurate.

The dashboard build also required deliberate prioritization: the analysis surfaced more findings than a single page could accommodate, so KPI cards, charts, and slicers were chosen specifically for their relevance to an executive audience, with less critical breakdowns left out of the final view.

---

## Conclusion

Nova Gear experienced strong and consistent growth throughout the five-year period, with revenue more than doubling and order volume growing in step. However, the business remains heavily dependent on a small number of product categories, and a measurable share of revenue  roughly 7.1%  is lost annually to cancellations and refunds, concentrated in its two largest markets.

Sustained growth going forward depends on recovering leaked revenue in the USA and China, continuing to optimize the mobile purchasing experience, diversifying beyond Computing and Phones & Tablets, and adopting market-specific strategies rather than a one-size-fits-all approach across all eight countries.

---

## Author Information

**Abdraheem Waliu Ajekunle**
Data Analyst
