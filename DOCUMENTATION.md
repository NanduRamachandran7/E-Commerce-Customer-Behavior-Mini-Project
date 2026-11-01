# Project Documentation: E-commerce Customer Behaviour Analytics

This document explains how the project was built: data, preparation, data model, every DAX formula, dashboard design and assumptions. It supports the summary in `README.md`.

## 1. Project overview

| Item | Detail |
|---|---|
| Goal | Analyse e-commerce customer behaviour and recommend actions |
| Tools | Excel (cleaning, pivot tables), Power BI Desktop (model, DAX, report) |
| Deliverables | Dataset source link, cleaned `.xlsx`, Power BI report (`.pbix` and PDF export), narrative, insights and recommendations (PDF) |
| Period | 1 Jan 2023 – 26 Mar 2024 |
| Volume | 5,000 orders, 5,000 unique customers |

## 2. Data dictionary

### 2.1 `Customer_Behaviour_Dataset` (orders, 5,000 rows)

| Column | Type | Description |
|---|---|---|
| Order_ID | Text | Unique order identifier |
| Customer_ID | Text | Customer identifier (links to engagement table) |
| Date | Date | Order date |
| Age | Whole number | Customer age (18–75) |
| Gender | Text | Female / Male / Other |
| City | Text | 10 cities |
| Product_Category | Text | 8 categories (links to catalogue) |
| Unit_Price | Decimal | Price per unit |
| Quantity | Whole number | Units ordered |
| Discount_Amount | Decimal | Discount in currency |
| Total_Amount | Decimal | Unit_Price × Quantity − Discount_Amount |
| Payment_Method | Text | Credit Card, Debit Card, Digital Wallet, Bank Transfer, Cash on Delivery |
| Device_Type | Text | Mobile / Desktop / Tablet |
| Session_Duration_Minutes | Decimal | Length of web session |
| Pages_Viewed | Whole number | Pages viewed in session |
| Is_Returning_Customer | True/False | Returning customer flag |
| Delivery_Time_Days | Whole number | Days to deliver (1–25) |
| Customer_Rating | Whole number | Rating 1–5 |

### 2.2 `Product_categorycatalog_dataset` (8 rows)

| Column | Description |
|---|---|
| Product_Category | Category name (key) |
| Department | Department name |
| Product_Count | Products in the category |
| Average_List_Price | Average list price |
| Estimated_Gross_Margin_Pct | Estimated gross margin (%) |
| Category_Target_Rating | Target rating |
| Return_Policy_Days | Return window |
| Free_Shipping_Threshold | Order value that unlocks free shipping |
| Category_Manager | Manager A, B, C or D |

### 2.3 `Customer_engagement_dataset` (5,000 rows)

| Column | Description |
|---|---|
| Customer_ID | Key (1:1 with orders) |
| Marketing_Channel | Direct, Email, Paid Ads, Referral, Search Engine, Social Media |
| Website_Visits_Last_30_Days | Visits in the last 30 days |
| Cart_Abandonment_Count | Abandoned carts |
| Email_Opened | Yes / No |
| Loyalty_Program_Member | Yes / No |
| Loyalty_Points | Points balance (0–1,000) |
| Customer_Satisfaction_Score | Score 1–5 |
| Preferred_Contact_Method | Email, App Notification, SMS (1,002 blanks) |

## 3. Data preparation (Excel)

| Check / step | Result |
|---|---|
| Duplicate rows (Remove Duplicates) | 0 found |
| Duplicate Order_ID | 0 found |
| Missing values | Only `Preferred_Contact_Method` (1,002 blanks), replaced with "Not Specified" |
| Consistency check | `Total_Amount = Unit_Price × Quantity − Discount_Amount` holds for all rows |
| Data types | Date as Date, amounts as decimal, counts as whole numbers, `Is_Returning_Customer` as True/False |
| Sorting / filtering | Used to inspect outliers and category/city values |
| Pivot tables | Used to validate Power BI totals (revenue by category, city, month) |

> Edit this table so it lists exactly the steps you performed in your Excel file.

## 4. Data model

| Relationship | Cardinality | Direction |
|---|---|---|
| `Product_categorycatalog_dataset[Product_Category]` → `Customer_Behaviour_Dataset[Product_Category]` | One to many | Single |
| `Customer_engagement_dataset[Customer_ID]` ↔ `Customer_Behaviour_Dataset[Customer_ID]` | One to one | Both |
| `Calendar[Date]` → `Customer_Behaviour_Dataset[Date]` | One to many | Single (active) |

Measures are stored in a separate table, `_Measures`.

## 5. Calendar table

```DAX
Calendar =
ADDCOLUMNS(
    CALENDAR(
        DATE(YEAR(MIN(Customer_Behaviour_Dataset[Date])), 1, 1),
        DATE(YEAR(MAX(Customer_Behaviour_Dataset[Date])), 12, 31)
    ),
    "Year", YEAR([Date]),
    "Quarter", "Q" & QUARTER([Date]),
    "Month No", MONTH([Date]),
    "Month", FORMAT([Date], "MMM"),
    "Month Year", FORMAT([Date], "MMM YYYY"),
    "Weekday No", WEEKDAY([Date], 2),
    "Weekday", FORMAT([Date], "ddd")
)
```
```DAX
Month Year Sort = YEAR('Calendar'[Date]) * 100 + MONTH('Calendar'[Date])
```
Marked as the date table. `Month` is sorted by `Month No`, `Month Year` by `Month Year Sort`, `Weekday` by `Weekday No`.

## 6. Calculated columns

### 6.1 `Customer_Behaviour_Dataset`

```DAX
Age Group =
SWITCH(TRUE(),
    Customer_Behaviour_Dataset[Age] < 25, "18-24",
    Customer_Behaviour_Dataset[Age] < 35, "25-34",
    Customer_Behaviour_Dataset[Age] < 45, "35-44",
    Customer_Behaviour_Dataset[Age] < 55, "45-54",
    "55+")

Age Group Sort =
SWITCH(TRUE(),
    Customer_Behaviour_Dataset[Age] < 25, 1,
    Customer_Behaviour_Dataset[Age] < 35, 2,
    Customer_Behaviour_Dataset[Age] < 45, 3,
    Customer_Behaviour_Dataset[Age] < 55, 4,
    5)

Customer Type = IF(Customer_Behaviour_Dataset[Is_Returning_Customer] = TRUE(), "Returning", "New")

Gross Order Value = Customer_Behaviour_Dataset[Unit_Price] * Customer_Behaviour_Dataset[Quantity]

Order Discount % = DIVIDE(Customer_Behaviour_Dataset[Discount_Amount], Customer_Behaviour_Dataset[Gross Order Value], 0)

Discount Flag = IF(Customer_Behaviour_Dataset[Discount_Amount] > 0, "Discounted", "No Discount")

Order Value Segment =
SWITCH(TRUE(),
    Customer_Behaviour_Dataset[Total_Amount] < 200, "Low (<200)",
    Customer_Behaviour_Dataset[Total_Amount] <= 1000, "Medium (200-1000)",
    "High (>1000)")

Delivery Speed =
SWITCH(TRUE(),
    Customer_Behaviour_Dataset[Delivery_Time_Days] <= 3, "Fast (1-3 days)",
    Customer_Behaviour_Dataset[Delivery_Time_Days] <= 7, "Standard (4-7 days)",
    "Slow (8+ days)")

Rating Category =
SWITCH(TRUE(),
    Customer_Behaviour_Dataset[Customer_Rating] >= 4, "Satisfied",
    Customer_Behaviour_Dataset[Customer_Rating] = 3, "Neutral",
    "Unsatisfied")

Session Type =
SWITCH(TRUE(),
    Customer_Behaviour_Dataset[Session_Duration_Minutes] <= 5, "Short",
    Customer_Behaviour_Dataset[Session_Duration_Minutes] <= 15, "Medium",
    "Long")

Gross Profit =
Customer_Behaviour_Dataset[Total_Amount]
    * RELATED(Product_categorycatalog_dataset[Estimated_Gross_Margin_Pct]) / 100

Free Shipping Eligible =
IF(Customer_Behaviour_Dataset[Total_Amount] >= RELATED(Product_categorycatalog_dataset[Free_Shipping_Threshold]), "Yes", "No")
```

### 6.2 `Customer_engagement_dataset`

```DAX
Loyalty Tier =
SWITCH(TRUE(),
    Customer_engagement_dataset[Loyalty_Points] >= 500, "Gold",
    Customer_engagement_dataset[Loyalty_Points] >= 100, "Silver",
    Customer_engagement_dataset[Loyalty_Points] > 0, "Bronze",
    "No Points")

Cart Abandonment Level =
SWITCH(TRUE(),
    Customer_engagement_dataset[Cart_Abandonment_Count] = 0, "None",
    Customer_engagement_dataset[Cart_Abandonment_Count] <= 2, "Low (1-2)",
    Customer_engagement_dataset[Cart_Abandonment_Count] <= 5, "Medium (3-5)",
    "High (6+)")

Visit Level =
SWITCH(TRUE(),
    Customer_engagement_dataset[Website_Visits_Last_30_Days] <= 3, "Low",
    Customer_engagement_dataset[Website_Visits_Last_30_Days] <= 8, "Medium",
    "High")
```

## 7. Measures (`_Measures`)

### 7.1 Core KPIs

```DAX
Total Revenue = SUM(Customer_Behaviour_Dataset[Total_Amount])
Total Orders = DISTINCTCOUNT(Customer_Behaviour_Dataset[Order_ID])
Total Customers = DISTINCTCOUNT(Customer_Behaviour_Dataset[Customer_ID])
Total Units Sold = SUM(Customer_Behaviour_Dataset[Quantity])
Avg Order Value = DIVIDE([Total Revenue], [Total Orders])
Total Discount = SUM(Customer_Behaviour_Dataset[Discount_Amount])
Gross Sales = SUMX(Customer_Behaviour_Dataset, Customer_Behaviour_Dataset[Unit_Price] * Customer_Behaviour_Dataset[Quantity])
Discount % = DIVIDE([Total Discount], [Gross Sales])
Total Gross Profit = SUM(Customer_Behaviour_Dataset[Gross Profit])
Profit Margin % = DIVIDE([Total Gross Profit], [Total Revenue])
```

### 7.2 Customer experience

```DAX
Avg Rating = AVERAGE(Customer_Behaviour_Dataset[Customer_Rating])
Avg Delivery Days = AVERAGE(Customer_Behaviour_Dataset[Delivery_Time_Days])
Avg Session Minutes = AVERAGE(Customer_Behaviour_Dataset[Session_Duration_Minutes])
Avg Pages Viewed = AVERAGE(Customer_Behaviour_Dataset[Pages_Viewed])

Returning Customer % =
DIVIDE(CALCULATE([Total Orders], Customer_Behaviour_Dataset[Is_Returning_Customer] = TRUE()), [Total Orders])

Satisfied Customers % =
DIVIDE(CALCULATE([Total Orders], Customer_Behaviour_Dataset[Customer_Rating] >= 4), [Total Orders])

Avg Rating vs Target = [Avg Rating] - AVERAGE(Product_categorycatalog_dataset[Category_Target_Rating])
```

### 7.3 Time intelligence

```DAX
Revenue PY = CALCULATE([Total Revenue], SAMEPERIODLASTYEAR('Calendar'[Date]))
Revenue YoY % = DIVIDE([Total Revenue] - [Revenue PY], [Revenue PY])
Revenue PM = CALCULATE([Total Revenue], DATEADD('Calendar'[Date], -1, MONTH))
Revenue MoM % = DIVIDE([Total Revenue] - [Revenue PM], [Revenue PM])
Revenue YTD = TOTALYTD([Total Revenue], 'Calendar'[Date])

Cumulative Revenue =
CALCULATE([Total Revenue], FILTER(ALL('Calendar'[Date]), 'Calendar'[Date] <= MAX('Calendar'[Date])))
```

### 7.4 Share, rank and top-N

```DAX
Revenue % of Total = DIVIDE([Total Revenue], CALCULATE([Total Revenue], ALLSELECTED()))

Category Rank = RANKX(ALLSELECTED(Customer_Behaviour_Dataset[Product_Category]), [Total Revenue], , DESC)

Top Category =
CALCULATE(MAXX(TOPN(1, ALL(Customer_Behaviour_Dataset[Product_Category]), [Total Revenue]), Customer_Behaviour_Dataset[Product_Category]))

Top City =
CALCULATE(MAXX(TOPN(1, ALL(Customer_Behaviour_Dataset[City]), [Total Revenue]), Customer_Behaviour_Dataset[City]))
```

### 7.5 Engagement

```DAX
Avg Website Visits = AVERAGE(Customer_engagement_dataset[Website_Visits_Last_30_Days])
Total Cart Abandonments = SUM(Customer_engagement_dataset[Cart_Abandonment_Count])
Avg Cart Abandonment = AVERAGE(Customer_engagement_dataset[Cart_Abandonment_Count])
Avg Satisfaction Score = AVERAGE(Customer_engagement_dataset[Customer_Satisfaction_Score])
Total Loyalty Points = SUM(Customer_engagement_dataset[Loyalty_Points])

Loyalty Member % =
DIVIDE(CALCULATE(COUNTROWS(Customer_engagement_dataset), Customer_engagement_dataset[Loyalty_Program_Member] = "Yes"), COUNTROWS(Customer_engagement_dataset))

Email Open Rate =
DIVIDE(CALCULATE(COUNTROWS(Customer_engagement_dataset), Customer_engagement_dataset[Email_Opened] = "Yes"), COUNTROWS(Customer_engagement_dataset))
```

### 7.6 Titles, colours and gauge targets

```DAX
Dynamic Title =
"Revenue by Month - " &
IF(ISFILTERED(Product_categorycatalog_dataset[Product_Category]),
   SELECTEDVALUE(Product_categorycatalog_dataset[Product_Category], "Multiple Categories"),
   "All Categories")

Rating Color = IF([Avg Rating] >= 4, "#2EC4B6", IF([Avg Rating] >= 3.5, "#FF9F1C", "#E71D36"))
```
The gauge target measures (`Target Email Open Rate`, `Target Loyalty Member`, `Target Satisfaction Score`) are constants chosen as analyst benchmarks.

## 8. Report pages

| Page | Visuals | Slicers |
|---|---|---|
| **1. Executive Overview** | 6 cards (Total Revenue, Total Orders, Avg Order Value, Total Gross Profit, Avg Rating, Returning Customer %); clustered bar of revenue by Product_Category (series: Year); clustered bar of revenue by City (series: Year); area chart of revenue by Month Year; donut of orders by Payment_Method | Year, Product_Category, City, Gender |
| **2. Customer Insights** | Column chart revenue by Age Group; bar chart revenue by Customer Type; matrix Age Group × Product_Category (conditional colour); column chart orders and AOV by Device_Type; donut orders by Gender | City, Gender, Device_Type |
| **3. Product, Profit & Delivery** | Column chart Total Gross Profit and Total Revenue by category; table (Product_Category, Department, Category_Manager, Total Revenue, Profit Margin %, Avg Rating, Avg Rating vs Target); scatter (X Unit_Price, Y Avg Rating, size Total Revenue); column chart Avg Rating by Delivery Speed; column chart orders and rating by Discount Flag | Category_Manager, Year |
| **4. Marketing & Engagement** | Bar chart revenue and AOV by Marketing_Channel; column chart Avg Cart Abandonment by channel; donut Total Customers by Loyalty Tier; column chart AOV by Loyalty_Program_Member; bar chart customers by Preferred_Contact_Method; gauges for Email Open Rate, Avg Satisfaction Score, Loyalty Member % (with targets) | Marketing_Channel, Loyalty_Program_Member, Email_Opened, Category_Manager, Year |

**Interactivity:** slicers (dropdown style), cross-filtering between visuals, conditional formatting (matrix gradient, red/green rating vs target), page navigation buttons (Back/Next), titles and tooltips.

## 9. Design system

| Role | Colour |
|---|---|
| Canvas | `#0B1626` |
| Wallpaper | `#08111F` |
| Header band | `#0A1220` |
| Visual cards | `#172A46` |
| Card border | `#2A3F5F` |
| Primary accent (teal) | `#2EC4B6` |
| Secondary accent (amber) | `#FF9F1C` |
| Highlight (sky blue) | `#5BC0EB` |
| Alert (coral) | `#E71D36` |
| Text / muted text | `#FFFFFF` / `#A9B8CC` |

Layout grid: 1280×720 canvas, 16 px margins, 12 px gaps, header band 64 px with a teal accent line.

## 10. Method notes

- **Gross profit** = order amount × category estimated margin ÷ 100 (catalogue margin, not accounting profit).
- **Like-for-like growth:** Q1 2023 (90 days) vs 1 Jan–26 Mar 2024 (86 days) to avoid the partial-month bias.
- **Validation:** totals in Power BI (revenue ₺4,915,545; 5,000 orders; AOV ₺983.11; average rating 3.90; returning 59.8%) were checked against independent calculations on the source data.

## 11. Assumptions and limitations

1. Currency assumed to be ₺ (TRY).
2. Data ends 26 Mar 2024; March 2024 is incomplete.
3. One order per customer: no repeat-purchase or lifetime-value analysis.
4. No cost, return or shipping data: profit is estimated.
5. Weak correlations between rating, satisfaction, delivery time and order value (|r| < 0.02): findings are descriptive.
6. Gauge targets are analyst benchmarks.

## 12. Future improvements

- Add cost, returns and refund data for true profit.
- Keep full order history per customer for cohort and lifetime-value analysis.
- Add a drill-through page per category and matched-period comparison.
- Refresh with full months and set business-approved gauge targets.

## 13. Related files

| File | Purpose |
|---|---|
| `reports/1_Report_Narrative.pdf` | Page-by-page story of the report |
| `reports/2_Key_Insights.pdf` | 21 evidence-based insights |
| `reports/3_Recommended_Solutions.pdf` | 11 actions, KPI scorecard and roadmap |
