# E-commerce Customer Behaviour Analytics | Excel + Power BI

End-to-end data analytics mini project: data preparation in **Excel** and a 4-page interactive **Power BI** dashboard that analyses 5,000 e-commerce orders (Jan 2023 – Mar 2024) across sales, customers, products and marketing.

![Power BI](https://img.shields.io/badge/Power%20BI-Dashboard-F2C811?logo=powerbi&logoColor=black)
![Excel](https://img.shields.io/badge/Excel-Data%20Prep-217346?logo=microsoftexcel&logoColor=white)
![DAX](https://img.shields.io/badge/DAX-Measures-2EC4B6)

## Objective

Understand **how the business is performing, who the customers are, which products make money, and how marketing and loyalty affect order value**, and turn the findings into actions.

## Dataset

| Item | Detail |
|---|---|
| Source | `https://huggingface.co/datasets/maorsoul/ecommerce_customer_behavior_dataset/tree/main` |
| Size | 5,000 orders, 5,000 customers, 8 product categories, 10 cities |
| Period | 1 Jan 2023 – 26 Mar 2024 |
| Tables | `Customer_Behaviour_Dataset` (orders, 18 columns), `Product_categorycatalog_dataset` (8 rows, 9 columns), `Customer_engagement_dataset` (5,000 rows, 9 columns) |
| Currency | ₺ (Turkish Lira), assumed from Turkish city names |

## Tools

Microsoft Excel (cleaning, checks, pivot tables) · Power BI Desktop (model, DAX, visuals) · DAX

## Approach

1. **Data preparation (Excel):** checked duplicates (none found), missing values (`Preferred_Contact_Method` blank for 1,002 customers, set to "Not Specified"), data types and formats; sorted and filtered; built pivot tables for validation.
2. **Data model (Power BI):** star-style model with a `Calendar` date table.
   - `Product_categorycatalog_dataset[Product_Category]` 1 → \* `Customer_Behaviour_Dataset[Product_Category]`
   - `Customer_engagement_dataset[Customer_ID]` 1 : 1 `Customer_Behaviour_Dataset[Customer_ID]`
   - `Calendar[Date]` 1 → \* `Customer_Behaviour_Dataset[Date]`
3. **DAX:** calculated columns, measures and a calendar table (below).
4. **Visualisation:** 4 report pages with slicers, drill-down, conditional formatting, gauges and page navigation; dark navy / teal / amber theme.

## Key formulas (DAX)

**Calculated columns**
```DAX
Age Group = SWITCH(TRUE(), Customer_Behaviour_Dataset[Age] < 25, "18-24", Customer_Behaviour_Dataset[Age] < 35, "25-34", Customer_Behaviour_Dataset[Age] < 45, "35-44", Customer_Behaviour_Dataset[Age] < 55, "45-54", "55+")
Gross Profit = Customer_Behaviour_Dataset[Total_Amount] * RELATED(Product_categorycatalog_dataset[Estimated_Gross_Margin_Pct]) / 100
Delivery Speed = SWITCH(TRUE(), Customer_Behaviour_Dataset[Delivery_Time_Days] <= 3, "Fast (1-3 days)", Customer_Behaviour_Dataset[Delivery_Time_Days] <= 7, "Standard (4-7 days)", "Slow (8+ days)")
Discount Flag = IF(Customer_Behaviour_Dataset[Discount_Amount] > 0, "Discounted", "No Discount")
Customer Type = IF(Customer_Behaviour_Dataset[Is_Returning_Customer] = TRUE(), "Returning", "New")
Loyalty Tier = SWITCH(TRUE(), Customer_engagement_dataset[Loyalty_Points] >= 500, "Gold", Customer_engagement_dataset[Loyalty_Points] >= 100, "Silver", Customer_engagement_dataset[Loyalty_Points] > 0, "Bronze", "No Points")
```

**Date table**
```DAX
Calendar = ADDCOLUMNS(CALENDAR(DATE(YEAR(MIN(Customer_Behaviour_Dataset[Date])),1,1), DATE(YEAR(MAX(Customer_Behaviour_Dataset[Date])),12,31)),
    "Year", YEAR([Date]), "Month No", MONTH([Date]), "Month", FORMAT([Date],"MMM"), "Month Year", FORMAT([Date],"MMM YYYY"))
Month Year Sort = YEAR('Calendar'[Date]) * 100 + MONTH('Calendar'[Date])
```

**Measures**
```DAX
Total Revenue = SUM(Customer_Behaviour_Dataset[Total_Amount])
Total Orders = DISTINCTCOUNT(Customer_Behaviour_Dataset[Order_ID])
Avg Order Value = DIVIDE([Total Revenue], [Total Orders])
Total Gross Profit = SUM(Customer_Behaviour_Dataset[Gross Profit])
Profit Margin % = DIVIDE([Total Gross Profit], [Total Revenue])
Discount % = DIVIDE([Total Discount], [Gross Sales])
Returning Customer % = DIVIDE(CALCULATE([Total Orders], Customer_Behaviour_Dataset[Is_Returning_Customer] = TRUE()), [Total Orders])
Avg Rating vs Target = [Avg Rating] - AVERAGE(Product_categorycatalog_dataset[Category_Target_Rating])
Revenue YoY % = DIVIDE([Total Revenue] - [Revenue PY], [Revenue PY])
Loyalty Member % = DIVIDE(CALCULATE(COUNTROWS(Customer_engagement_dataset), Customer_engagement_dataset[Loyalty_Program_Member] = "Yes"), COUNTROWS(Customer_engagement_dataset))
Email Open Rate = DIVIDE(CALCULATE(COUNTROWS(Customer_engagement_dataset), Customer_engagement_dataset[Email_Opened] = "Yes"), COUNTROWS(Customer_engagement_dataset))
```
The complete list of columns and measures is in [`docs/DOCUMENTATION.md`](docs/DOCUMENTATION.md).

## Dashboard pages

| Page | Question answered | Main visuals |
|---|---|---|
| Executive Overview | How is the business performing? | 6 KPI cards, monthly revenue (area), revenue by category and city (by year), payment-method donut |
| Customer Insights | Who is buying? | Revenue by age group, gender donut, orders and AOV by device, customer type, age × category heatmap |
| Product, Profit & Delivery | What makes money? | Profit vs revenue by category, category table with rating vs target, price vs rating scatter, discount and delivery impact |
| Marketing & Engagement | How do customers arrive and stay? | Channel revenue and AOV, cart abandonment, loyalty tier, contact method, 3 gauges |

## Key findings

- Revenue **₺4.92M**, **5,000 orders**, AOV **₺983**, estimated gross profit **₺1.36M (27.7%)**.
- **Electronics = 47%** of revenue but only 38% of gross profit; **Istanbul = 27%** of revenue.
- Like-for-like Q1 revenue **+7%**, driven by AOV (+10%) while orders fell 3%.
- **All 8 categories are below their target rating** (average 3.90).
- Discounts (₺124K) do not raise order value; loyalty members spend **12% less** per order, and 35% of members hold zero points.
- Search Engine has the highest AOV (₺1,108); Social Media has the most orders but the lowest AOV (₺926).

## Repository structure

```
├── README.md
├── docs/
│   └── DOCUMENTATION.md
├── data/
│   ├── raw/                  # original dataset (or link in README)
│   └── cleaned/              # cleaned Excel file
├── powerbi/
│   └── Mini_project.pbix
├── reports/
│   ├── Dashboard.pdf         # Power BI export
│   ├── 1_Report_Narrative.pdf
│   ├── 2_Key_Insights.pdf
│   └── 3_Recommended_Solutions.pdf
└── assets/
    ├── backgrounds/          # canvas background images
    └── theme/                # Power BI theme JSON
```

## How to use

1. Clone the repository and open `powerbi/Mini_project.pbix` in **Power BI Desktop**.
2. If the data source path changes: **Home → Transform data → Data source settings** and point to the cleaned Excel file in `data/cleaned/`.
3. In Desktop, use **Ctrl + Click** on navigation buttons.

## Assumptions and limitations

- Currency assumed to be ₺ (TRY). Verify against the original source.
- Gross profit is **estimated** from catalogue margins (no cost, return or shipping data).
- Data ends on 26 Mar 2024, so March 2024 is a partial month.
- Each customer has one order, so repeat purchase and lifetime value cannot be measured.
- Gauge targets are analyst benchmarks, not company targets.

## Author

**[YOUR NAME]** · [Course / Batch] · [LinkedIn or GitHub profile]

## License

For educational purposes. Dataset © original owner; see the source link above.
