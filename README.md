# Retail-Sales-Analytics-Data-Cleaning-Excel-Dashboards
Cleaned a messy 6-lakh-row retail sales dataset with Python (pandas) and built a 3-page Excel dashboard on revenue, profit, discounts, brands and payments, with business recommendations.
# Retail Sales Analytics: Data Cleaning & Excel Dashboards

An end-to-end analytics project: a very messy 6-lakh-row retail sales file is cleaned with Python (pandas), validated, and turned into a 3-page interactive Excel dashboard for management reporting.

---

## 1. Project Goal

Take a raw, error-filled retail sales export and deliver:

1. A **trustworthy, analysis-ready dataset**, with every cleaning decision documented.
2. A **manager-ready dashboard** that answers: *What sells, where, at what profit, and where are we losing money?*

---

## 2. The Data

| | Raw file | Cleaned file |
|---|---|---|
| Rows | 5,99,994 | 5,75,513 |
| Columns | 47 | 46 |
| Unique orders | 5,78,919 (after removing duplicates) | 5,75,513 |
| Period | 2022 to Oct 2026 | 2022 to Oct 2026 |

The raw workbook also contained two lookup sheets: **Product_Master** (39 products, MRP ranges, GST %, brands) and **Store_Master** (store, city, state, region).

**Dashboard scope:** the dashboard is built on **50,709 orders from 1 Jan to 19 Aug 2022** (28 columns). August is a partial month, so its dip is not a real decline.

### Main columns
Order & date (Order_ID, Order_Date, Year/Quarter/Month), customer (Customer_ID, Gender, Age_Group), geography (City, State, Region), product (Product, Category, Sub_Category, Brand), money (Qty, Unit_Price, Discount_Pct, Gross_Sales, Discount_Amt, Net_Sales, GST_Amt, Shipping, Cost_Price, Profit), and operations (Payment_Mode, Order_Status, Is_Returned, Rating, Delivery_Days, Coupon, Device, Data_Quality_Note).

---

## 3. Data Problems Found & How They Were Fixed

| Area | Problem | Fix |
|---|---|---|
| Structure | 1,166 blank rows, 352 "PAGE TOTAL" rows, 6 repeated header rows | Removed |
| Duplicates | 19,551 exact duplicate orders | Kept one row per Order ID |
| Order ID | Mixed formats (`ORD-2024-0000001`, `ORD9`, `#9`, lowercase) | Standardised to `ORD-YYYY-NNNNNNN` |
| Dates | 7 formats plus placeholders (2099-01-01, 1970-01-01, "TBD", "yesterday", 31/02) | Parsed each format; invalid dates blanked; ship date before order date rejected |
| Quantity | Words ("one", "do", "ek"), "2 pcs", 0, negatives, 999 | Cross-checked against Gross / Unit Price and corrected |
| Prices & amounts | ₹, Rs., $, commas, values ×1000 too large | Parsed, range-checked against Product_Master MRP, recomputed |
| Discount % | 150, -5, "Festive", fractions like 0.1 | Recovered from discount amount and net sales |
| Cost / Profit | Wrong scale or missing | Fixed scale, derived from profit, or estimated (flagged) |
| Product | Invalid IDs and name typos | Matched via product name to Product_Master |
| Geography | 306 city spellings, 197 state spellings, city-state conflicts | Mapped to 36 cities / 21 states; state and region derived from city |
| Customer | Missing IDs, name typos, missing gender/age | ID recovered from email; name/gender/age from the same customer |
| Categories | Gender, channel, payment, status, coupon, device variants | Standardised; missing values labelled `Unknown` |

### Validation checks (all passed with 0 violations)
- `Net_Sales = Gross_Sales - Discount_Amt`
- `Gross_Sales = Qty × Unit_Price`
- `Profit = Net_Sales - Cost_Price - Shipping`
- Ship date is never before order date

### Honest notes
- **3,406 rows (0.6%)** were removed because no usable amount or date could be recovered.
- Where a value was estimated (order date, cost, shipping, delivery days), the row is marked in the `Data_Quality_Note` column.
- `Ship_Date`, `Delivery_Days` and `Rating` are intentionally left blank where not applicable (e.g. orders not delivered, or not rated). They are not invented.
- Revenue in the dashboard includes Cancelled and Returned orders (about 13% of orders). Filter by `Order_Status` for strict net revenue.

---

## 4. Dashboards

### Page 1: Executive Overview
KPIs: Revenue, Profit, Orders, Quantity. Charts: monthly revenue & profit trend, revenue by channel, top 10 states, revenue & profit by region.

### Page 2: Sales & Product Analysis
KPIs: Revenue, Profit, Margin, Average Discount. Charts: category revenue vs profit, top 10 sub-categories, discount sensitivity, coupon performance, plus a Category > Sub-Category > Product drill-down table.


### Page 3: Brand & Transaction Analysis
KPIs: Avg delivery time, Avg rating, Return rate, Orders. Charts: brand share, top brands by payment mode, orders by status, revenue by payment mode, plus a Top-20 transaction audit log.


---

## 5. Key Insights (1 Jan – 19 Aug 2022)

- Revenue **₹60.09 Cr**, profit **₹12.78 Cr**, margin **21.3%**, average order value **₹11,850**.
- **North India = 42%** of revenue; UP, Maharashtra and Delhi together contribute about 40%.
- **Electronics = 49.5%** of revenue; **Laptops alone = 34.6%**.
- **Discounts erode profit:** margin is about 30% at 0% discount, about 12% at 20%, and turns **negative at 30%+**.
- **Snacks** is loss-making; Grocery, Skincare and Stationery have the thinnest margins.
- **UPI = 36%** of revenue; Cash on Delivery is still 18%.
- Only **69%** of orders are delivered; about **13%** are cancelled or returned. Return rate is 7.8%, average rating 3.9/5, average delivery about 6 days.

## 6. Recommendations

1. Cap discounts at about 20%. Beyond that, margin collapses.
2. Protect laptop supply, but grow mobiles and home appliances as a second engine.
3. Reprice or bundle low-margin items (snacks, stationery, skincare).
4. Run targeted campaigns in East and Central India.
5. Push UPI over COD and investigate cancellation/return reasons.
6. Bring delivery time under 5 days to lift the customer rating.

---

## 7. Tech Stack

- **Python**: pandas, NumPy (cleaning, validation, reconciliation)
- **Excel**: formulas, dropdown filters, charts, outline drill-down
- **xlsxwriter / openpyxl**: report and file generation

---

## 8. Repository Structure

```
retail-sales-analytics/
├── README.md
├── images/
│   ├── Dashboard_1_Overview.png
│   ├── Dashboard_2_Sales_Product.png
│   └── Dashboard_3_Brand_Transactions.png
├── dashboard/
│   └── Retail_Sales_Dashboard.xlsx
└── data/
    └── README.md   # link to the full cleaned dataset (see note below)
```
---

## 9. How to Use

1. Download `Retail_Sales_Dashboard.xlsx` from the `dashboard/` folder.
2. Open it in Excel and use the dropdown filters at the top right of each page. All charts update automatically.
3. The `Data` sheet holds the 50,709 cleaned orders behind the dashboard.

---

## 10. Author

**Harsh Bisht**: MBA (Business Analytics & AI)

Skills:  Excel

Feedback and suggestions are welcome. Feel free to open an issue or connect on LinkedIn.
