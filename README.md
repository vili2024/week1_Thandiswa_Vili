# ShopSmart Analysis - Week 1

**Intern:** Thandiswa Vili
**Status:** All parts Completed

### Progress: Done

- [x] PART 1 — Data Understanding (10 Marks) - DONE
  - Identified all 9 datasets and their purpose
  - Listed Primary Keys and Foreign Keys for each table
  - Mapped relationships between tables
    
- [x] PART 2 — Data Cleaning (15 Marks) - DONE
  - Removed 47 duplicate rows across datasets
  - Checked missing values and handled nulls
  - Corrected date formats to YYYY-MM-DD
  - Converted currency fields to numeric (removed currency symbols and text)
  - Converted all 9 sheets to Excel Tables (Ctrl+T)
  - Renamed worksheets to proper names

- [ ] PART 3 — Excel Formulas (20 Marks) - DONE

**Revenue:** 2,496,989,808
=SUMIF(L2:L50000,"Completed",M2:M50000)

**Quantity:** 49,432
=SUMIF(L2:L50000,"Completed",K2:K50000)

**Avg Selling Price:** 50,510.98
=Revenue / Quantity

**Avg Order Value:** 151,067.20
=Revenue / Orders

**Highest Sale:** 49,935
=MAXIFS(M2:M50000,L2:L50000,"Completed")

**Lowest Sale:** 1,029
=MINIFS(M2:M50000,L2:L50000,"Completed")

**Orders:** 16,529
=COUNTIF(L2:L50000,"Completed")

**Customers:** 4,795
=COUNTA(UNIQUE(FILTER(B2:B50000,L2:L50000="Completed")))

**Products:** 500
=COUNTA(UNIQUE(FILTER(D2:D50000,L2:L50000="Completed")))

**Stores:** 25
=COUNTA(UNIQUE(FILTER(E2:E50000,L2:L50000="Completed")))

- [ ] PART 4 — Lookup Functions (15 Marks)-NEXT
1. Customer Name: `=XLOOKUP(B2, Customers!A:A, Customers!B:B)`
2. Store Name: `=XLOOKUP(C2, Stores!A:A, Stores!B:B)`
3. Product Name: `=XLOOKUP(D2, Products!A:A, Products!B:B)`
4. Category: `=XLOOKUP(D2, Products!A:A, Products!C:C)`
5. Brand: `=XLOOKUP(D2, Products!A:A, Products!D:D)`
6. Selling Price: `=XLOOKUP(D2, Products!A:A, Products!G:G)`
7. Supplier Name: `=XLOOKUP(XLOOKUP(D2, Products!A:A, Products!H:H), Suppliers!A:A, Suppliers!B:B)`


- [ ] PART 5 — Pivot Tables (20 Marks)
Sheet: `Pivot Tables` | Grand Total = 2,292,404,381 verified across all pivots

- **Month:** Dec = 204.8M (peak), Jan lowest - holiday trend
- **Store:** STR003/STR007 top (~100M each)
- **Category:** Beauty = 533.7M (23%) top, Home lowest - invest in Beauty
- **Brand:** Thomas Ltd top
- **State:** Lagos = 384.7M top
- **Membership:** Gold = 778M top - Gold spends most
- **Gender:** Male 1.19B vs Female 1.09B
- **Top:** Customer CUST01257 = 2.05M, Product PROD0210, Bottom: PROD0430 lowest
  
- [ ] PART 6 — Charts (10 Marks)
Sheet: `Charts` - 6 charts from pivots

- **Monthly Sales:** Line chart - Dec peak
- **Revenue by Category:** Pie chart - Beauty 23% biggest
- **Top Products:** Bar chart - PROD0210 top
- **Revenue by Store:** Bar chart - all stores
- **Sales by Region:** Bar chart - South = 839.8M best region
- **Return Rate by Category:** Bar chart - Beauty 5.08% highest return

All charts have titles, labels, legend.
      
- [ ] PART 7 — Dashboard (30 Marks)
Sheet: `Dashboard` | KPIs: Revenue N2.29B, Orders 50k, Customers 5k, Avg N45k, Return 4.89%, Inventory N4.59B

**Slicers (6):** Year (2025), Month, Region (East/North/South/West), Category (Beauty etc), Store_ID (STR001-007), Customer Segment - all linked

**Charts (5):** Monthly Trend (Line - Dec peak), Revenue by Category (Pie - Beauty 23%), Revenue by Store (Bar), Top Products (Bar), Sales by State (Bar)

Interactive dashboard, filters work.

- [ ] PART 8 — Business Questions (30 Marks)
**Q1. Which product generated highest revenue?**
Ans: PROD0210 - Found via Pivot Sales by Product, sorted Sum of Total Amount descending

**Q2. Which store performed best?**
Ans: STR003 - Highest total sales in Sales by Store pivot

**Q3. Which category sells the most?**
Ans: Beauty - Highest sales in Category pivot (533.7M / 23%)

**Q4. Which month generated highest revenue?**
Ans: December - Seasonal peak, Sales by Month pivot

**Q5. Which customers spend the most?**
Ans: Gold Membership customers - Highest spend (778M) and avg order value

**Q6. Which products are frequently returned?**
Ans: PROD0263 most returned (38 times), PROD0349 (14). Home category 5.13% return rate. Fix quality for PROD0263

**Q7. Which supplier supplies highest number of products?**
Ans: All suppliers equal - each supplies 500 products. No dominant supplier

**Q8. Which store has low inventory?**
Ans: STR011 with only 9 inventory items - needs urgent restocking

**Q9. Which region performs best?**
Ans: South region - 839.8M highest

**Q10. Which category should management invest in?**
Ans: Invest in Beauty category, South region, December promotions. Also improve Home category to reduce 5.13% return rate and restock STR011
      
- [ ] PART 9 — Recommendations (20 Marks)
**3 KEY FINDINGS:**
1. Low Completion (33%) - Only 16,529 of 50k orders Completed. Rest Pending/Cancelled = data risk, inflates demand
2. Winners Concentrated - Beauty top (533.7M), PROD0210 top product, STR003 top store, South top region (839.8M), Dec peak. But Home 5.13% return highest, PROD0263 returned 38 times = quality issue
3. Inventory Imbalance - STR011 only 9 items (stock-out risk), other stores full. Suppliers equal (500 each) = no supplier risk

**3 RECOMMENDATIONS:**
1. Fix Order Process - Report on Completed only. Move Pending to pipeline, auto follow-up to raise completion from 33% to 60%
2. Fix Returns - Audit PROD0263 & Home category (5.13% return). Check quality, description, supplier. Remove/repackage PROD0263
3. Invest in Winners - Restock STR011 urgent, push Beauty + South + Dec promo. Standardise discounts (currently 10% Beauty/Grocery vs 5% Home)

**RISK:** Revenue overstatement by 67% if using 50k vs 16k Completed = bad cashflow & over-purchasing

**OPPORTUNITY:** Recover 20% of 33k Pending = double revenue without new customers. Focus Beauty/South/Dec = fastest growth
      

### Files (Final submission will be)
- ShopSmart_Analysis.xlsx
- Executive_Dashboard.pdf
- Business_Report.pdf
- README.md (this file)

Last updated: 19 Oct 2026 - All parts done
