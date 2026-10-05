# PR. 1 – Fundamental Booster (Microsoft Excel)

A hands-on Excel project that applies core spreadsheet skills to three small datasets: **student grades**, **sales records** and **employee data**. Every result is calculated with formulas, so the workbook updates automatically when the data changes.

**Author:** Priyanka
**File:** `PR_1_Fundamental_Booster_FINAL.xlsx`

---

## Topics Covered

| Area | Functions / Concepts |
|---|---|
| Data entry & formatting | Bold headers, ₹ currency, `dd-mmm-yyyy` dates, percentages, colour fills |
| Cell references | Relative (`A1`) and absolute (`$A$1`) references |
| Conditional logic | `IF`, nested `IF`, `IF + AND`, `IF + OR` |
| Conditional aggregation | `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` |
| Lookups | `VLOOKUP`, `XLOOKUP`, `XMATCH`, `INDEX + MATCH` |
| Text functions | `LEFT`, `FIND`, `UPPER`, `LOWER`, `TEXT` |
| Dynamic references | `INDIRECT`, `OFFSET` |
| Date & time | `YEAR`, `DATEDIF`, `TODAY`, date subtraction |
| Math functions | `ROUND`, `CEILING`, `FLOOR` |
| Dynamic arrays | `FILTER` |

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| **Project Instructions** | Original project brief and task list |
| **Students Grade** | 20 students with Math, Science, English scores. Added columns: grade classification, Math > 50, Math & Science > 80, average, first name, UPPER/lower name, enrollment year |
| **Sales Data** | 20 sales records. Added columns: discount %, discount amount, net sales, discount eligibility, month, year, quarter, GST amount |
| **Employee Data** | 20 employees. Added columns: joining year, joining month, years of service, salary band |
| **Analysis & Solutions** | All 14 task sections with formulas and results |
| **Requirement Checklist** | Each requirement from the brief mapped to its sheet, cell and function |

---

## Task Summary

| # | Task | Function(s) |
|---|---|---|
| 1 | Classify student grades (A+, A, B, C, D, F) | Nested `IF` |
| 2 | Count students scoring above 50 / 60 | `COUNTIFS` |
| 3 | Fetch product price from a product code | `VLOOKUP` |
| 4 | Total sales by region and product; average score above 60 | `SUMIFS`, `AVERAGEIFS` |
| 5 | Student name by ID | `VLOOKUP` |
| 6 | Sales for a salesperson in a given month; employee details by field | `INDEX + MATCH` |
| 7 | Employee salary and salesperson performance | `XLOOKUP` |
| 8 | Position of a product in the sales list | `XMATCH` |
| 9 | Total from a range typed as text | `INDIRECT` |
| 10 | 3-month moving sales total | `OFFSET` |
| 11 | Joining year, years of service, days between dates | `YEAR`, `DATEDIF` |
| 12 | Round sales to nearest 100, up and down to nearest 1000 | `ROUND`, `CEILING`, `FLOOR` |
| 13 | Top-performing students (average above 80) | `FILTER` |
| 14 | Age from date of birth | `DATEDIF` |

---

## Sample Results

- Students with Math > 50: **19**, Math > 60: **19**, Science > 60: **11**
- East region Keyboard sales: **₹80,349**
- Average Math score where Math > 60: **82.26**
- Person 10's sales in May 2023: **₹35,185**
- Days between first and last employee joining dates: **4,750**
- Top-performing students (average > 80): **Student 12** and **Student 13**

---

## How to Use

1. Download `PR_1_Fundamental_Booster_FINAL.xlsx`.
2. Open it in **Microsoft Excel 2021 or Microsoft 365**.
3. Start with the **Requirement Checklist** sheet to see where each task is done.
4. Open **Analysis & Solutions** to see all formulas and results.

> **Note:** `XLOOKUP`, `XMATCH` and `FILTER` need Excel 2021 or Microsoft 365. Older versions of Excel and some online viewers may show errors for these three formulas.

---

## Assumptions

- The dataset has no product price or date-of-birth column, so a small **sample price table** and **three sample dates of birth** were added. They are shown in **blue font**.
- A GST rate of **18%** is used as a sample input to demonstrate absolute references (`$P$1`).
- Years of Service shows `0` for employees whose joining date is in the future.

---

## Skills Demonstrated

Formula writing, cell referencing, data cleaning and classification, lookup techniques, dynamic ranges, date analysis, and presenting results clearly.
