# Retail-sales-data-report-Excel-SQL-and-powerBI-
End-to-end retail sales analytics project — messy transactional data cleaned in Excel/Power Query and PostgreSQL.

# Data Cleaning Documentation — Messy Retail Sales Dataset

## Overview

This document explains the end-to-end data cleaning process applied to a simulated
e-commerce orders dataset (~1,270 raw rows). Cleaning was done in two stages:

1. **Excel / Power Query** — structural and formatting-level cleaning (duplicates,
   text formatting, categorical standardization, date formats).
2. **SQL (PostgreSQL)** — logical/business-rule cleaning after the data was loaded
   into a database (data-entry errors, cross-column consistency, imputation, and a
   derived metric).

---

## Stage 1: Excel / Power Query Cleaning

| Column | Issue Found | Method Used | Result |
|---|---|---|---|
| **OrderID** | Exact full-row duplicates | Data tab → Remove Duplicates → entire selection (also verified with `=IF(COUNTIF(C2:C1276,C2)>1,"",C2)`) | 44 duplicate rows removed |
| **OrderID** | Same OrderID appearing with different data in other columns | Data tab → Remove Duplicates → current selection (OrderID column only) | 69 duplicate OrderIDs resolved (kept one row per ID) |
| **CustomerID** | Can legitimately repeat (one customer, multiple orders) | No changes | N/A — correct as-is |
| **CustomerName** | Stray commas in names | Find & Replace `,` → `""` | 107 replacements |
| **CustomerName** | Inconsistent casing, leading/trailing spaces | `=PROPER(TRIM(CustomerName))` | Standardized to Proper Case, trimmed |
| **CustomerName** | Blank name fields, or name mismatched with the name embedded in the email address | `=IF(G2="",IF(C2&" "&D2=" ","",C2&" "&D2),IF(AND(LEFT(G2,FIND(" ",G2)-1)=D2,RIGHT(G2,LEN(G2)-FIND(" ",G2))=C2),G2,C2&" "&D2))` | Cross-checks FirstName/LastName against the name parsed from the email; uses whichever source is valid, falls back to blank if all sources are empty |
| **Phone** | Special characters (`()`, `-`, `.`, `+`) inconsistently applied | Format → General | Special characters removed, formatting standardized |
| **Product** | Leading/trailing whitespace | `=TRIM()` | Whitespace removed |
| **Category** | `&` symbol, extra spaces, inconsistent capitalization (e.g. "Electroncis", "beauty") | Find & Replace `&` → `and`, `TRIM()`, capitalize | Standardized category labels |
| **UnitPrice** | Mixed data types (currency symbols, commas, text like "unknown") | Converted column data type to Decimal | Column is now fully numeric |
| **Discount** | Null values | Replaced null with `0` | No missing discounts (0 = no discount applied) |
| **OrderDate / ShipDate** | Multiple inconsistent date formats in the same column (`YYYY-MM-DD`, `MM/DD/YYYY`, `DD-MM-YYYY`, `DD Mon YYYY`, etc.) | Data → Text to Columns → Delimited → set "Column data format" to Date (matching the source format, e.g. MDY) → Finish, then applied a consistent custom format (`m/d/yyyy`) via Format Cells → Custom | Both columns converted to true Excel date values in one consistent format |
| **Country** | Multiple spellings/abbreviations for the same country (`USA`, `U.S.A.`, `usa`, `UK`, `england`, etc.) | Power Query — value mapping/replace | Standardized to full country names |
| **PaymentMethod** | Inconsistent labels/casing (`credit card`, `CREDIT CARD`, `Cr. Card`, etc.) | Find & Replace in Excel | Standardized to consistent payment method labels |

### Note on date conversion (detailed steps used)
Because Excel stored the raw dates as **text**, formatting alone did not work. The fix required two steps:
1. **Convert text to real dates:** select the column → *Data* tab → *Text to Columns* → *Delimited* → *Next* → *Next* → under "Column data format" choose *Date* and select the matching source format (e.g. MDY) → *Finish*.
2. **Apply a consistent display format:** select the column → right-click → *Format Cells* (`Ctrl+1`) → *Custom* → enter `m/d/yyyy` → *OK*.

---

## Stage 2: SQL Cleaning (PostgreSQL)

After loading the Excel-cleaned data into PostgreSQL, the following issues — mostly
**logical/business-rule errors** that formatting tools like Power Query can't catch —
were identified and fixed with SQL.

### 1. Duplicate CustomerIDs (contact/address table)
```sql
DELETE FROM Customerdetails
WHERE CustomerId IN (
    SELECT CustomerId
    FROM (
        SELECT CustomerId,
               ROW_NUMBER() OVER (PARTITION BY CustomerId ORDER BY ctid) AS rn
        FROM Customerdetails
    ) AS sub
    WHERE rn > 1
);
```
Removes duplicate `CustomerId` entries, keeping one row per customer.

### 2. Negative / zero Quantity
```sql
-- Zero quantity: investigated and kept — represents cancelled orders
SELECT * FROM orderdetails WHERE Quantity = 0;

-- Negative quantity: data-entry sign error, corrected
UPDATE orderdetails
SET quantity = ABS(quantity)
WHERE quantity < 0;
```
**Decision:** Zero-quantity rows were **kept as-is** (interpreted as cancelled orders,
still valid business records). Negative quantities were **converted to absolute
values** (sign entry error).

### 3. ShipDate earlier than OrderDate (logically impossible)
```sql
SELECT * FROM orderdetails WHERE ShipDate < OrderDate;

UPDATE orderdetails
SET ShipDate = NULL
WHERE ShipDate < OrderDate;
```
**Decision:** Since the true ship date can't be reconstructed, these values were set
to `NULL` ("unknown") rather than guessed.

### 4. UnitPrice outliers (decimal-entry errors)
```sql
SELECT * FROM orderdetails WHERE UnitPrice > 1000 ORDER BY UnitPrice DESC;

UPDATE orderdetails
SET UnitPrice = UnitPrice / 100
WHERE UnitPrice > 1000;
```
**Decision:** Overall average UnitPrice was ~529; rows above 1000 averaged ~669 once
scaled down — dividing the outlier values by 100 brought them back in line with
normal pricing, consistent with a misplaced decimal point at data entry.

### 5. Country: merge "England" into "United Kingdom"
```sql
UPDATE customercontactaddressdetails
SET Country = 'United Kingdom'
WHERE Country = 'England';
```

### 6. PaymentMethod: merge "Wire Transfer" into "Bank Transfer"
```sql
UPDATE orderdetails
SET PaymentMethod = 'Bank Transfer'
WHERE PaymentMethod = 'Wire Transfer';
```

### 7. Email: unify "N/A" text with true NULLs
```sql
UPDATE customercontactaddressdetails
SET Email = NULL
WHERE Email = 'N/A';
```

### 8. UnitPrice nulls: imputed with the product's median price
```sql
UPDATE orderdetails o
SET UnitPrice = sub.median_price
FROM (
    SELECT Product, PERCENTILE_CONT(0.5) WITHIN GROUP (ORDER BY UnitPrice) AS median_price
    FROM orderdetails
    WHERE UnitPrice IS NOT NULL
    GROUP BY Product
) sub
WHERE o.Product = sub.Product
AND o.UnitPrice IS NULL;
```
**Why median, not average:** median is not skewed by the outlier prices found and
corrected in step 4, making it a more reliable stand-in value.

### 9. Derived column: Net Revenue
```sql
ALTER TABLE orderdetails ADD COLUMN NetRevenue NUMERIC(10,2);

UPDATE orderdetails
SET NetRevenue = (quantity * unitprice) * (1 - discount);
```
Adds a calculated `NetRevenue` column (`Quantity × UnitPrice × (1 - Discount)`) for
downstream reporting and KPI calculations.

---

## Summary of Cleaning Decisions

| Type of Issue | Approach Taken |
|---|---|
| Exact duplicate rows | Deleted |
| Duplicate keys with conflicting data | Kept first occurrence, removed rest |
| Text formatting (case, spacing, symbols) | Standardized via Excel formulas / Find & Replace |
| Inconsistent categorical values (Country, Category, PaymentMethod) | Mapped to a single standard label |
| Inconsistent date formats | Converted to a single true date format |
| Logically invalid data (negative quantity, ShipDate < OrderDate) | Corrected where the fix was inferable (sign error); nulled out where it wasn't (date) |
| Outlier values (UnitPrice) | Corrected using a scale-factor inferred from comparison to the average |
| Missing values | Zero-filled (Discount), NULL-standardized (Email), median-imputed (UnitPrice) |
| Derived metrics | Added NetRevenue as a calculated column |

## Tools Used
- **Microsoft Excel** — Find & Replace, formulas (`TRIM`, `PROPER`, `COUNTIF`, `IF`),
  Text to Columns, Format Cells, Remove Duplicates
- **Power Query** — value mapping/standardization for categorical columns
- **PostgreSQL** — window functions (`ROW_NUMBER`, `PERCENTILE_CONT`), `UPDATE`/`DELETE`
  statements, `ALTER TABLE` for schema changes
