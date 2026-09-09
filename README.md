Sales & Order Analytics — Microsoft Excel

Project Overview

This project demonstrates an end-to-end Excel data analytics workflow using a deliberately messy retail sales-order dataset. The goal was to take raw transactional data through data auditing, cleaning, transformation, validation, analysis, visualization, and refresh automation using Microsoft Excel and Power Query.

The original dataset contained 244 rows and 15 columns covering orders across regions, product categories, sales channels, payment methods, sales representatives, and order statuses. Additional test rows were later added to validate the refresh process and confirm that the transformation logic re-runs correctly on new data.

Tools & Techniques

Microsoft Excel 2022

Excel Tables and structured references

Power Query

PivotTables and PivotCharts

Data Validation and Conditional Formatting

Advanced Excel formulas

Charts and Sparklines

Key Excel Concepts Applied

Data Auditing & Cleaning

Audited raw data using COUNTA, COUNTBLANK, filters, and duplicate checks.

Standardized customer names, regions, categories, channels, payment methods, order statuses, and notes.

Used functions such as TRIM, CLEAN, PROPER, SUBSTITUTE, VALUE, TEXT, and date functions.

Converted inconsistent dates, quantities, prices, and discounts into analysis-ready formats.

Created helper fields including cleaned categories, cleaned dates, month labels, revenue, high-value flags, and normalized discount values.

Formula-Based Analysis

Applied:

XLOOKUP

VLOOKUP

IF / IFERROR

SUMIFS

COUNTIFS

MAXIFS

FILTER

UNIQUE

SORT

EOMONTH

These were used for lookups, conditional calculations, multi-criteria analysis, latest-order identification, dynamic filtering, and monthly reporting.

Power Query ETL

Built a reusable transformation workflow including:

Data type correction

Trim / Clean / capitalization

Standardization using Replace Values

Null handling

Date normalization

Duplicate identification and removal

Merge with regional target data using a Left Outer Join

Append of Online and In-Store order queries

Unpivot of wide monthly data into a tidy format

Group By for category and channel/status summaries

Custom calculated fields such as Revenue and Percent of Target

Load-to-Excel and automated refresh testing

A full refresh cycle was tested by adding deliberately messy new source rows and confirming that Power Query automatically re-applied the cleaning and transformation steps.

PivotTables & Business Analysis

Created PivotTables to analyze:

Revenue by Category and Channel

Regional order count and average order value

Payment-method performance

Monthly revenue trends

SalesRep performance

Order-status distribution

Region-by-Category revenue

Average discount by category

Data Quality & Validation

Implemented:

Raw vs cleaned row-count reconciliation

Blank Quantity and UnitPrice checks

Distinct Region and Order Status checks

Data Validation lists for controlled Region entry

Manual cross-checks of automated calculations

Validation of Power Query refresh outputs

Data Visualization

Created:

Clustered Column Charts

Line Charts

Interactive PivotCharts

Horizontal Bar Charts

Combo Charts with a Secondary Axis

SalesRep and Category Sparklines

Charts were formatted for readability using meaningful titles, axis labels, legends, sorting, field-button cleanup, and simple professional styles.

Repository Files

sales-order-analytics-excel/
│
├── README.md
├── excel/
│   ├── Orders_Working.xlsx
│   └── Orders_PowerQuery_Practice.xlsx
│
└── assets/
    ├── revenue_by_category.png
    ├── monthly_revenue_trend.png
    ├── revenue_discount_combo.png
    └── sparklines.png

File Description

Orders_Working.xlsx
Main Excel analysis workbook containing formula-based cleaning, helper columns, advanced formulas, PivotTables, PivotCharts, charts, QA checks, and Sparklines.

Orders_PowerQuery_Practice.xlsx
Power Query workbook containing the repeatable ETL workflow, including cleaning, duplicate handling, Merge, Append, Unpivot, Group By, custom columns, loading, and refresh validation.

assets/
Recommended screenshots of key outputs so the project can be reviewed directly on GitHub without downloading the Excel files.

How to Review the Project

Read this README for an overview of the workflow and skills demonstrated.

Review the screenshots in the assets folder for quick visual examples.

Download and open the Excel workbooks in Microsoft Excel to inspect formulas, PivotTables, charts, and Power Query queries.

For the Power Query workbook, use Data → Queries & Connections to inspect the saved transformation steps.

Power Query functionality is best reviewed in the desktop version of Microsoft Excel.

Key Learning Outcomes

This project demonstrates the ability to:

Transform messy transactional data into analysis-ready data.

Build repeatable Excel and Power Query workflows instead of relying on one-time manual cleaning.

Apply advanced formulas to answer business questions.

Create PivotTables and interactive visualizations for management reporting.

Validate outputs using structured data-quality checks.

Automate recurring data preparation through Power Query refresh.
