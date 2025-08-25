# 📘 Excel Automation Practice (Power Automate Desktop)

This section contains **hands-on practice cases** for learning Excel Automation in Power Automate Desktop (PAD).  
The compact 10 cases cover core scenarios from reading/writing, data shaping, analytics, to report generation.

---

## 📂 Cases Overview

### 01. Reading & Writing Basics
- Open a workbook, read a range, write values to specific cells, then save & close.  
- **Covers:** `Launch Excel`, `Read from Excel worksheet`, `Write to Excel worksheet`, `Save Excel`, `Close Excel`.

### 02. Data Filtering (No Loop)
- Filter rows by conditions (e.g., Department = IT and Salary > 5000) and output to a new file.  
- **Covers:** `Read from Excel worksheet`, `Filter list`, `Write to Excel worksheet`, `Save Excel`.

### 03. Multi-Level Sorting (No Loop)
- Sort by Department (A→Z) and then by Salary (Z→A), export sorted result.  
- **Covers:** `Read from Excel worksheet`, `Sort list` (multi-pass), `Write to Excel worksheet`, `Save Excel`.

### 04. Mathematical Operations
- Compute total, average, min, max from a numeric column and write a summary table.  
- **Covers:** `Read from Excel worksheet`, `Get column from data table` / list building, `Calculate`, `Get maximum value of list`, `Get minimum value of list`, `Write to Excel worksheet`.

### 05. Conditional Formatting
- Highlight salaries below 5000 (red) and above 7000 (green) directly in the sheet.  
- **Covers:** `Read from Excel worksheet`, (indexing rows) `Set cell color`, `Save Excel`.

### 06. Lookup Operations (Join Two Sources)
- Enrich EmployeeData.xlsx with DepartmentBudget.xlsx by matching Department → Budget.  
- **Covers:** `Read from Excel worksheet` (two files), `Find in data table` / `Filter list`, `Write to Excel worksheet`, `Save Excel`.

### 07. Pivot Table Automation
- Create/refresh a Pivot Table summarizing Total Sales by Region and Product via macro.  
- **Covers:** `Run Excel Macro`, template/macro setup, `Save Excel`, optional `Export to PDF`.

### 08. Charts Automation
- Generate/update a chart (e.g., clustered column) from Pivot/summary range via macro.  
- **Covers:** `Run Excel Macro`, `Save Excel`, optional `Export to PDF`.

### 09. Data Cleaning Pipeline
- Remove duplicates, trim spaces, standardize text case, replace blanks, and fix totals.  
- **Covers:** `Run Excel Macro` (batch cleanup), or `For each` + table operations, `Save Excel`.

### 10. Report Generation (End-to-End)
- Build a weekly report: pivot summaries, chart, formatting, export to Excel & PDF.  
- **Covers:** `Run Excel Macro` (pivot, chart, formatting), `Save Excel`, `Export to PDF`.

---

## 🎯 Goal
By completing these 10 compact cases, you will gain hands-on experience in:  
- Opening, reading, writing, and saving Excel workbooks reliably  
- Filtering and sorting datasets efficiently without loops  
- Performing numerical aggregations and building summary tables  
- Enriching data via lookups across multiple workbooks  
- Automating Pivot Tables and Charts with reusable macros/templates  
- Cleaning raw data (deduplication, trimming, normalization) at scale  
- Generating polished reports (Excel & PDF) suitable for business stakeholders

This practice set is designed to mimic **real-world Excel automation** tasks commonly found in business processes using **Power Automate Desktop**.
