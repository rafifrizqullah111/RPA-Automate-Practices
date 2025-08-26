## 🔹 Case 02: Extract Tables from PDF and Save to Excel

### 🎯 Objective
Automate the process of **extracting tabular data** (e.g., line items, reports, transactions) from a PDF and storing them into an Excel file for further analysis or reporting.

---

### 🛠️ Steps

1. **Action: Extract tables from PDF**
   - Parameters:
     - `PDF file path`: `C:\PDFs\Report.pdf`
     - `Page range`: All
   - Output: List of DataTables → `%PDFTables%`

2. **Loop Through Extracted Tables**
   - `For each %Table% in %PDFTables%`:
     - Write each `%Table%` into an Excel sheet.

3. **Action: Create/Open Excel file**
   - File path: `C:\PDFs\ReportData.xlsx`
   - Output: `%ExcelInstance%`

4. **Action: Write DataTable to Excel worksheet**
   - Write `%Table%` to a new sheet (e.g., `Sheet1`, `Sheet2`, ...).

5. **Save and Close Excel**
   - Save changes and close `%ExcelInstance%`.

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Extract Tables from PDF
   File: C:\PDFs\Report.pdf
   Pages: All
   → %PDFTables%

Launch Excel → %ExcelInstance%
Create New Workbook

Set %SheetIndex% = 1

For Each %Table% in %PDFTables%
   Write DataTable to Excel
      DataTable: %Table%
      Worksheet: Sheet%SheetIndex%
   Increment %SheetIndex%

Save Workbook as C:\PDFs\ReportData.xlsx
Close Excel
```

---

## ✅ Expected Output
- `ReportData.xlsx` created in `C:\PDFs\`.
- Each extracted table from the PDF is written into a separate sheet:
  - Sheet1 → Table 1 (e.g., summary)
  - Sheet2 → Table 2 (e.g., detailed line items)
  - Sheet3 → Table 3 (if exists)

Example (Sheet1):

| Date       | Item            | Qty | Price | Total   |
| ---------- | --------------- | --- | ----- | ------- |
| 10/08/2025 | Product A       | 2   | 100   | 200     |
| 10/08/2025 | Product B       | 1   | 150   | 150     |
|            | **Grand Total** |     |       | **350** |
