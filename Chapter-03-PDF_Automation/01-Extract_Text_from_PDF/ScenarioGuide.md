## 🔹 Case 01: Extract Invoice Data from PDF and Save to Excel

### 🎯 Objective
Automate the process of **extracting text from a PDF invoice** and storing the extracted content into an Excel file for further reporting or reconciliation.

---

### 🛠️ Steps

1. **Action: Extract text from PDF**
   - Parameters:
     - `PDF file path`: `C:\PDFs\Invoice.pdf`
     - `Page range`: All
   - Output: Extracted text → `%PDFText%`

2. **Action: Create new Excel file**
   - File path: `C:\PDFs\InvoiceData.xlsx`
   - Output: Excel instance → `%ExcelInstance%`

3. **Action: Write to Excel worksheet**
   - Write raw `%PDFText%` into first cell (`A1`) for logging.

4. **(Optional - Parsing Step)**
   - Use **Find text position in text** or regex to extract structured fields like:  
     - `Invoice Number:` → `%InvoiceNo%`  
     - `Date:` → `%InvoiceDate%`  
     - `Total:` → `%InvoiceTotal%`
   - Write these values into separate columns in Excel:
     - Column A: Invoice Number  
     - Column B: Date  
     - Column C: Total  

5. **Action: Save and close Excel**
   - Save changes to `InvoiceData.xlsx`.

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Extract Text from PDF
   File: C:\PDFs\Invoice.pdf
   Pages: All
   → %PDFText%

Launch Excel → %ExcelInstance%
Create New Workbook

Write to Excel
   Cell: A1
   Value: %PDFText%

Parse %PDFText% with Regex:
   Invoice No: (\d+)
   Date: (\d{2}/\d{2}/\d{4})
   Total: (\d+\.\d{2})

Write to Excel:
   A2: %InvoiceNo%
   B2: %InvoiceDate%
   C2: %InvoiceTotal%

Save Workbook as C:\PDFs\InvoiceData.xlsx
Close Excel
```

---

## ✅ Expected Output
- `InvoiceData.xlsx` created in `C:\PDFs\`.
- First sheet contains:
  - Full raw invoice text in cell A1.
  - Extracted structured data (Invoice Number, Date, Total) in row 2.
 
| Invoice Number | Date       | Total  |
| -------------- | ---------- | ------ |
| 12345          | 10/08/2025 | 550.00 |
