## 🔹 Case 05: Merge Multiple PDFs into Single File

### 🎯 Objective
Automate the process of **combining multiple PDF files** into a single PDF document.  
Readers can use any PDF files available (e.g., reports, invoices, scanned documents).

---

### 🛠️ Steps

1. **Action: Merge PDF files**
   - Parameters:
     - `PDF file list`: Provide a list of existing PDF files (e.g., `["C:\PDFs\File1.pdf", "C:\PDFs\File2.pdf"]`)
     - `Destination file`: Choose where to save the merged file (e.g., `C:\PDFs\MergedFile.pdf`)

2. **Log Result**
   - Log message to confirm the new merged file path.

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Merge PDF Files
   Input files: [C:\Path\To\File1.pdf, C:\Path\To\File2.pdf]
   Destination: C:\Path\To\MergedFile.pdf

Log "Merged PDF saved to C:\Path\To\MergedFile.pdf"
```

---

## ✅ Expected Output
- A new PDF file is created that combines all input files into one document.
- Example: If File1 has 3 pages and File2 has 5 pages, the merged file will contain 8 pages in total.
