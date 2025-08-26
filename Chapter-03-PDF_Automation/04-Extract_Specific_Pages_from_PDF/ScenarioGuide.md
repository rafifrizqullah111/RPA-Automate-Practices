## 🔹 Case 04: Extract Specific Pages from PDF to a New File

### 🎯 Objective
Automate the process of **splitting a PDF** by extracting specific pages and saving them into a new PDF file.  
Readers can use **any PDF file available** for testing.

---

### 🛠️ Steps

1. **Action: Extract PDF file pages to new PDF file**
   - Parameters:
     - `Source PDF file`: Use any available PDF file (e.g., `C:\PDFs\YourFile.pdf`)
     - `Page range`: Define which pages you want to extract (e.g., 1–3)
     - `Destination file`: Choose the folder and file name where the new PDF will be saved (e.g., `C:\PDFs\ExtractedPages.pdf`)

2. **Log Result**
   - Log message to confirm where the new PDF is saved.

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Extract PDF File Pages to New PDF File
   Source: C:\Path\To\YourFile.pdf
   Pages: 1–3
   Destination: C:\Path\To\ExtractedPages.pdf

Log "Extracted pages saved to C:\Path\To\ExtractedPages.pdf"
```

---

## ✅ Expected Output
- A new PDF file is created containing only the specified page range.
- Example: If the original file has 20 pages and you extract pages 1–3, the new file will only contain those 3 pages.
