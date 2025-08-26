## 🔹 Case 03: Extract Images from PDF and Save to Folder

### 🎯 Objective
Automate the process of **extracting images embedded inside PDF files** (e.g., logos, scanned documents, product pictures) and saving them as image files (PNG/JPG) in a folder for further use.

---

### 🛠️ Steps

1. **Action: Extract images from PDF**
   - Parameters:
     - `PDF file path`: `C:\PDFs\Brochure.pdf`
     - `Save to folder`: `C:\PDFs\ExtractedImages\`
   - Output: List of saved image file paths → `%ExtractedImages%`

2. **Loop Through Extracted Images (Optional)**
   - `For each %ImagePath% in %ExtractedImages%`:
     - Log image file path.
     - (Optional) Insert into Excel for catalog/reporting.

3. **Action: Write image paths to Excel (Optional)**
   - Open/create Excel file `C:\PDFs\ImageCatalog.xlsx`.
   - Write `%ImagePath%` into column A.

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Extract Images from PDF
   File: C:\PDFs\Brochure.pdf
   Save to: C:\PDFs\ExtractedImages\
   → %ExtractedImages%

Launch Excel → %ExcelInstance%
Create New Workbook

Set %Row% = 1
For Each %ImagePath% in %ExtractedImages%
   Write to Excel
      Cell: A%Row%
      Value: %ImagePath%
   Increment %Row%

Save Workbook as C:\PDFs\ImageCatalog.xlsx
Close Excel
```

---

## ✅ Expected Output
- All images from Brochure.pdf are saved in `C:\PDFs\ExtractedImages\`.
- An Excel file ImageCatalog.xlsx contains the list of saved image file paths:

| Image Path                         |
| ---------------------------------- |
| C:\PDFs\ExtractedImages\Image1.png |
| C:\PDFs\ExtractedImages\Image2.png |
| C:\PDFs\ExtractedImages\Image3.png |
