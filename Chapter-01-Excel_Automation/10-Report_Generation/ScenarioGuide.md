# 📘 Hands-on 10: Report Generation in Excel

## 🎯 Objective
Automate the process of **creating a structured sales report** in Excel:
- Generate summary tables (per Region & Product).
- Insert chart automatically.
- Format report for readability.
- Export final report as Excel & PDF.

---

## 📂 Scenario
After cleaning sales data, the manager requests a **weekly report**:
- Total sales by Region
- Total sales by Product
- A chart showing sales by Region
- Exported to PDF

---

## 📑 Input Data (CleanedSalesData.xlsx)

| Order ID | Region   | Product  | Quantity | Unit Price | Total Sale |
|----------|----------|----------|----------|------------|------------|
| 1001     | North    | Laptop   | 3        | 800        | 2400       |
| 1002     | South    | Laptop   | 2        | 800        | 1600       |
| 1003     | East     | Phone    | 5        | 500        | 2500       |
| 1004     | West     | Tablet   | 0        | 300        | 0          |
| 1005     | North    | Phone    | 6        | 500        | 3000       |
| 1006     | South    | Tablet   | 3        | 300        | 900        |
| 1007     | East     | Laptop   | 1        | 800        | 800        |
| 1008     | West     | Phone    | 2        | 500        | 1000       |
| 1009     | North    | Tablet   | 4        | 300        | 1200       |
| 1010     | East     | Phone    | 7        | 500        | 3500       |

---

## 🛠 Steps to Implement

1. **Launch Excel**
   - Action: `Launch Excel` → open `CleanedSalesData.xlsx`.

2. **Insert PivotTable - Sales by Region**
   - Action: `Run Excel Macro`
   - VBA:
   ```vba
   Sub CreatePivotByRegion()
       Dim ws As Worksheet
       Dim pc As PivotCache
       Dim pt As PivotTable
        
       Set ws = Sheets.Add
       ws.Name = "RegionSummary"
        
       Set pc = ActiveWorkbook.PivotCaches.Create( _
           SourceType:=xlDatabase, _
           SourceData:="Sheet1!A1:F100")
        
       Set pt = pc.CreatePivotTable( _
           TableDestination:=ws.Range("A3"), _
           TableName:="PivotRegion")
        
       With pt
           .PivotFields("Region").Orientation = xlRowField
           .PivotFields("Total Sale").Orientation = xlDataField
           .PivotFields("Total Sale").Function = xlSum
       End With
   End Sub
   ```
3. Insert PivotTable - Sales by Product
   ```vba
   Sub CreatePivotByProduct()
     Dim ws As Worksheet
     Dim pc As PivotCache
     Dim pt As PivotTable
    
     Set ws = Sheets.Add
     ws.Name = "ProductSummary"
    
     Set pc = ActiveWorkbook.PivotCaches.Create( _
         SourceType:=xlDatabase, _
         SourceData:="Sheet1!A1:F100")
    
     Set pt = pc.CreatePivotTable( _
         TableDestination:=ws.Range("A3"), _
         TableName:="PivotProduct")
    
     With pt
         .PivotFields("Product").Orientation = xlRowField
         .PivotFields("Total Sale").Orientation = xlDataField
         .PivotFields("Total Sale").Function = xlSum
     End With
   End Sub
   ```
4. Insert Chart (Sales by Region)
   ```vba
   Sub CreateRegionChart()
       Dim ws As Worksheet
       Set ws = Sheets("RegionSummary")
       
       Dim chtObj As ChartObject
       Set chtObj = ws.ChartObjects.Add(Left:=300, Width:=400, Top:=50, Height:=250)
      
       chtObj.Chart.SetSourceData Source:=ws.Range("A3:B10")
       chtObj.Chart.ChartType = xlColumnClustered
       chtObj.Chart.HasTitle = True
       chtObj.Chart.ChartTitle.Text = "Sales by Region"
   End Sub
   ```
6. Format Report
   - Action: Run Excel Macro (auto bold headers, auto-fit columns).
   ```vba
   Sub FormatSheets()
       Dim ws As Worksheet
       For Each ws In ActiveWorkbook.Sheets
           ws.Rows("3:3").Font.Bold = True
           ws.Columns.AutoFit
       Next ws
   End Sub
   ```
7. Save as Excel
   - Action: `Save Excel` → `Weekly_Sales_Report.xlsx`.
8. Export to PDF
   - VBA:
   ```vba
   Sub ExportReportPDF()
       ActiveSheet.ExportAsFixedFormat _
           Type:=xlTypePDF, _
           Filename:=ThisWorkbook.Path & "\Weekly_Sales_Report.pdf", _
           OpenAfterPublish:=False
   End Sub
   ```
9. Close Excel

---

## ✅ Expected Output

Worksheets
- RegionSummary (pivot + chart)
- ProductSummary (pivot)
- Clean formatting (auto-fit, bold headers)

Files Generated
- `Weekly_Sales_Report.xlsx`
- `Weekly_Sales_Report.pdf`

---

## 💡 Notes
- Bisa diperluas dengan menambahkan trendline (weekly) jika dataset punya tanggal.
- Bisa juga dijadikan automated schedule (misalnya report setiap Senin pagi) lewat Power Automate Cloud.
