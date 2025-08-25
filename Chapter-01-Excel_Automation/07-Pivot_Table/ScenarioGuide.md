# 📘 Hands-on 07: Pivot Table Automation in Excel

## 🎯 Objective
Learn how to **automatically create and refresh Pivot Tables** in Excel using Power Automate Desktop.

---

## 📂 Scenario
You have sales transaction data stored in Excel.  
You want to **summarize total sales per region and product** using a Pivot Table, and then export the result into a report file.

---

## 📑 Sample Data (SalesData.xlsx)

| Order ID | Region   | Product     | Quantity | Unit Price | Total Sale |
|----------|----------|-------------|----------|------------|------------|
| 1001     | North    | Laptop      | 3        | 800        | 2400       |
| 1002     | South    | Laptop      | 2        | 800        | 1600       |
| 1003     | East     | Phone       | 5        | 500        | 2500       |
| 1004     | West     | Tablet      | 4        | 300        | 1200       |
| 1005     | North    | Phone       | 6        | 500        | 3000       |
| 1006     | South    | Tablet      | 3        | 300        | 900        |
| 1007     | East     | Laptop      | 1        | 800        | 800        |
| 1008     | West     | Phone       | 2        | 500        | 1000       |
| 1009     | North    | Tablet      | 4        | 300        | 1200       |
| 1010     | East     | Phone       | 7        | 500        | 3500       |

---

## 🛠 Steps to Implement

1. **Launch Excel (SalesData.xlsx)**  
   - Action: `Launch Excel` → open `SalesData.xlsx`.

2. **Read Sales Data**  
   - Action: `Read from Excel worksheet` → store in `%SalesData%`.

3. **Create Pivot Table**  
   - PAD tidak punya action khusus "Create Pivot Table", jadi gunakan salah satu cara:
     - **Cara A (Recommended):** Use `Run Excel Macro` → panggil VBA yang bikin Pivot Table otomatis.  
     - **Cara B:** Simpan template Excel yang sudah ada Pivot Table, lalu gunakan **`Refresh Pivot Table`** action untuk update data baru.

   👉 Untuk latihan ini, kita pakai **Cara A dengan Macro**.

4. **Insert Macro (VBA)**  
   - Example VBA Macro (stored in Excel workbook):

      ```vba
      Sub CreateSalesPivot()
          Dim ws As Worksheet
          Dim pc As PivotCache
          Dim pt As PivotTable
          Dim rng As Range
          
          ' Define source data
          Set ws = ThisWorkbook.Sheets("Sheet1")
          Set rng = ws.Range("A1").CurrentRegion
          
          ' Create Pivot Cache
          Set pc = ThisWorkbook.PivotCaches.Create( _
              SourceType:=xlDatabase, _
              SourceData:=rng)
          
          ' Add new sheet for Pivot
          On Error Resume Next
          Set pt = ThisWorkbook.Sheets("PivotReport").PivotTables("SalesPivot")
          On Error GoTo 0
          
          If pt Is Nothing Then
              ThisWorkbook.Sheets.Add.Name = "PivotReport"
              Set pt = pc.CreatePivotTable( _
                  TableDestination:="PivotReport!R3C1", _
                  TableName:="SalesPivot")
          End If
          
          ' Configure Pivot
          With pt
              .PivotFields("Region").Orientation = xlRowField
              .PivotFields("Product").Orientation = xlColumnField
              .PivotFields("Total Sale").Orientation = xlDataField
              .PivotFields("Total Sale").Function = xlSum
          End With
      End Sub
      ```
5. Run Macro in PAD
   - Action: Run `Excel Macro` → Macro name = `CreateSalesPivot`.
6. Export Pivot Result
   - Action: `Export Excel to PDF` OR` Save Workbook` As `"SalesReport.xlsx"`.
7. Close Excel

---

## ✅ Expected Output (PivotReport)

| Region    | Laptop | Phone | Tablet | Grand Total |
| --------- | ------ | ----- | ------ | ----------- |
| North     | 2400   | 3000  | 1200   | 6600        |
| South     | 1600   | 0     | 900    | 2500        |
| East      | 800    | 6000  | 0      | 6800        |
| West      | 0      | 1000  | 1200   | 2200        |
| **Total** | 4800   | 10000 | 3300   | 18100       |

---

## 💡 Notes
- Dengan cara ini, PAD tidak perlu bikin pivot table manual → cukup sekali bikin macro/template, lalu bisa otomatis di-trigger.
- Bisa dikembangkan jadi multiple pivot reports dalam satu flow.
- Kalau mau cepat, pakai cara template (Pivot Table sudah ada, tinggal Refresh saja).
