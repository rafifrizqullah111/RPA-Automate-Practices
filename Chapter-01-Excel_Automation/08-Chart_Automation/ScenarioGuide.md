# 📘 Hands-on 08: Charts Automation in Excel

## 🎯 Objective
Learn how to **automatically create and update charts** (bar, column, line, etc.) in Excel using Power Automate Desktop.

---

## 📂 Scenario
After creating a **Pivot Table** report (previous exercise), now you want to **visualize sales data per region and product** as a chart, and then export the chart into a report file (PDF/Excel).

---

## 📑 Sample Data (SalesData.xlsx)
(Same as used in Pivot Table Automation)

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

2. **Run Pivot Macro (Optional)**  
   - If Pivot Table already exists, skip.  
   - Else, use Macro `CreateSalesPivot` (from Hands-on 07).

3. **Insert Chart Macro (VBA)**  
   - Add macro inside workbook to create chart:
      ```vba
      Sub CreateSalesChart()
          Dim ws As Worksheet
          Dim chartObj As ChartObject
          
          ' Use PivotReport sheet
          Set ws = ThisWorkbook.Sheets("PivotReport")
          
          ' Delete old chart if exists
          For Each chartObj In ws.ChartObjects
              chartObj.Delete
          Next chartObj
          
          ' Add new chart
          Set chartObj = ws.ChartObjects.Add(Left:=300, Top:=20, Width:=500, Height:=300)
          chartObj.Chart.SetSourceData Source:=ws.Range("A3:E7")
          chartObj.Chart.ChartType = xlColumnClustered
          
          ' Add chart title
          chartObj.Chart.HasTitle = True
          chartObj.Chart.ChartTitle.Text = "Sales by Region and Product"
      End Sub
      ```

4. Run Macro in PAD
   - Action: Run `Excel Macro` → Macro name = `CreateSalesChart`.
5. Export Report
   - Action: `Export Excel to PDF` OR `save as "SalesReport_WithChart.xlsx"`.
6. Close Excel

---
 
## ✅ Expected Output (Chart)
- A clustered column chart showing sales breakdown per region & product.
- Example:
  - X-axis = Region
  - Columns = Product categories (Laptop, Phone, Tablet)
  - Y-axis = Total Sales

## 💡 Notes
- You can switch ChartType to:
  - `xlPie` → Pie Chart
  - `xlLine` → Line Chart
  - `xlBarClustered` → Bar Chart
- For automation scale-up:
  - Create multiple charts (sales trend, top products, etc.).
  - Export all into a dashboard PDF.
