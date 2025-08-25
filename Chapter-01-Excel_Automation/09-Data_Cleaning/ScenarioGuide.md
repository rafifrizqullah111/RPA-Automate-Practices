# 📘 Hands-on 09: Data Cleaning in Excel

## 🎯 Objective
Learn how to **automatically clean raw Excel data** using Power Automate Desktop:
- Remove duplicates
- Trim spaces
- Replace null/blank cells
- Standardize formatting

---

## 📂 Scenario
You receive raw sales data from multiple sources. Before analysis, the data must be cleaned to ensure accuracy.  

PAD will automate cleaning operations and save a "Cleaned" version of the file.

---

## 📑 Sample Raw Data (RawSalesData.xlsx)

| Order ID | Region     | Product     | Quantity | Unit Price | Total Sale |
|----------|------------|-------------|----------|------------|------------|
| 1001     | North      | Laptop      | 3        | 800        | 2400       |
| 1002     | South      | Laptop      | 2        | 800        | 1600       |
| 1002     | South      | Laptop      | 2        | 800        | 1600       |  ← duplicate
| 1003     | EAST       | Phone       | 5        | 500        | 2500       |
| 1004     | West       | Tablet      |          | 300        |            |  ← missing Quantity & Total
| 1005     |  North     | Phone       | 6        | 500        | 3000       |  ← extra spaces in Region
| 1006     | South      | tablet      | 3        | 300        | 900        |  ← lowercase Product
| 1007     | East       | Laptop      | 1        | 800        | 800        |
| 1008     | West       | Phone       | 2        | 500        | 1000       |
| 1009     | North      | Tablet      | 4        | 300        | 1200       |
| 1010     | East       | Phone       | 7        | 500        | 3500       |

---

## 🛠 Steps to Implement

1. **Launch Excel**
   - Action: `Launch Excel` → open `RawSalesData.xlsx`.

2. **Remove Duplicates**
   - Action: `Run Excel Macro`
   - VBA code:
      ```vba
      Sub RemoveDuplicates()
          Sheets("Sheet1").Range("A1:F100").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6), Header:=xlYes
      End Sub
      ```
3. Trim Spaces in Text Fields
   - VBA code:
     ```vba
     Sub TrimSpaces()
         Dim rng As Range
         For Each rng In Sheets("Sheet1").Range("B2:B100") ' Region column
             If Not IsEmpty(rng.Value) Then rng.Value = Trim(rng.Value)
         Next rng
     End Sub
     ```
4. Fix Case (Proper Formatting)
   - VBA code:
     ```vba
     Sub StandardizeProductNames()
         Dim rng As Range
         For Each rng In Sheets("Sheet1").Range("C2:C100") ' Product column
             If Not IsEmpty(rng.Value) Then rng.Value = StrConv(rng.Value, vbProperCase)
         Next rng
     End Sub
     ```
5. Replace Blank Cells with Default Values
   - VBA code:
     ```vba
     Sub ReplaceBlanks()
         Dim ws As Worksheet
         Set ws = Sheets("Sheet1")
         
         ws.Range("D2:D100").SpecialCells(xlCellTypeBlanks).Value = 0 ' Quantity
         ws.Range("F2:F100").SpecialCells(xlCellTypeBlanks).Value = 0 ' Total Sale
     End Sub
     ```
6. Recalculate Total Sale if Missing
   - VBA code:
     ```vba
     Sub RecalculateTotal()
         Dim i As Long
         With Sheets("Sheet1")
             For i = 2 To 100
                 If .Cells(i, 6).Value = 0 And .Cells(i, 4).Value > 0 Then
                     .Cells(i, 6).Value = .Cells(i, 4).Value * .Cells(i, 5).Value
                 End If
             Next i
         End With
     End Sub
     ```
7. Save Cleaned File
   - Action: `Save Excel` → as `CleanedSalesData.xlsx`.
8. Close Excel

---

## ✅ Expected Output (Cleaned Data)

| Order ID | Region | Product | Quantity | Unit Price | Total Sale |
| -------- | ------ | ------- | -------- | ---------- | ---------- |
| 1001     | North  | Laptop  | 3        | 800        | 2400       |
| 1002     | South  | Laptop  | 2        | 800        | 1600       |
| 1003     | East   | Phone   | 5        | 500        | 2500       |
| 1004     | West   | Tablet  | 0        | 300        | 0          |
| 1005     | North  | Phone   | 6        | 500        | 3000       |
| 1006     | South  | Tablet  | 3        | 300        | 900        |
| 1007     | East   | Laptop  | 1        | 800        | 800        |
| 1008     | West   | Phone   | 2        | 500        | 1000       |
| 1009     | North  | Tablet  | 4        | 300        | 1200       |
| 1010     | East   | Phone   | 7        | 500        | 3500       |

---

## 💡 Notes
- Bisa digabung semua macro jadi 1 Master Macro biar PAD cukup sekali panggil.
- Untuk dataset besar, bisa pakai `For Each Row in Excel Worksheet` action tanpa VBA (tapi lebih lambat).
- Hasil bisa langsung dipakai untuk analisis/pivot.
