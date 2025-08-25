# 📘 Hands-on 01: Reading and Writing Excel Data

## 🎯 Objective
Learn how to **open an Excel file**, **read data** from specific cells or ranges, and **write new values** into Excel using Power Automate Desktop (PAD).

---

## 📂 Scenario
You have an Excel file `EmployeeData.xlsx` that contains a list of employees with the following columns:
- Employee ID
- Name
- Department
- Salary

Your task is to:
1. Open the Excel file.
2. Read the **first 10 rows** of data from the sheet.
3. Write a new column **Bonus** (10% of Salary) into the Excel file.

---

## 📑 Sample Data (EmployeeData.xlsx)

| Employee ID | Name       | Department | Salary |
|-------------|-----------|------------|--------|
| 001         | Alice     | HR         | 5000   |
| 002         | Bob       | IT         | 6000   |
| 003         | Carol     | Finance    | 7000   |
| 004         | David     | Marketing  | 5500   |
| 005         | Eva       | IT         | 6200   |
| 006         | Frank     | Sales      | 4800   |
| 007         | Grace     | HR         | 5100   |
| 008         | Henry     | Finance    | 7200   |
| 009         | Irene     | Marketing  | 5300   |
| 010         | Jack      | Sales      | 4700   |

---

## 🛠 Steps to Implement
1. **Launch Excel**  
   - Action: `Launch Excel` → open `EmployeeData.xlsx`.  
   - Parameter: Make instance visible.

2. **Read Data from Worksheet**  
   - Action: `Read from Excel worksheet`  
   - Range: `A2:D11` (covers 10 rows).  
   - Output: List variable (e.g., `%EmployeeData%`).

3. **Calculate Bonus in PAD**  
   - Action: `For Each` → iterate through `%EmployeeData%`.  
   - Inside loop: `Set Variable` → `Bonus = Salary * 0.1`.

4. **Write Data Back into Excel**  
   - Action: `Write to Excel worksheet`  
   - Target Cell: Column E (starting from `E2`)  
   - Value: `%Bonus%`.

5. **Save and Close Excel**  
   - Action: `Save Excel` → then `Close Excel`.

---

## ✅ Expected Output
- Excel file now has a new column **Bonus** with calculated values.  
- Example:

| Employee ID | Name       | Department | Salary | Bonus |
|-------------|-----------|------------|--------|-------|
| 001         | Alice     | HR         | 5000   | 500   |
| 002         | Bob       | IT         | 6000   | 600   |
| 003         | Carol     | Finance    | 7000   | 700   |
| 004         | David     | Marketing  | 5500   | 550   |
| 005         | Eva       | IT         | 6200   | 620   |
| 006         | Frank     | Sales      | 4800   | 480   |
| 007         | Grace     | HR         | 5100   | 510   |
| 008         | Henry     | Finance    | 7200   | 720   |
| 009         | Irene     | Marketing  | 5300   | 530   |
| 010         | Jack      | Sales      | 4700   | 470   |

---

## 💡 Notes
- Ensure the Excel file is **closed** before running the flow (to avoid lock issues).  
- This hands-on focuses on **basic Excel read/write**, which will be extended in later exercises (filtering, sorting, etc.).
