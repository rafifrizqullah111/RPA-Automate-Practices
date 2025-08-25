# 📘 Hands-on 02: Filtering Excel Data

## 🎯 Objective
Filter data in Excel using **Power Automate Desktop** by applying conditions directly (without looping).

---

## 📂 Scenario
You have `EmployeeData.xlsx`.  
Your manager wants only employees from **IT department** with a **Salary greater than 5000**.  

You need to:
1. Read the Excel data.  
2. Apply filter using `Filter list`.  
3. Save results to `FilteredEmployees.xlsx`.

---

## 📑 Sample Data (EmployeeData.xlsx)

| Employee ID | Name   | Department | Salary |
|-------------|--------|------------|--------|
| 001         | Alice  | HR         | 5000   |
| 002         | Bob    | IT         | 6000   |
| 003         | Carol  | Finance    | 7000   |
| 004         | David  | Marketing  | 5500   |
| 005         | Eva    | IT         | 6200   |
| 006         | Frank  | Sales      | 4800   |
| 007         | Grace  | HR         | 5100   |
| 008         | Henry  | Finance    | 7200   |
| 009         | Irene  | Marketing  | 5300   |
| 010         | Jack   | Sales      | 4700   |

---

## 🛠 Steps to Implement
1. **Launch Excel**  
   - Action: `Launch Excel` → open `EmployeeData.xlsx`.

2. **Read Data from Worksheet**  
   - Action: `Read from Excel worksheet`  
   - Range: `A2:D11`.  
   - Output: `%EmployeeData%`.

3. **Filter List**  
   - Action: `Filter list`  
   - Input list: `%EmployeeData%`  
   - Condition:  
     - `item['Department'] == "IT"`  
     - `item['Salary'] > 5000`  
   - Output: `%FilteredEmployees%`.

4. **Write to New Excel File**  
   - Action: `Launch Excel` (new, blank workbook).  
   - Action: `Write to Excel worksheet` → write `%FilteredEmployees%` starting `A1`.  
   - Action: `Save Excel` → `FilteredEmployees.xlsx`.

5. **Close Excel**  
   - Close both instances.

---

## ✅ Expected Output (FilteredEmployees.xlsx)

| Employee ID | Name | Department | Salary |
|-------------|------|------------|--------|
| 002         | Bob  | IT         | 6000   |
| 005         | Eva  | IT         | 6200   |

---

## 💡 Notes
- Using `Filter list` is simpler and faster than looping with `For Each`.  
- You can combine multiple conditions with `AND` / `OR` logic.  
- This approach is more efficient for medium/large datasets.
