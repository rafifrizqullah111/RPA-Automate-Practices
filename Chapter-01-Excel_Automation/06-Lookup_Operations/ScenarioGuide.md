# 📘 Hands-on 06: Lookup Operations in Excel Automation

## 🎯 Objective
Automate **lookup operations** in Excel using Power Automate Desktop.  
This is similar to using Excel's `VLOOKUP` or `INDEX-MATCH`.

---

## 📂 Scenario
You have two Excel files:

1. `EmployeeData.xlsx` → contains basic employee info.  
2. `DepartmentBudget.xlsx` → contains department budget data.  

You want to **match employees with their department budgets** automatically.

---

## 📑 Sample Data

### File 1: EmployeeData.xlsx

| Employee ID | Name   | Department |
|-------------|--------|------------|
| 001         | Alice  | HR         |
| 002         | Bob    | IT         |
| 003         | Carol  | Finance    |
| 004         | David  | Marketing  |
| 005         | Eva    | IT         |
| 006         | Frank  | Sales      |
| 007         | Grace  | HR         |
| 008         | Henry  | Finance    |
| 009         | Irene  | Marketing  |
| 010         | Jack   | Sales      |

---

### File 2: DepartmentBudget.xlsx

| Department | Budget  |
|------------|---------|
| HR         | 20000   |
| IT         | 30000   |
| Finance    | 25000   |
| Marketing  | 22000   |
| Sales      | 18000   |

---

## 🛠 Steps to Implement

1. **Launch Excel (EmployeeData.xlsx)**  
   - Action: `Launch Excel` → open `EmployeeData.xlsx`.

2. **Read Employee Data**  
   - Action: `Read from Excel worksheet` → `%EmployeeData%`.

3. **Launch Excel (DepartmentBudget.xlsx)**  
   - Action: `Launch Excel` → open `DepartmentBudget.xlsx`.

4. **Read Department Budget**  
   - Action: `Read from Excel worksheet` → `%BudgetData%`.

5. **For Each Row in EmployeeData**  
   - Action: `For each %CurrentRow% in %EmployeeData%`.  
   - Get department = `CurrentRow['Department']`.

6. **Lookup Matching Budget**  
   - Use Action: `Find first free row` OR filter with **“Find in data table”** (PAD built-in).  
   - Example:  
     - Action: `Find first occurrence of %Department% in column Department of %BudgetData%`.  
     - Store result in `%FoundRow%`.

7. **Write Budget Back**  
   - Action: `Write to Excel worksheet` → write `%FoundRow['Budget']%` into column D of EmployeeData.xlsx.  

8. **Save as EmployeeData_WithBudget.xlsx**  
   - Action: `Save Excel`.

9. **Close Excel Files**  

---

## ✅ Expected Output (EmployeeData_WithBudget.xlsx)

| Employee ID | Name   | Department | Budget |
|-------------|--------|------------|--------|
| 001         | Alice  | HR         | 20000  |
| 002         | Bob    | IT         | 30000  |
| 003         | Carol  | Finance    | 25000  |
| 004         | David  | Marketing  | 22000  |
| 005         | Eva    | IT         | 30000  |
| 006         | Frank  | Sales      | 18000  |
| 007         | Grace  | HR         | 20000  |
| 008         | Henry  | Finance    | 25000  |
| 009         | Irene  | Marketing  | 22000  |
| 010         | Jack   | Sales      | 18000  |

---

## 💡 Notes
- Cara ini **lebih fleksibel daripada Excel formula** karena lookup bisa digabung dengan kondisi lain.  
- Jika jumlah datanya ribuan, lebih efisien pakai **PAD DataTable actions** ketimbang looping manual.  
- Bisa dikembangkan untuk **multi-criteria lookup** (misalnya berdasarkan Department + Job Title).  

