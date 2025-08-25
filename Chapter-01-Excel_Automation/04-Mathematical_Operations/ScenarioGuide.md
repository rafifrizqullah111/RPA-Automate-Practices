# 📘 Hands-on 04: Mathematical Operations in Excel Automation

## 🎯 Objective
Use **Power Automate Desktop** to perform mathematical operations (Sum, Average, Min, Max) on Excel data.

---

## 📂 Scenario
You have `EmployeeData.xlsx`.  
Your manager wants to know:
1. Total salary of all employees.  
2. Average salary per employee.  
3. Highest and lowest salary.  

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

3. **Extract Salary Column Only**  
   - Action: `For each` → iterate `%EmployeeData%`.  
   - Inside loop: `Append to list` → append `CurrentItem['Salary']` into `%SalaryList%`.

   *(Alternative: If data structured, use `Get column from data table` action).*

4. **Calculate Total Salary**  
   - Action: `Calculate` → `Sum list %SalaryList%`.  
   - Output: `%TotalSalary%`.

5. **Calculate Average Salary**  
   - Action: `Calculate` → `%TotalSalary% / List.Count(%SalaryList%)`.  
   - Output: `%AverageSalary%`.

6. **Find Maximum and Minimum Salary**  
   - Action: `Get maximum value of list %SalaryList%` → `%MaxSalary%`.  
   - Action: `Get minimum value of list %SalaryList%` → `%MinSalary%`.

7. **Write Results to New Excel File**  
   - Action: `Launch Excel` (new workbook).  
   - Action: `Write to Excel worksheet` with results:

| Metric            | Value   |
|-------------------|---------|
| Total Salary      | 56,000  |
| Average Salary    | 5,600   |
| Highest Salary    | 7,200   |
| Lowest Salary     | 4,700   |

8. **Save as `SalaryReport.xlsx`**  
   - Action: `Save Excel`.  

9. **Close Excel**  
   - Close all instances.

---

## ✅ Expected Output (SalaryReport.xlsx)

| Metric            | Value |
|-------------------|-------|
| Total Salary      | 56,000 |
| Average Salary    | 5,600 |
| Highest Salary    | 7,200 |
| Lowest Salary     | 4,700 |

---

## 💡 Notes
- For larger datasets, prefer `Get column from data table` to avoid looping.  
- You can extend this by calculating per **Department** (using `Group by` logic).  
- This method allows building **automated reports** without formulas inside Excel.  

