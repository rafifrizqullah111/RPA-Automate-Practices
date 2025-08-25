# 📘 Hands-on 03: Sorting Excel Data

## 🎯 Objective
Sort Excel data automatically using **Power Automate Desktop** with `Sort list` action.

---

## 📂 Scenario
You have `EmployeeData.xlsx`.  
Your manager wants to:
1. Sort employees by **Department** (A → Z).  
2. Within each department, sort by **Salary** (highest first).  

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

3. **Sort List - First by Department**  
   - Action: `Sort list`  
   - Input: `%EmployeeData%`  
   - Key: `item['Department']`  
   - Order: `Ascending`.  
   - Output: `%SortedByDept%`.

4. **Sort List - Then by Salary (within Department)**  
   - Action: `Sort list`  
   - Input: `%SortedByDept%`  
   - Key: `item['Salary']`  
   - Order: `Descending`.  
   - Output: `%FinalSortedList%`.

5. **Write to New Excel File**  
   - Action: `Launch Excel` (new workbook).  
   - Action: `Write to Excel worksheet` → write `%FinalSortedList%` starting `A1`.  
   - Action: `Save Excel` → `SortedEmployees.xlsx`.

6. **Close Excel**  
   - Close both instances.

---

## ✅ Expected Output (SortedEmployees.xlsx)

| Employee ID | Name   | Department | Salary |
|-------------|--------|------------|--------|
| 003         | Carol  | Finance    | 7000   |
| 008         | Henry  | Finance    | 7200   |
| 007         | Grace  | HR         | 5100   |
| 001         | Alice  | HR         | 5000   |
| 005         | Eva    | IT         | 6200   |
| 002         | Bob    | IT         | 6000   |
| 004         | David  | Marketing  | 5500   |
| 009         | Irene  | Marketing  | 5300   |
| 006         | Frank  | Sales      | 4800   |
| 010         | Jack   | Sales      | 4700   |

---

## 💡 Notes
- You can apply multiple `Sort list` actions to mimic **multi-level sorting**.  
- For advanced scenarios, combine `Sort list` with `Filter list`.  
- Sorting does not modify the original Excel file until you write back the results.

