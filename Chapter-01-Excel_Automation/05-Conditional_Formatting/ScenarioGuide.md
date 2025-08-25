# 📘 Hands-on 05: Conditional Formatting in Excel Automation

## 🎯 Objective
Use **Power Automate Desktop** to apply conditional formatting rules in Excel automatically.

---

## 📂 Scenario
You have `EmployeeData.xlsx`.  
Manager wants:
1. Highlight salaries **below 5000** in **red background**.  
2. Highlight salaries **above 7000** in **green background**.  

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

2. **Read Data**  
   - Action: `Read from Excel worksheet` → Range: `A2:D11`.  
   - Output: `%EmployeeData%`.

3. **Loop Through Rows**  
   - Action: `For each %CurrentRow% in %EmployeeData%`.  
   - Extract Salary = `CurrentRow['Salary']`.

4. **Check Salary Condition**  
   - Action: `If Salary < 5000` →  
     - Use: `Set cell color` → Red background (`RGB(255,0,0)`) for the salary cell.  
   - Action: `Else If Salary > 7000` →  
     - Use: `Set cell color` → Green background (`RGB(0,255,0)`).

   *(You can directly use `Set cell color` action without formulas.)*

5. **Save File as `EmployeeData_Formatted.xlsx`**  
   - Action: `Save Excel`.

6. **Close Excel**  

---

## ✅ Expected Output (EmployeeData_Formatted.xlsx)

- **Frank (4800)** → highlighted Red.  
- **Jack (4700)** → highlighted Red.  
- **Henry (7200)** → highlighted Green.  
- Others → no highlight.

---

## 💡 Notes
- You can extend rules (e.g., different colors for departments, highlight top 3 salaries).  
- Instead of looping, PAD supports **Excel Advanced Actions** → `Set cell format` with formula-based conditions, but looping is more explicit for learning.  

