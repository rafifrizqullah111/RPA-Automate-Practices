## 🔹 Case 02: Transactional Processing with RE Framework

### 🎯 Objective
Extend the RE Framework skeleton into a **transaction-based workflow** using Excel as the data source.  
The automation should process transactions row by row, mark their status, and handle errors gracefully.

---

### 🛠️ Steps

1. **Action: Initialize Environment (`Init`)**
   - Create a subflow `Init` to:
     - Open Excel file (e.g., `C:\Data\Transactions.xlsx`).
     - Read data into a datatable → `%TransactionList%`.
     - Initialize index variable `%TransactionIndex% = 0`.
     - Log "Initialization complete".

2. **Action: Get Transaction (`GetTransaction`)**
   - Create subflow `GetTransaction`:
     - If `%TransactionIndex%` < total rows:
       - Extract current row as `%CurrentTransaction%`.
       - Increment `%TransactionIndex%`.
     - Else:
       - Set `%CurrentTransaction% = Null` (no more transactions).
   - Output: `%CurrentTransaction%`.

3. **Action: Process Transaction (`ProcessTransaction`)**
   - Subflow to process one transaction (dummy example: open Notepad, type Customer Name, save file).
   - Log success or failure.

4. **Action: Set Transaction Status (`SetTransactionStatus`)**
   - Update Excel file:
     - If success → write "Success" in `Status` column.
     - If failure → write "Failed" in `Status` column.

5. **Action: End Process (`End Process`)**
   - Save and close Excel.
   - Log "Execution ended".

6. **Action: Main Flow**
   - Sequence:
     - `Init`
     - Loop:
       - `GetTransaction`
       - If `%CurrentTransaction% != Null` → `ProcessTransaction` → `SetTransactionStatus`
       - Else → Break Loop
     - `End Process`

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Main Flow:
   Call → Init
   LOOP
      Call → GetTransaction
      IF %CurrentTransaction% != Null THEN
         TRY
            Call → ProcessTransaction
            Call → SetTransactionStatus (Success)
         CATCH Exception
            Call → SetTransactionStatus (Failed)
      ELSE
         Exit Loop
   END LOOP
   Call → End Process

Subflow: Init
   Log: "Execution started"
   Launch Excel → Open C:\Data\Transactions.xlsx
   Read into %TransactionList%
   Set %TransactionIndex% = 0

Subflow: GetTransaction
   IF %TransactionIndex% < Rows(%TransactionList%)
      %CurrentTransaction% = Row[%TransactionIndex%]
      %TransactionIndex% = %TransactionIndex% + 1
   ELSE
      %CurrentTransaction% = Null

Subflow: ProcessTransaction
   Launch Notepad
   Send Keys: %CurrentTransaction["CustomerName"]%
   Save as C:\Temp\%CurrentTransaction["ID"]%.txt
   Close Notepad

Subflow: SetTransactionStatus
   Write "Success" or "Failed" into Excel cell [Row, "Status"]

Subflow: End Process
   Save and Close Excel
   Log: "Execution ended"
```

---

## ✅ Expected Output
- `Transactions.xlsx` updated with transaction status:

| ID  | CustomerName | Status  |
| --- | ------------ | ------- |
| 001 | John Doe     | Success |
| 002 | Jane Smith   | Failed  |
| 003 | Bob Lee      | Success |

- Text files generated in `C:\Temp\` for each transaction.
- Log messages recorded for each execution stage.
