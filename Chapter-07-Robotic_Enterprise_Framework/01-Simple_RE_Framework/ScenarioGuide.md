## 🔹 Case 01: Build a Simple RE Framework

### 🎯 Objective
Create a **lightweight Robotic Enterprise (RE) Framework** in Power Automate Desktop, focusing on structuring automation into clear stages (`Init`, `Process Transaction`, `End Process`).  
This case helps establish the foundation for scalable automations before moving into transactional design.

---

### 🛠️ Steps

1. **Action: Initialize Environment (`Init`)**
   - Create a subflow `Init` to:
     - Log the start of execution.
     - Initialize variables (e.g., `%LogFilePath%`, `%TransactionCounter%`).
     - Prepare working directories (optional).
   - Output: Ready-to-use environment.

2. **Action: Main Process (`Process Transaction`)**
   - Create a subflow `Process Transaction` to:
     - Execute a dummy process (e.g., open Notepad, type text, save file).
     - Log process status (`Success` or `Failed`).
   - Output: `%TransactionResult%`.

3. **Action: End Process (`End Process`)**
   - Create a subflow `End Process` to:
     - Log that the framework has ended.
     - Close any open applications/resources.
   - Output: Clean termination.

4. **Action: Main Flow**
   - Call subflows in sequence:
     - `Init`
     - `Process Transaction`
     - `End Process`

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Main Flow:
   Call → Init
   Call → Process Transaction
   Call → End Process

Subflow: Init
   Log Message: "Execution started"
   Set Variable: %TransactionCounter% = 0
   Set Variable: %LogFilePath% = "C:\Logs\PAD_REFramework.log"

Subflow: Process Transaction
   Try:
      Launch Notepad
      Send Keys: "Hello, this is RE Framework skeleton"
      Save File as C:\Temp\DummyOutput.txt
      Log Message: "Transaction processed successfully"
   Catch Exception:
      Log Message: "Transaction failed"

Subflow: End Process
   Log Message: "Execution ended"
   Close Notepad (if open)
```

---

## ✅ Expected Output
- A Power Automate Desktop flow project with:
  - Three subflows: Init, Process Transaction, End Process.
  - A main flow calling each subflow in order.
- A log file (or console messages) showing:
  - Start of execution
  - Process transaction result (success/fail)
  - End of execution
- A dummy file created in C:\Temp\DummyOutput.txt containing:

```palintext
Hello, this is RE Framework skeleton
```
