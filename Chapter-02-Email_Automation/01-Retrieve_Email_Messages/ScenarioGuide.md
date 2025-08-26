## 🔹 Case 01: Retrieve Email Messages

### 🎯 Objective
Automate the process of **retrieving emails from Outlook**, optionally with filters (by folder, sender, subject, or time).  
This allows the robot to process incoming messages for further workflows.

---

### 🛠️ Steps

1. **Action: Launch Outlook**
   - Make sure Outlook is running.  
   - Save instance to variable: `%OutlookInstance%`.

2. **Action: Retrieve email messages from Outlook**
   - Parameters:
     - `Account`: Default or specify account (e.g., user@company.com).
     - `Folder`: Inbox (or another, e.g., "Reports").
     - `Retrieve unread only`: ✔️ (optional).
     - `Retrieve attachments`: ✔️ (optional).
     - `Number of emails to retrieve`: e.g., 10 latest.
   - Output: List of email messages → `%RetrievedEmails%`.

3. **Loop Through Messages (Optional)**
   - `For each %Email% in %RetrievedEmails%`:
     - Extract `Sender`, `Subject`, `ReceivedTime`, `Body`, etc.
     - Store into DataTable or log file for later processing.

4. **Close Outlook (Optional)**
   - Action: `Close Outlook` with `%OutlookInstance%`.

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Launch Outlook → Save as %OutlookInstance%

Retrieve Email Messages
  Account: default
  Folder: Inbox
  Retrieve unread only: True
  Top: 10
  → Save as %RetrievedEmails%

For Each %Email% in %RetrievedEmails%
   Log %Email.Subject%
   Log %Email.Sender%
   Log %Email.ReceivedTime%

Close Outlook (%OutlookInstance%)
```

---

## ✅ Expected Output
- The bot retrieves the 10 most recent unread emails from Inbox.
- Logs sender, subject, and received time into `console/log` file.
- Outlook session is closed at the end.
