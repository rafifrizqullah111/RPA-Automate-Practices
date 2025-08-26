## 🔹 Case 07: Organize and Move Emails

### 🎯 Objective
Automate the process of **organizing emails** by moving them into specific folders (e.g., "Reports", "Invoices", "Approvals") based on predefined rules like sender, subject, or attachments.  
This helps keep the mailbox structured and reduces manual sorting.

---

### 🛠️ Steps

1. **Action: Launch Outlook**
   - Make sure Outlook is running.  
   - Save instance to variable: `%OutlookInstance%`.

2. **Action: Retrieve email messages from Outlook**
   - Parameters:
     - `Account`: Default or specify account.
     - `Folder`: Inbox.
     - `Retrieve unread only`: (optional).
     - `Number of emails to retrieve`: e.g., 10 latest.
   - Output: List of email messages → `%RetrievedEmails%`.

3. **Loop Through Messages**
   - `For each %Email% in %RetrievedEmails%`:
     - **Condition**:  
       - If `Subject` contains "Invoice" → Move to "Invoices" folder.  
       - If `Subject` contains "Report" → Move to "Reports" folder.  
       - Else → Move to "General" folder.  
     - **Action: Move Email**  
       - Parameters:  
         - `Email`: `%Email%`  
         - `Destination folder`: e.g., `Inbox\Invoices` or `Inbox\Reports`.

4. **Close Outlook (Optional)**
   - Action: `Close Outlook` with `%OutlookInstance%`.

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Launch Outlook → Save as %OutlookInstance%

Retrieve Email Messages
  Account: default
  Folder: Inbox
  Top: 10
  → Save as %RetrievedEmails%

For Each %Email% in %RetrievedEmails%
   If %Email.Subject% Contains "Invoice"
      Move Email → Folder: Inbox\Invoices
   Else If %Email.Subject% Contains "Report"
      Move Email → Folder: Inbox\Reports
   Else
      Move Email → Folder: Inbox\General

Close Outlook (%OutlookInstance%)
```

---

## ✅ Expected Output
- The bot retrieves the 10 latest emails from Inbox.
- Emails with "Invoice" in subject → moved to Invoices folder.
- Emails with "Report" in subject → moved to Reports folder.
- All other emails → moved to General folder.
- Outlook session is closed at the end.
