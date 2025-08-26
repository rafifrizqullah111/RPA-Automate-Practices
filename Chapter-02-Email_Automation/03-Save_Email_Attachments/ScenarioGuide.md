## 🔹 Case 03: Save Email Attachments

### 🎯 Objective
Automate the process of **downloading and saving email attachments** from Outlook messages into a specific folder.  
This ensures all important files are collected and stored for further processing or archival.

---

### 🛠️ Steps

1. **Action: Launch Outlook**
   - Make sure Outlook is running.  
   - Save instance to variable: `%OutlookInstance%`.

2. **Action: Retrieve email messages from Outlook**
   - Parameters:
     - `Account`: Default or specify account (e.g., user@company.com).
     - `Folder`: Inbox (or custom folder like "Reports").
     - `Retrieve attachments`: ✔️  
     - `Retrieve unread only`: (optional).
     - `Number of emails to retrieve`: e.g., 5 latest.
   - Output: List of email messages → `%RetrievedEmails%`.

3. **Loop Through Messages**
   - `For each %Email% in %RetrievedEmails%`:
     - **Action: Save Email Attachments**
       - Parameters:
         - `Email`: `%Email%`
         - `Save to folder`: e.g., `C:\Attachments\`
       - Output: List of file paths of saved attachments → `%SavedFiles%`.
     - (Optional) Log `%SavedFiles%`.

4. **Close Outlook (Optional)**
   - Action: `Close Outlook` with `%OutlookInstance%`.

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Launch Outlook → Save as %OutlookInstance%

Retrieve Email Messages
  Account: default
  Folder: Inbox
  Retrieve attachments: True
  Top: 5
  → Save as %RetrievedEmails%

For Each %Email% in %RetrievedEmails%
   Save Email Attachments
      Email: %Email%
      Save to: "C:\Attachments\"
      → %SavedFiles%
   Log "Saved: %SavedFiles%"

Close Outlook (%OutlookInstance%)
```

---

## ✅ Expected Output
- The bot retrieves the 5 most recent emails from Inbox.
- Downloads all attachments and saves them into C:\Attachments\.
- Logs the file paths of saved attachments for confirmation.
- Outlook session is closed at the end.
