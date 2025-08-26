## 🔹 Case 09: Save Outlook Email Messages

### 🎯 Objective
Automate the process of **saving Outlook email messages as files** (e.g., `.msg`, `.txt`, or `.html`) into a local folder.  
This is useful for archiving, auditing, or sharing raw email content outside of Outlook.

---

### 🛠️ Steps

1. **Action: Launch Outlook**
   - Make sure Outlook is running.  
   - Save instance to variable: `%OutlookInstance%`.

2. **Action: Retrieve email messages from Outlook**
   - Parameters:
     - `Account`: Default or specify account.
     - `Folder`: Inbox (or another folder like "Approvals").
     - `Retrieve unread only`: (optional).
     - `Number of emails to retrieve`: e.g., 5 latest.
   - Output: List of email messages → `%RetrievedEmails%`.

3. **Loop Through Messages**
   - `For each %Email% in %RetrievedEmails%`:
     - **Action: Save Outlook email messages**
       - Parameters:  
         - `Email`: `%Email%`  
         - `Save to folder`: `C:\SavedEmails\`  
         - `Format`: `.msg` (can also choose `.txt` or `.html`).  
       - Output: Saved file path(s) → `%SavedEmailPath%`.  
     - Log `%SavedEmailPath%`.

4. **Close Outlook (Optional)**
   - Action: `Close Outlook` with `%OutlookInstance%`.

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Launch Outlook → Save as %OutlookInstance%

Retrieve Email Messages
  Account: default
  Folder: Inbox
  Top: 5
  → Save as %RetrievedEmails%

For Each %Email% in %RetrievedEmails%
   Save Outlook Email Message
      Email: %Email%
      Save to: "C:\SavedEmails\"
      Format: .msg
      → %SavedEmailPath%
   Log "Saved email at: %SavedEmailPath%"

Close Outlook (%OutlookInstance%)
```

---

## ✅ Expected Output
- The bot retrieves the 5 most recent emails.
- Saves each email as a `.msg` file into `C:\SavedEmails\`.
- File paths of saved emails are logged.
- Outlook session is closed at the end.
