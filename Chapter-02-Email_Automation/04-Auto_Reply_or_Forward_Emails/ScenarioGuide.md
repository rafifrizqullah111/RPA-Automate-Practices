## 🔹 Case 04: Auto Reply or Forward Emails

### 🎯 Objective
Automate the process of **replying** or **forwarding emails** in Outlook based on predefined rules (e.g., sender, subject, or keywords).  
This helps reduce manual effort in responding to common queries or escalating messages.

---

### 🛠️ Steps

1. **Action: Launch Outlook**
   - Make sure Outlook is running.  
   - Save instance to variable: `%OutlookInstance%`.

2. **Action: Retrieve email messages from Outlook**
   - Parameters:
     - `Account`: Default or specify account.
     - `Folder`: Inbox (or another, e.g., "Support").
     - `Retrieve unread only`: ✔️ (optional).
     - `Number of emails to retrieve`: e.g., 5 latest.
   - Output: List of email messages → `%RetrievedEmails%`.

3. **Loop Through Messages**
   - `For each %Email% in %RetrievedEmails%`:
     - **Condition (Optional)**: Check if `Subject` contains keyword (e.g., "Request") or `Sender` matches specific address.
     - **Action: Reply to email**  
       - Parameters:  
         - `OriginalEmail`: `%Email%`  
         - `Body`: "Thank you for your message. We have received your request and will process it shortly."  
         - `Include original message`: ✔️ (optional).  
     - OR  
     - **Action: Forward email**  
       - Parameters:  
         - `OriginalEmail`: `%Email%`  
         - `Recipients`: `manager@company.com`  
         - `Body`: "Forwarding this email for your attention."  

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
  Top: 5
  → Save as %RetrievedEmails%

For Each %Email% in %RetrievedEmails%
   If %Email.Subject% Contains "Request"
      Reply to Email
         OriginalEmail: %Email%
         Body: "Thanks, your request is being processed."
         Include original: True
   Else
      Forward Email
         OriginalEmail: %Email%
         Recipients: manager@company.com
         Body: "Forwarding this email for your review."

Close Outlook (%OutlookInstance%)
```

---

## ✅ Expected Output
- The bot retrieves the 5 most recent unread emails.
- If subject contains `"Request"`, it replies automatically with a predefined template.
- If not, the email is forwarded to a specified recipient.
- Outlook session is closed at the end.
