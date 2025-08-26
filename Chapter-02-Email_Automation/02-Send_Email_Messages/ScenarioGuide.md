## 🔹 Case 03: Send Email Message

### 🎯 Objective
Automate the process of **sending email messages via Outlook**, including subject, body, recipients, and optional attachments.

---

### 🛠️ Steps

1. **Action: Launch Outlook**
   - Make sure Outlook is running.  
   - Save instance to variable: `%OutlookInstance%`.

2. **Action: Send email message through Outlook**
   - Parameters:
     - `Account`: Default or specific account (e.g., user@company.com).
     - `To`: test.receiver@company.com
     - `Cc`: (optional) manager@company.com
     - `Bcc`: (optional)
     - `Subject`: "Monthly Sales Report"
     - `Body`: "Hi Team, please find attached the sales report for this month."
     - `Attachments`: e.g., `C:\Reports\Sales_July.xlsx`
   - Output: (no output, just confirmation the email was sent).

3. **(Optional) Retrieve sent item confirmation**
   - After sending, retrieve last sent mail from "Sent Items" folder to verify.

4. **Close Outlook (Optional)**
   - Action: `Close Outlook` with `%OutlookInstance%`.

---

### 📂 Example Workflow (Pseudo-code)
```plaintext
Launch Outlook → Save as %OutlookInstance%

Send Email Message
  Account: default
  To: test.receiver@company.com
  Cc: manager@company.com
  Subject: "Monthly Sales Report"
  Body: "Hi Team, please find attached the sales report for this month."
  Attachment: C:\Reports\Sales_July.xlsx

Close Outlook (%OutlookInstance%)
```

---

## ✅ Expected Output
- Outlook sends an email to test.receiver@company.com.
- Email subject: "Monthly Sales Report".
- Email body includes the provided text.
- Excel file Sales_July.xlsx is attached.
