# Outlook & Excel Bulk Email Automation (VBA)

A lightweight, powerful VBA macro designed to send personalized bulk emails directly from Microsoft Excel using Microsoft Outlook—completely free and without requiring third-party tools, paid plugins, or external API keys.

---

## 🚀 Features

* **Multi-Account Support:** Automatically detects all configured accounts in your local Outlook session and prompts you to choose which email address you want to send from.
* **Dynamic Placeholders:** Easily personalize your emails using custom tags like `{{NAME}}` and `{{COMPANY}}` pulled directly from your Excel data.
* **Centralized Templates:** Reads the email subject from a single master cell (`D2`) and the body text template from another (`E2`).
* **Zero Credentials Stored:** Secure by design—it interacts directly with your running local Outlook application and leaves zero passwords or tokens hardcoded in the script.

---

## 📋 Excel Sheet Structure Setup

To ensure the macro runs smoothly, configure your active worksheet (`Sheets(1)`) with the following column layout starting from **Row 2**:

| Column | Header Example | Description |
| :--- | :--- | :--- |
| **Column A** | `Name` | Recipient's name (replaces `{{NAME}}`) |
| **Column B** | `Company` | Recipient's company (replaces `{{COMPANY}}`) |
| **Column C** | `Email` | Recipient's email address |
| **Column D** | `Subject` | Email subject text (referenced in cell `D2`) |
| **Column E** | `Body Template` | Master email body with placeholders (referenced in cell `E2`) |

---

## ⚙️ How to Install and Use

1. Open your Excel workbook containing your recipient list.
2. Press `ALT + F11` to open the Visual Basic for Applications (VBA) editor.
3. In the top menu, go to **Insert > Module** to create a blank code module.
4. Copy and paste the `SendBulkEmails_FromDraft` macro into the module window.
5. In the VBA editor menu, go to **Tools > References** and ensure **Microsoft Outlook x.x Object Library** is checked.
6. Close the VBA editor and save your workbook as an **Excel Macro-Enabled Workbook (`.xlsm`)**.
7. Press `ALT + F8`, select `SendBulkEmails_FromDraft`, and click **Run**.

---

## 🔒 Safety & Best Practice Warning

By default, the script includes the line `.Send` inside the loop, which dispatches emails immediately. 

> **Tip for Testing:** Before running this on a large list of real clients, it is highly recommended to change `.Send` to `.Display` in the code block. This will open each email as a draft on your screen so you can manually verify that placeholders and formatting render correctly before going live.

```vba
With OutMail
    .To = ws.Cells(i, 3).Value
    .Subject = subjectText
    .Body = finalBody
    .SendUsingAccount = OutAccount
    
    ' Change to .Display for review, or .Send to send automatically
    .Display  
End With
