### English
# Automation fo send Emails due Outlook by VBA
## Description
This Excel macro automates the process of sending emails to customers based on data from the "Customers" worksheet and matching PDF files.

Features:
Collects email addresses from columns C onward.

Displays customer notes from column B before sending.

Uses a .msg email template.

Attaches a matching PDF file per customer.

Moves the sent PDF file to the sent folder.

Automatically skips customers without a PDF file.

How to Use
Open the Excel file and run the SendEmails macro.

Make sure a file named maintenance.msg exists in the same folder.

Ensure that PDF files are named exactly like the customers (e.g., John Doe.pdf).

Make sure Outlook is open and configured for sending.

After sending, each PDF is moved to the sent folder.

Requirements:
Microsoft Excel (macro-enabled).

Microsoft Outlook.

Permissions to send emails and move files on your computer.

### עברית
# אוטומציה לשליחת אימיילים על ידי Outlook באמצעות VBA
## תיאור
המאקרו בקובץ Excel זה מאפשר שליחה אוטומטית של מיילים ללקוחות, על סמך נתונים מגיליון "Customers" וקובצי PDF תואמים.

תכונות:
שליפת כתובות אימייל עבור כל לקוח מעמודות C והלאה.

בדיקת הערות ללקוח בעמודה B והצגת הודעה לפני שליחה.

שימוש בקובץ .msg כתבנית מייל.

צירוף קובץ PDF תואם לפי שם הלקוח.

העברת הקובץ שנשלח לתיקיית sent.

דילוג אוטומטי על לקוחות ללא קובץ PDF.

הוראות שימוש
פתח את הקובץ Excel והפעל את המאקרו SendEmails.

ודא שקובץ .msg בשם maintenance.msg נמצא בתיקיית הקובץ.

ודא שקיימים קובצי PDF בשם זהה לשם הלקוח (לדוגמה: דני כהן.pdf).

ודא ש-Outlook פתוח ומוגדר לשליחה.

לאחר השליחה, הקבצים שנשלחו יועברו לתיקיית sent.

דרישות:
Microsoft Excel (תומך במאקרו).

Microsoft Outlook.

הרשאות לשליחת דוא"ל והעברת קבצים במחשב.

