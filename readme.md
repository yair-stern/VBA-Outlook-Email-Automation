##### English
# Automating Email Sending via Outlook using VBA and Excel
### Description
This Excel macro automates the process of sending emails to customers based on data from the "Customers" worksheet and matching PDF files.

### Features:
Collects email addresses from columns C onward.

Displays customer notes from column B before sending.

Uses a .msg email template.

Attaches a matching PDF file per customer.

Moves the sent PDF file to the sent folder.

Automatically skips customers without a PDF file.

### About
This program was originally written as a homework project in an automation course.
After some adjustments, it was used in my workplace (as a back office representative at HOT) as an automatic tool to send batches of emails to customers.

The script reads customer data from an Excel sheet, and matches each name with a corresponding PDF file stored in a folder.
Each customer who appears in the Excel file and has a matching PDF will receive a personalized email with the file attached.

If a customer has a note in the Excel file, a pop-up message will appear showing the note, and the user can choose whether to send the email or skip that customer.
### How to Use
Open the Excel file and run the SendEmails macro.

Make sure a file named maintenance.msg exists in the same folder.

Ensure that PDF files are named exactly like the customers (e.g., John Doe.pdf).

Make sure Outlook is open and configured for sending.

After sending, each PDF is moved to the sent folder.

### Requirements:
Microsoft Excel (macro-enabled).

Microsoft Outlook.

Permissions to send emails and move files on your computer.

##### עברית
# אוטומציה לשליחת מיילים מ־Outlook באמצעות VBA
### תיאור
המאקרו בקובץ Excel זה מאפשר שליחה אוטומטית של מיילים ללקוחות, על סמך נתונים מגיליון "Customers" וקובצי PDF תואמים.

### אודות
התוכנית נכתבה במקור כשיעורי בית בלימודי אוטומציה.
לאחר התאמות, שימשה במקום עבודתי (כנציג בק אופיס בהוט) ככלי אוטומטי לשליחת קבוצות מיילים ללקוחות.

התוכנית קוראת רשימת לקוחות מקובץ Excel, ומשתמשת בתיקייה המכילה קובצי PDF מותאמים אישית לפי שם הלקוח.
כל לקוח שמופיע גם באקסל וגם יש עבורו קובץ PDF, יקבל מייל אוטומטי עם הקובץ המצורף.

במקרה שלקוח מופיע עם הערה מיוחדת ברשימה – תופיע הודעה קופצת עם פרטי ההערה, והמשתמש יוכל לבחור האם לשלוח את המייל או לדלג עליו.

### הוראות הורדה והכנה
**דרישות**
Microsoft Excel (תומך במאקרו).
Microsoft Outlook.
הרשאות לשליחת דוא"ל והעברת קבצים במחשב.

**הורדה**
בטרמינל
git clone https://github.com/yair-stern/VBA-Outlook-Email-Automation.git

**התקנה**
ככלל, התקנה אינה נדרשת.
יש להכין קובץ template.msg
כלומר, קובץ תבנית הודעת מייל של Outlook
לשם הנוחות ניתן ליצור קובץ זה באמצעות תכנית פייתון בפרויקט זה.

תוכנית זו דורשת:
דרישות קדם:
פייתון מותקן
מכונה לסביבה וירטואלית virtualenv מותקן
אפליקציית Ountlook desktop

התקנה נוספת:
ספריית win32
רצוי להשתמש בסביבה וירטואלית

הוראות צעד אחר צעד:
נווט לתקית
python_for_auto_mail
פתח את app.py
ערוך את תבנית המייל שבקובץ ושמור

בטרמינל
python -m virtualenv env
env\Scripts\activate

pip install pywin32
או
pip install -r requirements.txt

ולאחר מכן בטרמינל
python python_for_auto_mail\app.py

גרור את הקובץ
your_template_file_name.msg (או את השם שנתת לו)
לתיקית
VBA-Outlook-Email-Automation
או אם נתת לה שם אחר, לתיקית קובץ המאקרו
Auto_Mails.xlsm
אם לא נתת לו שם אחר

**הכנה**
*לפני הפעם הראשונה בלבד*
אם לא יצרת תבנית אימייל (הוראות תחת כותרת ההתקנה), צור קובץ Outlook Email, כלומר קובץ מסוג .msg ושים אותו באותה תיקיה של קובץ האקסל
קובץ זה צריך להיות בעל השם
"your_template_file_name.msg"
אם נתת לו שם אחר, שנה בהתאם את הסקריפט ב-VBA שבקובץ האקסל בשורה 30:
    templatePath = ThisWorkbook.Path & "\your_template_file_name.msg"

בעמודה A בקובץ ה-אקסל רשום את שמות כל הלקוחות שלך
בעמודה B רשום הערות עבור לקוחות שאתה מעוניין לאשר אחד אחד לפני שאתה שולח להם מיילים.
בעמודות C והלאה רשום לכל לקוח את כתובות האימייל שלו
חובה לרשום לכל לקוח שם אחד לפחות אחרת תיתקל בשגיאה בנסיון לשלוח לו הודעה והתכנית לא תיתן מידע לגבי השגיאה
### תכונות:
שליפת כתובות אימייל עבור כל לקוח מעמודות C והלאה.

בדיקת הערות ללקוח בעמודה B והצגת הודעה לפני שליחה.

שימוש בקובץ .msg כתבנית מייל.

צירוף קובץ PDF תואם לפי שם הלקוח.

העברת הקובץ שנשלח לתיקיית sent.

דילוג אוטומטי על לקוחות ללא קובץ PDF.

### הוראות שימוש
פתח את הקובץ Excel והפעל את המאקרו SendEmails.

ודא שקובץ .msg בשם maintenance.msg נמצא בתיקיית הקובץ.

ודא שקיימים קובצי PDF בשם זהה לשם הלקוח (לדוגמה: דני כהן.pdf).

ודא ש-Outlook פתוח ומוגדר לשליחה.

לאחר השליחה, הקבצים שנשלחו יועברו לתיקיית sent.

### סיבוכיות
תוכנית זו מיועדת לשימוש של כמות סבירה של מיילים (מאות)
מתוך רשימה סבירה של אנשי קשר (עשרות אלפים)
משום כך לא תוכננה ביעילות כקוד Real Time
לפיכך סיבוכיות זמן ריצה הינה מסדר M*N (O(m*n)) כאשר N מייצג את מספר הלקוחות שישנם ברשימה באקסל ו-M את מספר קבצי ה-PDF, כלומר המיילים שיש לשלוח כעת
לשם ייעול לריצה לינארית יש לבצע מיון לשתי הרשימות לפני תחילת החיפושים, ולחפש עבור כל קובץ PDF את שם הלקוח בקובץ האקסל מאותה נקודה שעצרנו עבור הלקוח המתאים לקובץ ה-PDF הקודם שנשלח.

### שיפורים עתידיים:
שיפור הסיבוכיות:
לאפשר בחירה של סיבוכיות לינארית עבור כמויות גדולות
שימוש בקובץ משתני סביבה כדי להקל על המשתמש לשנות שמות כמו עמודות לבדיקה, הערות, שמות הקבצים והתקיות שלו

