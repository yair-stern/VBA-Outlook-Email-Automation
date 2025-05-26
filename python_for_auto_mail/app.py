from win32com.client import Dispatch

outlook = Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)  # 0 = olMailItem

mail.Subject = "החשבונית שלך/תלוש השכר שלך. לכבוד"
mail.Body = """שלום עובד/לקוח יקר

החשבונית/תלוש השכר שלך לחודש זה נמצא בקובץ המצורף למייל זה.

בברכה,
יאיר שטרן

-- 
יאיר ש.
סמנכ"ל כספים | Stern Enterprises Ltd.
טלפון: 03-1234567 | נייד: 054-7654321
"""
mail.CC = "account@office.com"

mail.SaveAs("your_template_file_name.msg", 3) # 3 = olMsg
