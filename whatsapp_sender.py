import time
import urllib.parse
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ✅ ANSI color codes for colored terminal output
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
CYAN = "\033[96m"
RESET = "\033[0m"

# ✅ تحميل بيانات الأرقام والرسائل من ملف Excel
df = pd.read_excel("contacts.xlsx")  # يجب أن يحتوي على الأعمدة: ClientName, PhoneNumber, Message

# ✅ إعداد WebDriver لمتصفح Chrome
driver = webdriver.Chrome()  # تأكد من أن chromedriver في نفس مجلد المشروع

# ✅ فتح واتساب ويب
driver.get("https://web.whatsapp.com/") # يجب تسجيل الدخول يدويًا

# ✅ عرض اسم المطوّر وحقوقه في البداية
print(f"{CYAN}----------------------------------------------{RESET}")
print(f"{YELLOW}🚀 Program created by: {GREEN}Yaser Mohammed Alshikh{RESET}")
print(f"{YELLOW}✉ Email: {CYAN}yaseralshikh@gmail.com{RESET}")
print(f"{YELLOW}🔒 Protected and owned. {RESET}📌 Contact for improvements or licensing.")
print(f"{YELLOW}© All rights reserved. {RESET}💡 Idea and implementation by Yaser.")
print(f"{CYAN}----------------------------------------------\n{RESET}")

input("🔵 After logging into WhatsApp, press Enter to continue...")

# ✅ قائمة لتخزين الأرقام التي فشل إرسال الرسائل إليها
failed_numbers = []

# ✅ إرسال الرسائل للأرقام الموجودة في الملف
for index, row in df.iterrows():
    client_name = row["ClientName"]
    number = str(row["PhoneNumber"])  # تحويل الرقم إلى نص لضمان عدم حدوث أخطاء
    message = row["Message"]

    # ✅ دمج اسم العميل في الرسالة
    full_message = f"مرحبًا {client_name}، {message}"
    encoded_message = urllib.parse.quote(full_message)  # ✅ ترميز الرسالة لتجنب الأخطاء
    url = f"https://web.whatsapp.com/send?phone={number}&text={encoded_message}"
    
    driver.get(url)
    
    try:
        # ✅ انتظار تحميل الصفحة بالكامل قبل البحث عن زر الإرسال
        send_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, '//button[contains(@aria-label, "إرسال")]'))
        )
        send_button.click()
        time.sleep(10)  # انتظار بسيط للتأكد من الإرسال
        print(f"✅ Message sent to {client_name} ({number})")

    except Exception as e:
        # ✅ التحقق مما إذا كان الرقم غير مسجل في واتساب
        error_message = driver.find_elements(By.XPATH, '//div[contains(text(), "غير مسجل في واتساب")]')
        if error_message:
            print(f"❌ The number {number} is not registered on WhatsApp.")
            failed_numbers.append({
                "ClientName": client_name,
                "PhoneNumber": number,
                "Error": "رقم غير مسجل في واتساب"
            })
        else:
            print(f"{RED}❌ Failed to send the message to {client_name} ({number}) - {str(e)}{RESET}")
            failed_numbers.append({
                "ClientName": client_name,
                "PhoneNumber": number,
                "Error": str(e)  # عرض الخطأ الفعلي هنا مفيد للتشخيص
            })
        continue  # الانتقال إلى الرقم التالي

# ✅ إغلاق المتصفح بعد الانتهاء
driver.quit()

# ✅ حفظ الأرقام الفاشلة في ملف Excel
if failed_numbers:
    failed_df = pd.DataFrame(failed_numbers)
    failed_df.to_excel("failed_numbers.xlsx", index=False)
    print("\n📁 The failed numbers have been saved in the file failed_numbers.xlsx")
else:
    print("\n🎉 All messages were sent successfully, no failed numbers.")


# ✳️ إغلاق البرنامج بعد الضغط على Enter
input("\n🔚 Press Enter to close the program...")
