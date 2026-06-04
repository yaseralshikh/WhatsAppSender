import time
import random
import urllib.parse
from datetime import datetime
import sys

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys

# ✅ ألوان الكونسول
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
CYAN = "\033[96m"
RESET = "\033[0m"

# ✅ تحميل ملف الأرقام
try:
    df = pd.read_excel("contacts.xlsx")
except Exception as e:
    print(f"{RED}❌ تعذّر قراءة ملف contacts.xlsx - تأكد من وجوده بجانب البرنامج. التفاصيل: {e}{RESET}")
    input("اضغط Enter للخروج...")
    sys.exit(1)

required_cols = {"ClientName", "PhoneNumber", "Message"}
if not required_cols.issubset(df.columns):
    print(f"{RED}❌ ملف contacts.xlsx يجب أن يحتوي الأعمدة: ClientName, PhoneNumber, Message{RESET}")
    input("اضغط Enter للخروج...")
    sys.exit(1)

# ✅ إعداد متصفح Chrome مع بروفايل ثابت (خارج Temp لضمان الاحتفاظ بالجلسة)
options = webdriver.ChromeOptions()
options.add_argument(r"--user-data-dir=C:\Users\Yaser Alshikh\AppData\Local\chrome_profile_whatsapp")
options.add_argument("--profile-directory=Default")
options.add_argument("--disable-extensions")
options.add_argument("--disable-infobars")
options.add_argument("--disable-notifications")
options.add_argument("--log-level=3")

# ✅ استخدام chromedriver المحلي بدلاً من التحميل كل مرة
service = Service(r"chromedriver-win64\chromedriver-win64\chromedriver.exe")
driver = webdriver.Chrome(service=service, options=options)

# ✅ فتح واتساب ويب
driver.get("https://web.whatsapp.com/")

print(f"{CYAN}----------------------------------------------{RESET}")
print(f"{YELLOW}🚀 Program created by: {GREEN}Yaser Mohammed Alshikh{RESET}")
print(f"{YELLOW}✉ Email: {CYAN}yaseralshikh@gmail.com{RESET}")
print(f"{YELLOW}🔒 Protected and owned. {RESET}📌 Contact for improvements or licensing.")
print(f"{YELLOW}© All rights reserved. {RESET}💡 Idea and implementation by Yaser.")
print(f"{CYAN}----------------------------------------------\n{RESET}")

# ✅ التحقق من تسجيل الدخول — إذا لم يكن مسجلاً نطلب مسح QR
print(f"{CYAN}⏳ جاري التحقق من الجلسة المحفوظة...{RESET}")
try:
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                '//div[@role="grid"] | //div[@data-testid="chat-list"]'
            )
        )
    )
    print(f"{GREEN}✅ تم استعادة الجلسة — لا حاجة لمسح QR.{RESET}")
except Exception:
    print(f"{YELLOW}🔵 لم يتم العثور على جلسة محفوظة. افتح الجوال ومسح QR في واتساب ويب.{RESET}")
    input("اضغط Enter بعد ظهور محادثاتك...")
    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '//div[@role="grid"] | //div[@data-testid="chat-list"]'
                )
            )
        )
        print(f"{GREEN}✅ تم تسجيل الدخول بنجاح.{RESET}")
    except Exception:
        print(f"{RED}❌ لم يتم اكتشاف واجهة المحادثات. أعد تشغيل البرنامج.{RESET}")
        driver.quit()
        input("اضغط Enter للخروج...")
        sys.exit(1)

failed_numbers = []
log_records = []

# ✅ دالة مساعدة لإرسال رسالة لرقم واحد
def send_message(client_name, number, message_text):
    full_message = f"مرحبًا {client_name}، {message_text}"
    encoded_message = urllib.parse.quote(full_message)
    url = f"https://web.whatsapp.com/send?phone={number}&text={encoded_message}"

    driver.get(url)

    try:
        # ننتظر تحميل واجهة المحادثة أو رسالة الخطأ إن الرقم غير مسجل
        # انتظار صندوق الكتابة في أسفل الشاشة
        msg_box = WebDriverWait(driver, 25).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    # صندوق إدخال الرسالة الرئيسي (data-testid/textbox غالباً الأحدث)
                    '//div[@contenteditable="true" and @data-testid="conversation-compose-box-input"]'
                    ' | //div[@contenteditable="true" and @data-testid="msg-input"]'
                    ' | //div[@contenteditable="true" and starts-with(@data-tab,"10")]'
                )
            )
        )

        time.sleep(1.5)  # لتأكد من تحميل النص الجاهز

        # نرسل Enter لإرسال الرسالة
        msg_box.send_keys(Keys.ENTER)

        print(f"{GREEN}✅ تم إرسال الرسالة إلى {client_name} ({number}){RESET}")
        log_records.append({
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "ClientName": client_name,
            "PhoneNumber": number,
            "Status": "Success",
            "Error": ""
        })
        # تأخير عشوائي لتقليل الشبه الآلي
        time.sleep(random.randint(8, 15))

    except Exception as e:
        # التحقق إذا واتساب أظهر أن الرقم غير مسجل
        not_registered = driver.find_elements(
            By.XPATH,
            '//*[contains(text(), "غير مسجل في واتساب") or contains(text(), "not on WhatsApp")]'
        )

        if not_registered:
            error_text = "الرقم غير مسجل في واتساب"
            print(f"{RED}❌ {error_text}: {number}{RESET}")
        else:
            error_text = f"{type(e).__name__}: {str(e)}"
            print(f"{RED}❌ فشل الإرسال إلى {client_name} ({number}) - {error_text}{RESET}")

        failed_numbers.append({
            "ClientName": client_name,
            "PhoneNumber": number,
            "Error": error_text
        })
        log_records.append({
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "ClientName": client_name,
            "PhoneNumber": number,
            "Status": "Failed",
            "Error": error_text
        })

# ✅ الحلقة الأساسية لإرسال الرسائل
for _, row in df.iterrows():
    client_name = str(row["ClientName"]).strip()
    raw_number = str(row["PhoneNumber"]).strip()
    message = str(row["Message"])

    # معالجة الأرقام التي تخرج 9665xxxx.0
    if "." in raw_number:
        raw_number = raw_number.split(".")[0].strip()

    number = raw_number

    if not number.isdigit() or not number.startswith("966"):
        msg = "تنسيق رقم غير صحيح (يجب أن يبدأ بـ 966 وبدون رموز أو مسافات)"
        print(f"{YELLOW}⚠ {msg}: {number}{RESET}")
        failed_numbers.append({
            "ClientName": client_name,
            "PhoneNumber": number,
            "Error": msg
        })
        log_records.append({
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "ClientName": client_name,
            "PhoneNumber": number,
            "Status": "Failed",
            "Error": msg
        })
        continue

    send_message(client_name, number, message)

# ✅ إغلاق المتصفح
driver.quit()

# ✅ حفظ failed_numbers.xlsx
if failed_numbers:
    pd.DataFrame(failed_numbers).to_excel("failed_numbers.xlsx", index=False)
    print(f"\n📁 تم حفظ الأرقام التي فشل إرسالها في failed_numbers.xlsx")
else:
    print(f"\n🎉 تم إرسال جميع الرسائل بنجاح دون أي فشل!")

# ✅ حفظ log.xlsx
if log_records:
    pd.DataFrame(log_records).to_excel("log.xlsx", index=False)
    print(f"📝 تم حفظ سجل كامل بجميع المحاولات في log.xlsx")

input("\n🔚 اضغط Enter لإغلاق البرنامج...")
