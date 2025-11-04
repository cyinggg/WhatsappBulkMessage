import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from urllib.parse import quote
from datetime import datetime
import os

# -----------------------------
# 1️⃣ Read contacts from Excel
# -----------------------------
contacts_file = "contacts.xlsx"
df = pd.read_excel(contacts_file)

# -----------------------------
# 2️⃣ Setup Selenium Chrome WebDriver
# -----------------------------
service = Service("chromedriver.exe")  # ensure chromedriver.exe is in project folder
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)  # keeps browser open after script ends
driver = webdriver.Chrome(service=service, options=options)

# Open WhatsApp Web
driver.get("https://web.whatsapp.com/")
print("🔑 Please scan the QR code to log in to WhatsApp Web...")
time.sleep(25)  # wait for manual login (adjust if needed)

# -----------------------------
# 3️⃣ Prepare log DataFrame
# -----------------------------
log_file = "sent_log.xlsx"
if os.path.exists(log_file):
    log_df = pd.read_excel(log_file)
else:
    log_df = pd.DataFrame(columns=["Timestamp", "Name", "Phone", "Status"])

# -----------------------------
# 4️⃣ Send personalized message
# -----------------------------
for index, row in df.iterrows():
    name = row['Name']
    phone = str(row['Phone']).strip()

    # Personalize message
    message = (
        f"Hello {name}, thank you for being Student Coach serving ProjectHub!\n\n"
        "As we now review and check to ensure all current SCs have access to ProjectHub, "
        "may I trouble you to help me submit your SIT Student Card CAN ID via the following link "
        "https://forms.office.com/r/FQndAcGAQm *latest by today, 5th November, 1800hrs*\n\n"
        "Please *reply to this message after* you have *submitted your SIT Student Card CAN ID*, thank you!"
    )

    try:
        # Encode message for URL
        encoded_msg = quote(message)
        url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded_msg}"

        driver.get(url)
        time.sleep(10)  # wait for chat box to load

        # Focus on input box and press ENTER to send
        input_box = driver.switch_to.active_element
        input_box.send_keys(Keys.ENTER)

        print(f"✅ Message sent to {name} ({phone})")

        # Log success
        log_df = pd.concat([log_df, pd.DataFrame([{
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Name": name,
            "Phone": phone,
            "Status": "Sent"
        }])], ignore_index=True)

        time.sleep(8)  # wait before sending next message

    except Exception as e:
        print(f"❌ Failed to send to {name} ({phone}) -> {e}")

        # Log failure
        log_df = pd.concat([log_df, pd.DataFrame([{
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Name": name,
            "Phone": phone,
            "Status": f"Failed: {e}"
        }])], ignore_index=True)
        continue

# -----------------------------
# 5️⃣ Save log to Excel
# -----------------------------
log_df.to_excel(log_file, index=False)
print(f"📄 Log saved to {log_file}")

print("🎉 All messages processed!")
