import pandas as pd
import requests
import time

# WATI API setup
API_KEY = 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJqdGkiOiJkOWU2MGNjZi1lOGQ4LTQ4ZmYtYmM2ZS01YTlmNzIwZTk4YmYiLCJ1bmlxdWVfbmFtZSI6InRla2dyYWRzQGdtYWlsLmNvbSIsIm5hbWVpZCI6InRla2dyYWRzQGdtYWlsLmNvbSIsImVtYWlsIjoidGVrZ3JhZHNAZ21haWwuY29tIiwiYXV0aF90aW1lIjoiMDQvMTMvMjAyNSAwNDo0NDo0OCIsInRlbmFudF9pZCI6IjQzMDU3OSIsImRiX25hbWUiOiJtdC1wcm9kLVRlbmFudHMiLCJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL3dzLzIwMDgvMDYvaWRlbnRpdHkvY2xhaW1zL3JvbGUiOiJBRE1JTklTVFJBVE9SIiwiZXhwIjoyNTM0MDIzMDA4MDAsImlzcyI6IkNsYXJlX0FJIiwiYXVkIjoiQ2xhcmVfQUkifQ.NSbnEW_6UCxXVhfZyeGl7xTi5bGKdcNVE5KGbTXUye0'
BASE_URL = 'https://live-mt-server.wati.io/430579/sendSessionMessage'
MEDIA_ENDPOINT = "https://app.wati.io/api/v1/mediaMessage"

HEADERS = {
    'Authorization': f'{API_KEY}',
    'Content-Type': 'application/json'
}

# Media URLs hosted on GitHub
PDF_URL_1 = "https://github.com/tekgrads/images/blob/3c01aa71fb79ada03c91ffc5bc27296f4c5d1db1/tekgrads/Internshift-CTC-Presentation.pdf"
PDF_URL_2 = "https://github.com/tekgrads/images/blob/3c01aa71fb79ada03c91ffc5bc27296f4c5d1db1/tekgrads/TekGrads-Brochure.pdf"
IMAGE_URL = "https://raw.githubusercontent.com/tekgrads/images/refs/heads/main/tekgrads/tekgrads-internship-handout.jpeg"

# Load contacts from Excel
df = pd.read_excel(r'C:\Users\rajes\OneDrive\Documents\GitHub\images\tekgrads\contacts_2.xlsx')

# Personalized message template
def create_message(name):
    return f"""Dear {name},   

🎯 Let’s be real—only this program can get you a job.

❌ No certificate can do it.
❌ No college internship can match it.
✅ InternShift-CTC by TekGrads = Real Projects + Job Offers

📅 Starting April 19
🔗 Register Now: https://tekgrads.com/contact-us/  
📞 +91-9652544054 | 🌐 www.tekgrads.com
One program. One shot at your IT career. Don’t miss it."""

# Send plain text message
def send_text(number, message):
    payload = {
        "phone": str(number),
        "message": message
    }
    response = requests.post(BASE_URL, headers=HEADERS, json=payload)
    print(f"Text sent to {number}: {response.status_code}")

# Send media message (PDF/Image via URL)
def send_media(number, media_url, media_type, caption):
    payload = {
        "phone": str(number),
        "url": media_url,
        "media_type": media_type,
        "caption": caption
    }
    response = requests.post(MEDIA_ENDPOINT, headers=HEADERS, json=payload)
    print(f"{media_type.upper()} sent to {number}: {response.status_code}")

# Iterate through contacts and send messages
for index, row in df.iterrows():
    name = row['Name']
    phone = row['Phone Number']
    message = create_message(name)

    send_text(phone, message)
    time.sleep(1)

    send_media(phone, PDF_URL_1, "document", "InternShift Program Details")
    time.sleep(1)

    send_media(phone, PDF_URL_2, "document", "InternShift Schedule")
    time.sleep(1)

    send_media(phone, IMAGE_URL, "image", "InternShift Handout")
    time.sleep(1)
