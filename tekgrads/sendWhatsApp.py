import pandas as pd
import requests
import time

# WhatsApp API setup
ACCESS_TOKEN = 'YOUR_ACCESS_TOKEN'
PHONE_NUMBER_ID = 'YOUR_PHONE_NUMBER_ID'
PDF_MEDIA_ID = 'MEDIA_ID_FOR_PDF'       # Get this from /media upload
JPEG_MEDIA_ID = 'MEDIA_ID_FOR_IMAGE'    # Get this from /media upload

# Load Excel
df = pd.read_excel('contacts2.xlsx')

# Base URLs
API_URL = f'https://graph.facebook.com/v18.0/{PHONE_NUMBER_ID}/messages'
HEADERS = {
    'Authorization': f'Bearer {ACCESS_TOKEN}',
    'Content-Type': 'application/json'
}

# Personalized message template
def create_message(name):
    return f"""Dear {name},   

🎯 Let’s be real—only this program can get you a job.

❌ No certificate can do it.
❌ No college internship can match it.
✅ InternShift-CTC by TekGrads = Real Projects + Job Offers

✅ 8 Weeks of Hardcore, Industry-Grade Training
✅ Daily Standups, Live Code Reviews, Project Delivery
✅ Mentorship from 15+ Years Experienced IT Pros
✅ Unlimited AI-Powered Mock Interviews
✅ Resume & LinkedIn Profile Optimization
✅ Industry-Recognized Certification
✅ Job-Ready Profile + Placement Support

📅 Starting April 19
🔗 Register Now: https://tekgrads.com/contact-us/  
📞 +91-9652544054 | 🌐 www.tekgrads.com
one program. One shot at your IT career. Don’t miss it."""

# Send message
def send_text_message(number, message):
    payload = {
        "messaging_product": "whatsapp",
        "to": str(number),
        "type": "text",
        "text": {
            "body": message
        }
    }
    response = requests.post(API_URL, headers=HEADERS, json=payload)
    print(f"Text sent to {number}: {response.status_code}")

# Send attachment
def send_attachment(number, media_id, media_type, caption=None):
    payload = {
        "messaging_product": "whatsapp",
        "to": str(number),
        "type": media_type,
        media_type: {
            "id": media_id,
        }
    }
    if caption and media_type == "document":
        payload["document"]["caption"] = caption
    elif caption and media_type == "image":
        payload["image"]["caption"] = caption

    response = requests.post(API_URL, headers=HEADERS, json=payload)
    print(f"{media_type.capitalize()} sent to {number}: {response.status_code}")

# Loop through all contacts
for index, row in df.iterrows():
    name = row['Name']
    phone = row['Phone Number']
    msg = create_message(name)

    send_text_message(phone, msg)
    time.sleep(1)

    send_attachment(phone, PDF_MEDIA_ID, "document", caption="InternShift Program Details")
    time.sleep(1)

    send_attachment(phone, JPEG_MEDIA_ID, "image", caption="InternShift Handout")
    time.sleep(1)
