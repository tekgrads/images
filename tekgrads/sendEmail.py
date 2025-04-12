import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd

# Load your contacts from CSV or Excel
# contacts = pd.read_csv("contacts.xlsx")  # Should have 'Email' and 'Name' columns

contacts = pd.read_excel(r"C:\Users\pavan\OneDrive\Desktop\tekgrads\contacts_2.xlsx")


# print(contacts['Email ID'])
# print(contacts['Phone Number'])
# print(contacts['Name'])

# Email setup
sender_email = "tekgrads@gmail.com"
password = "tekgrads042025"

subject = "🎯 One Internship. One Chance. One Way to Get Hired."
body = """Dear [Name],

Let’s be blunt:
Certificates won’t get you a job.
College projects won’t impress recruiters.
Only real-world experience matters.

That’s exactly what InternShift-CTC by TekGrads gives you. And no one else does.

🚀 6 Weeks of Hardcore, Industry-Grade Training
💻 Daily Standups, Live Code Reviews, Project Delivery
🧠 Mentorship from 15+ Years Experienced IT Pros
🎯 Hands-On SDLC + Final Deployment Demo
💼 Job-Ready Profile + Placement Support

No theory. No fluff. Only results.
If you're serious about landing a high-paying tech job—this is your only shot.

📅 Starts April 12
⏳ Limited Seats | 100% Job-Oriented
📞 Call Now: +91 96185 44054
🌐 www.tekgrads.com

Act now—or get left behind.

– Ravindar Kothapally
Founder & CEO, TekGrads

"""

# SMTP setup (for Gmail)
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(sender_email, password)

# Send emails
for index, row in contacts.iterrows():
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = row["Email"]
    msg["Subject"] = subject

    # Attach message body
    msg.attach(MIMEText(body, "plain"))

    # Attach a file
    with open("your_attachment.pdf", "rb") as f:
        part = MIMEApplication(f.read(), Name="your_attachment.pdf")
        part["Content-Disposition"] = 'attachment; filename="your_attachment.pdf"'
        msg.attach(part)

    # Send it
    server.send_message(msg)

server.quit()
print("Emails sent successfully!")
