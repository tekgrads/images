import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage

import os

# Load contacts
contacts = pd.read_excel(r"C:\Users\pavan\OneDrive\Desktop\tekgrads\contacts_2.xlsx")  # Should contain 'Name' and 'Email'

# Email account setup
sender_email = "pavanimadireddy3@gmail.com"
password = "cnzv hprl vmjp tefg"

# Email content
subject = "{Name} One Internship. One Chance. One Way to Get Hired."

# Corrected Email Template
body_template = """
<html>
<body style="font-family: Arial, sans-serif; color: #333; margin: 0; padding: 0; background-color: #f0f8f4;">

<table width="100%" bgcolor="#008080" cellpadding="0" cellspacing="0">
    <tr>
        <td align="center" style="padding: 20px;">
            <img src="tek_grads_logo.png" alt="TekGrads Logo" width="150">            
        </td>
    </tr>
</table>

<table width="100%" bgcolor="#e0f2f0" cellpadding="0" cellspacing="0">
    <tr>
        <td align="center" style="padding: 30px;">
            <img src="ctc.png" alt="Internship Opportunity" style="max-width: 100%; height: auto; border-radius: 8px;">
        </td>
    </tr>
</table>

<table width="100%" cellpadding="0" cellspacing="0" style="padding: 20px;">
    <tr>
        <td style="font-size: 18px; line-height: 1.6; padding: 20px; background-color: #ffffff; border-radius: 8px; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);">
            <p style="font-size: 22px; color: #008080; font-weight: bold;">Dear {Name},</p>

            <p style="font-size: 16px; color: #555;">Let’s be blunt:</p>
            <ul style="font-size: 16px; color: #555; line-height: 1.6; padding-left: 20px;">
                <li>Certificates won’t get you a job.</li>
                <li>College projects won’t impress recruiters.</li>
                <li><span style="color: #00a699; font-weight: bold;">Only real-world experience matters.</span></li>
            </ul>

            <p style="font-size: 16px; color: #555;">That’s exactly what <span style="color: #00a699; font-weight: bold;">InternShift-CTC by TekGrads</span> gives you. And no one else does.</p>

            <p style="font-size: 18px; color: #008080; font-weight: bold;">What We Offer:</p>
            <ul style="font-size: 16px; color: #555; line-height: 1.6; padding-left: 20px;">
                <li><span style="color: #00a699; font-weight: bold;">6 Weeks of Hardcore, Industry-Grade Training</span></li>
                <li><span style="color: #00a699; font-weight: bold;">Daily Standups, Live Code Reviews, Project Delivery</span></li>
                <li><span style="color: #00a699; font-weight: bold;">Mentorship from 15+ Years Experienced IT Pros</span></li>
                <li><span style="color: #00a699; font-weight: bold;">Hands-On SDLC + Final Deployment Demo</span></li>
                <li><span style="color: #00a699; font-weight: bold;">Job-Ready Profile + Placement Support</span></li>
            </ul>

            <p style="font-size: 16px; color: #555; font-weight: bold;">No theory. No fluff. Only results.</p>
            <p style="font-size: 16px; color: #555;">If you're serious about landing a <span style="color: #00a699; font-weight: bold;">high-paying tech job</span>—this is your only shot.</p>

            <p style="font-size: 18px; color: #008080; font-weight: bold;">Program Details:</p>
            <ul style="font-size: 16px; color: #555; line-height: 1.6; padding-left: 20px;">
                <li>📅 <span style="color: #00a699; font-weight: bold;">Starts April 19</span></li>
                <li>⏳ <span style="color: #00a699; font-weight: bold;">Limited Seats | 100% Job-Oriented</span></li>
                <li>📞 <span style="color: #00a699; font-weight: bold;">Call Now: +91-9652544054</span></li>
                <li>🌐 <a href="http://www.tekgrads.com" style="color: #00a699;">www.tekgrads.com</a></li>
            </ul>

            <p style="font-size: 16px; color: #555; font-weight: bold;">Act now—or get left behind.</p>
        </td>
    </tr>
</table>

<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#008080">
    <tr>
        <td align="center" style="padding: 10px; color: #b2dfdb; font-size: 14px;">
            – Ravindar Kothapally<br>Founder & CEO, TekGrads
        </td>
    </tr>
</table>

</body>
</html>

"""



# Attachments - add full paths if not in same folder
attachments = [r"C:\Users\pavan\OneDrive\Desktop\tekgrads\Internshift-CTC-Presentation.pdf", r"C:\Users\pavan\OneDrive\Desktop\tekgrads\TekGrads-Brochure.pdf", r"C:\Users\pavan\OneDrive\Desktop\tekgrads\tekgrads-internship-handout.jpeg"]  # Add as many as needed

# Setup SMTP
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(sender_email, password)

# Loop through contacts
for index, row in contacts.iterrows():
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = row["Email ID"]
    msg["Subject"] = subject.format(Name=row["Name"])

    # Personalize body
    body = body_template.format(Name=row["Name"])
    msg.attach(MIMEText(body, "html"))

    # Attach multiple files
    for file_path in attachments:
        with open(file_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(file_path)}"'
            msg.attach(part)

    # Send email
    server.send_message(msg)

server.quit()
print("All emails sent successfully!")
