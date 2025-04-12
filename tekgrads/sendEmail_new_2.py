import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os

# Load contacts
contacts = pd.read_excel(r"C:\Users\pavan\OneDrive\Desktop\tekgrads\contacts_2.xlsx")

# Email setup
sender_email = "pavanimadireddy3@gmail.com"
password = "cnzv hprl vmjp tefg"

# Subject
subject = "{Name} One Internship. One Chance. One Way to Get Hired."

# HTML body template
body_template = """
<html>
<head>
  <style>
    body, html {{
      margin: 0;
      padding: 0;
      width: 100%;
    }}
    img {{
      display: block;
      max-width: 100%;
      height: auto;
    }}
    .container {{
      padding: 0;
      margin: 0;
      width: 100%;
      background-color: #f0f8f4;
      font-family: Arial, sans-serif;
      color: #333;
    }}
    .content {{
      padding: 20px;
      background-color: #ffffff;
      border-radius: 8px;
      box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
      font-size: 16px;
    }}
    .footer {{
      text-align: center;
      background-color: #008080;
      color: #b2dfdb;
      font-size: 14px;
      padding: 15px 10px;
    }}
    @media only screen and (max-width: 600px) {{
      .content {{
        padding: 15px;
        font-size: 14px;
      }}
      .content p {{
        font-size: 14px !important;
      }}
      .footer {{
        font-size: 12px;
      }}
    }}
  </style>
</head>
<body>
  <div class="container">
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center">
          <img src="https://raw.githubusercontent.com/tekgrads/images/main/tek_grads_logo.png" alt="TekGrads Banner" width="100%" style="max-width: 100%; height: auto; display: block; border: 0; margin: 0; padding: 0;">
        </td>
      </tr>
    </table>

    <div class="content">
      <p style="font-size: 22px; color: #008080; font-weight: bold;">Dear {Name},</p>

      <p style="font-size: 16px;">Let’s be blunt:</p>
      <ul style="line-height: 1.6;">
        <li>Certificates won’t get you a job.</li>
        <li>College projects won’t impress recruiters.</li>
        <li><strong style="color: #00a699;">Only real-world experience matters.</strong></li>
      </ul>

      <p>That’s exactly what <strong style="color: #00a699;">InternShift-CTC by TekGrads</strong> gives you. And no one else does.</p>

      <p style="font-size: 18px; color: #008080; font-weight: bold;">What We Offer:</p>
      <ul style="line-height: 1.6;">
        <li><strong style="color: #00a699;">8 Weeks of Hardcore, Industry-Grade Training</strong></li>
        <li><strong style="color: #00a699;">Daily Standups, Live Code Reviews, Project Delivery</strong></li>
        <li><strong style="color: #00a699;">Mentorship from 15+ Years Experienced IT Pros</strong></li>
        <li><strong style="color: #00a699;"> Unlimited AI-Powered Mock Interviews</strong></li>
         <li><strong style="color: #00a699;"> Resume & LinkedIn Profile Optimization</strong></li>
          <li><strong style="color: #00a699;">Industry-Recognized Certification</strong></li>
        <li><strong style="color: #00a699;">Job-Ready Profile + Placement Support</strong></li>
      </ul>

      <p><strong>No theory. No fluff. Only results.</strong></p>
      <p>If you're serious about landing a <strong style="color: #00a699;">high-paying tech job</strong>—this is your only shot.</p>

      <p style="font-size: 18px; color: #008080; font-weight: bold;">Program Details:</p>
      <table width="100%" cellspacing="0" cellpadding="5" style="font-size: 16px; color: #555;">
        <tr>
          <td style="vertical-align: top; width: 24px;">📅</td>
          <td><strong style="color: #00a699;">Starts April 19</strong></td>
        </tr>
        <tr>
          <td style="vertical-align: top;">⏳</td>
          <td><strong style="color: #00a699;">Limited Seats | 100% Job-Oriented</strong></td>
        </tr>
        <tr>
          <td style="vertical-align: top;">📞</td>
          <td>
            <a href="tel:+919652544054" style="color: #00a699; text-decoration: none;">
              <strong><span style="white-space: nowrap;">Call Now: +91-9652544054</span></strong>
            </a>
          </td>
        </tr>
        <tr>
          <td style="vertical-align: top;">🌐</td>
          <td><a href="http://www.tekgrads.com" style="color: #00a699;"><strong>www.tekgrads.com</strong></a></td>
        </tr>
      </table>

      <p style="font-weight: bold;">Act now—or get left behind.</p>
    </div>

    <p style="text-align: center; margin-top: 30px;">
  <a href="https://tekgrads.com/contact-us/" target="_blank" 
     style="background-color: #008080; color: #ffffff; padding: 12px 25px; 
            text-decoration: none; border-radius: 5px; font-size: 16px; 
            display: inline-block;">
     Register Now
  </a>
</p>


    <div class="footer">
      – Ravindar Kothapally<br>Founder & CEO, TekGrads
    </div>
  </div>
</body>
</html>
"""

# Attachments
attachments = [
    r"C:\Users\pavan\OneDrive\Desktop\tekgrads\Internshift-CTC-Presentation.pdf",
    r"C:\Users\pavan\OneDrive\Desktop\tekgrads\TekGrads-Brochure.pdf",
    r"C:\Users\pavan\OneDrive\Desktop\tekgrads\tekgrads-internship-handout.jpeg"
]

# Setup SMTP
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(sender_email, password)

# Loop through contacts
for index, row in contacts.iterrows():
    msg = MIMEMultipart('related')
    msg["From"] = sender_email
    msg["To"] = row["Email ID"]
    msg["Subject"] = subject.format(Name=row["Name"])

    msg_alternative = MIMEMultipart('alternative')
    msg.attach(msg_alternative)

    # Personalize HTML
    html_content = body_template.format(Name=row["Name"])
    msg_alternative.attach(MIMEText(html_content, "html"))

    # Attach files
    for file_path in attachments:
        with open(file_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(file_path)}"'
            msg.attach(part)

    # Send email
    server.send_message(msg)

server.quit()
print("✅ All emails sent successfully!")
