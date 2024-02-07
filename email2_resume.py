import smtplib
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from string import Template
from pathlib import Path
import pandas as pd

# Load the data from the Excel file
df = pd.read_excel("./consultant.xlsx")

# Read the HTML template
html = Template(Path('./index.html').read_text())

# Setting up the server
with smtplib.SMTP(host='smtp.gmail.com', port=587) as smtp:
    smtp.ehlo()
    smtp.starttls()

    # Login to your Gmail account
    smtp.login("smskanchu@gmail.com", "cooj xpvh qdno woat")

    # Iterate over rows in the DataFrame
    for index, row in df.iterrows():
        # Create a multipart message
        email = MIMEMultipart()
        email['from'] = 'Yogesh Pant'

        # Set the recipient's email address
        email['to'] = row['Email']

        # Set the email subject
        email['subject'] = 'Looking for the opportunity'

        # Attach the HTML body
        email.attach(MIMEText(html.substitute(name=row['Name']), 'html'))

        # Attach the resume
        resume_path = './Supriya_Maharjan_Sapkota.docx'  # Update this to the path of your resume file
        with open(resume_path, 'rb') as f:
            resume = MIMEApplication(f.read(), _subtype='docx')
            resume.add_header('Content-Disposition', 'attachment', filename=Path(resume_path).name)
            email.attach(resume)

        # Send the email
        smtp.send_message(email)
        print(f"Email sent to {row['Name']} ({row['Email']})")

print("All emails sent successfully.")
