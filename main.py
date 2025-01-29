import smtplib
import pandas as pd
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import schedule
import time

# Email Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_ADDRESS = "Vikramadityanaruka@gmail.com"  # Your email
EMAIL_PASSWORD = "rfwa zbsj mfnu cyjv"  # Replace with your Gmail app password

# Function to send email
def send_email(to_email, candidate_name):
    try:
        subject = "Scheduled Email Notification"
        body = f"Hello {candidate_name},\n\nThis is your scheduled email notification.\n\nBest regards,\nYour Team"

        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)

        print(f"✅ Email sent to {to_email} for {candidate_name}.")
    except Exception as e:
        print(f"❌ Error sending email to {to_email}: {e}")

# Function to read Excel and send emails
def check_and_send_emails():
    try:
        df = pd.read_excel("email_schedule.xlsx")  # Make sure this file exists

        current_date = datetime.now().strftime("%Y-%m-%d")

        for _, row in df.iterrows():
            candidate_name = row['Name']
            scheduled_date = row['Date'].strftime("%Y-%m-%d")
            email = row['Email']

            if scheduled_date == current_date:
                send_email(email, candidate_name)
    except Exception as e:
        print(f"❌ Error processing the Excel file: {e}")

# Schedule to run daily at 11:59 AM
schedule.every().day.at("15:00").do(check_and_send_emails)

print("📧 Email scheduler is running...")
while True:
    schedule.run_pending()
    time.sleep(60)
