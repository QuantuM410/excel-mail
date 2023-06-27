import pandas as pd
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

load_dotenv()


def send_email(smtp_server, sender_email, sender_password, receiver_email, subject, message):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    with smtplib.SMTP(smtp_server, 587) as server:
        server.ehlo()
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)


data = pd.read_excel('sample_excel_mail.xlsx')

for _, row in data.iterrows():
    for column_name in data.columns:
        if 'email' in column_name.lower():
            email = str(row[column_name])
            if email and '@' in email:
                subject = "Your Subject"
                message = f"Dear recipient,\n\nHere is your data:\n\n{row.dropna().to_string(index=False)}\n\nRegards,\nYour Name"
                smtp_server = os.getenv('smtp_server')
                sender_email = os.getenv('sender_email')
                sender_password = os.getenv('sender_password')
                send_email(smtp_server, sender_email,
                           sender_password, email, subject, message)
