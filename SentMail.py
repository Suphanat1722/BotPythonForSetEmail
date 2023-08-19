import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from BotMail import selected_columns

# ตรวจสอบว่าไฟล์บันทึกอีเมลอยู่หรือไม่
sent_emails_file = 'sent_emails.txt'
if os.path.exists(sent_emails_file):
    with open(sent_emails_file, 'r') as f:
        sent_emails = f.read().splitlines()
else:
    sent_emails = []

# ข้อมูลเข้าสู่ระบบ SMTP ของ Gmail
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = 'angnew1155@gmail.com'
smtp_password = 'lhxvhnuwmyhbpkto'

# ข้อมูลผู้ส่งและผู้รับ
from_email = 'angnew1155@gmail.com'
to_email = 'nduba1722@gmail.com'

# ตรวจสอบว่าเคยส่งไปแล้วหรือไม่
if to_email in sent_emails:
    print('Email has already been sent to this recipient.')
else:
    # สร้างออบเจกต์ MIMEMultipart เพื่อสร้างโครงสร้างของอีเมล
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = 'แจ้งเตือนวันหมดอายุ'

    print(selected_columns)
    # เพิ่มเนื้อหาของอีเมล
    body = str(selected_columns)
    msg.attach(MIMEText(body, 'plain'))

    # เชื่อมต่อกับเซิร์ฟเวอร์ SMTP ของ Gmail และส่งอีเมล
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()

        # เพิ่มอีเมลล์ที่ส่งไปลงในรายการบันทึก
        sent_emails.append(to_email)
        with open(sent_emails_file, 'a') as f:
            f.write(to_email + '\n')

        print('Email sent successfully!')
    except Exception as e:
        print('Email could not be sent:', str(e))
