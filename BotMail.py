#_______________________________________________________________________
#----------------------ส่วนของการดึงข้อมูลจากไฟล์ Excel มาคำนวณ-------------
#_______________________________________________________________________
import pandas as pd
from datetime import datetime
from BotMailGUI import Path

# อ่านข้อมูลจาก Excel 
excel_file = Path  # แทนที่ด้วยชื่อไฟล์ที่คุณใช้
df = pd.read_excel(excel_file, sheet_name='SheetData')

# แปลงคอลัมน์ที่มีข้อมูลวันที่เป็นรูปแบบ datetime
df['DateUs'] = pd.to_datetime(df['DateUs'])

# วันที่ปัจจุบัน
current_date = datetime.now()

# คำนวณจำนวนวันระหว่างวันที่ในคอลัมน์และวันที่ปัจจุบัน
df['dateSoon'] = (df['DateUs'] - current_date).dt.days

# กรองและดึงข้อมูลตามเงื่อนไข
filtered_data = df[(df['dateSoon'] == 9 )]
selected_columns = filtered_data

haveContract = False
# ตรวจสอบว่ามีข้อมูลหรือไม่
if filtered_data.empty:
    print("ไม่มีสัญญาวันหมดอายุในช่วง 0 ถึง 7 วัน")
    haveContract = False
else:
    # ถ้ามีข้อมูลตรงกับเงื่อนไข
    selected_columns = filtered_data.iloc[:, 2:7]
    haveContract = True
    print(selected_columns)

#_______________________________________________________________________
#----------------------ส่วนของการส่ง Email--------------------------------
#_______________________________________________________________________
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ข้อมูลเข้าสู่ระบบ SMTP ของ Gmail
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = 'angnew1155@gmail.com'
smtp_password = 'lhxvhnuwmyhbpkto'

# ข้อมูลผู้ส่งและผู้รับ
from_email = 'angnew1155@gmail.com'
to_email = 'nduba1722@gmail.com'

# ตรวจสอบว่ามีข้อมูลหรือไม่
if haveContract == False:
    print("ไม่มีสัญญาที่ใกล้หมดอายุ")
else:
    # สร้างออบเจกต์ MIMEMultipart เพื่อสร้างโครงสร้างของอีเมล
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = 'แจ้งเตือนสัญญาหมดอายุในอีก 7 วัน'

    # เพิ่มเนื้อหาของอีเมล
    body = selected_columns.to_html(index=False)  # แปลง DataFrame เป็น HTML table
    msg.attach(MIMEText(body, 'html'))

    # เชื่อมต่อกับเซิร์ฟเวอร์ SMTP ของ Gmail และส่งอีเมล
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()

        print('Email sent successfully!')
    except Exception as e:
        print('Email could not be sent:', str(e))

