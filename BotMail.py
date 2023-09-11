#_______________________________________________________________________
#----------------------ส่วนของ GUI 2-------------------------------------
#_______________________________________________________________________
import tkinter as tk
from tkinter import filedialog

# อ่านที่อยู่ไฟล์ล่าสุดจากไฟล์ข้อความ (PathData.txt) ถ้ามี
try:
    with open("PathData.txt", "r", encoding="utf-8") as file:
        Path = file.read()
except FileNotFoundError:
    Path = ""

# อ่านค่า Email ผู้ส่งและ Email ผู้รับล่าสุดจากไฟล์ข้อความ (EmailData.txt) ถ้ามี
try:
    with open("EmailData.txt", "r", encoding="utf-8") as file:
        emailData = file.read().splitlines()
        if len(emailData) == 2:
            emailSentInput, emailGetInput = emailData
        else:
            emailSentInput, emailGetInput = "", ""
except FileNotFoundError:
    emailSentInput, emailGetInput = "", ""

# ถ้ามีข้อมูลใน "PathData.txt" และ "EmailData.txt" ไม่ต้องแสดง GUI
if Path and emailSentInput and emailGetInput:
    # ใช้ค่าที่อ่านได้จากไฟล์
    pass
else:
    root = tk.Tk()
    root.title("เลือกที่อยู่ไฟล์")
    root.geometry("400x300")

    # ฟังก์ชันสำหรับการเลือกไฟล์
    def browse_file():
        file_path = filedialog.askopenfilename()
        if file_path:
            # บันทึกที่อยู่ไฟล์ล่าสุดลงในไฟล์ข้อความ
            with open("PathData.txt", "w", encoding="utf-8") as file:
                file.write(file_path)
                file_path_label.config(text="ที่อยู่ไฟล์: " + file_path)
        else:
            file_path_label.config(text="")

    # สร้างปุ่ม "เลือกไฟล์"
    browse_button = tk.Button(root, text="เลือกไฟล์", command=browse_file)
    browse_button.pack(pady=20)

    # สร้างป้ายกำกับแสดงที่อยู่ไฟล์ที่ถูกเลือก
    file_path_label = tk.Label(root, text="", wraplength=400)
    file_path_label.pack()

    # สร้างป้ายกำกับและกล่องข้อความสำหรับชื่อ Email ผู้ส่งและผู้รับ
    sender_email_label = tk.Label(root, text="Email ผู้ส่ง:")
    sender_email_label.pack()
    sender_email_entry = tk.Entry(root, width=35)
    sender_email_entry.pack()

    receiver_email_label = tk.Label(root, text="Email ผู้รับ:")
    receiver_email_label.pack()
    receiver_email_entry = tk.Entry(root, width=35)
    receiver_email_entry.pack()


    # ฟังก์ชันสำหรับปิดหน้าต่าง GUI และบันทึกค่า Email ผู้ส่งและ Email ผู้รับลงในไฟล์ EmailData.txt
    # ฟังก์ชันสำหรับปิดหน้าต่าง GUI และบันทึกค่า Email ผู้ส่งและ Email ผู้รับลงในไฟล์ EmailData.txt
    def save_and_exit():
        global Path, emailSentInput, emailSenderInput
        Path = file_path_label.cget("text").replace("ที่อยู่ไฟล์: ", "")
        emailSentInput = sender_email_entry.get()
        emailSenderInput = receiver_email_entry.get()
        with open("EmailData.txt", "w", encoding="utf-8") as file:
            file.write(emailSentInput + "\n")
            file.write(emailSenderInput)
        root.destroy()

    # ถ้ามีข้อมูลใน "PathData.txt" และ "EmailData.txt" ให้ใช้ค่าที่อ่านได้จากไฟล์
    if Path and emailSentInput and emailSenderInput:
        sender_email_entry.insert(0, emailSentInput)
        receiver_email_entry.insert(0, emailSenderInput)


    # สร้างปุ่ม "ตกลง" เพื่อบันทึกข้อมูลและปิดหน้าต่าง GUI
    save_button = tk.Button(root, text="ตกลง", command=save_and_exit)
    save_button.pack(pady=20)

    # ในกรณีนี้เราไม่ต้องอ่านค่าจาก EmailData.txt เนื่องจากเราเปิด GUI ใหม่ทุกครั้ง
    emailSentInput, emailSenderInput = "", ""

    # เริ่มการรัน GUI
    root.mainloop()

#_______________________________________________________________________
#----------------------ส่วนของการดึงข้อมูลจากไฟล์ Excel มาคำนวณ-------------
#_______________________________________________________________________
import pandas as pd
from datetime import datetime

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
filtered_data = df[(df['dateSoon'] == 7 )]
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
smtp_username = emailSentInput
smtp_password = 'lhxvhnuwmyhbpkto'

# ข้อมูลผู้ส่งและผู้รับ
from_email = emailSentInput
to_email = emailGetInput

# ตรวจสอบว่ามีข้อมูลหรือไม่
if haveContract == False:
    print("ไม่มีสัญญาที่ใกล้หมดอายุ")
else:
    # สร้างออบเจกต์ MIMEMultipart เพื่อสร้างโครงสร้างของอีเมล
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = emailGetInput  # แก้ไขตรงนี้เป็น emailGetInput
    msg['Subject'] = 'แจ้งเตือนสัญญาหมดอายุในอีก 7 วัน'

    # เพิ่มเนื้อหาของอีเมล
    body = selected_columns.to_html(index=False)  # แปลง DataFrame เป็น HTML table
    msg.attach(MIMEText(body, 'html'))

    # เชื่อมต่อกับเซิร์ฟเวอร์ SMTP ของ Gmail และส่งอีเมล
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(from_email, emailGetInput, msg.as_string())  # แก้ไขตรงนี้เป็น emailGetInput
        server.quit()

        print('Email sent successfully!')
    except Exception as e:
        print('Email could not be sent:', str(e))




