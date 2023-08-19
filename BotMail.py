import pandas as pd
from datetime import datetime

# อ่านข้อมูลจาก Excel ใน sheet1
excel_file = 'ทะเบียนคุมสัญญา.xlsx'  # แทนที่ด้วยชื่อไฟล์ที่คุณใช้
df = pd.read_excel(excel_file, sheet_name='Sheet1')

# แปลงคอลัมน์ที่มีข้อมูลวันที่เป็นรูปแบบ datetime
df['DateUs'] = pd.to_datetime(df['DateUs'])

# วันที่ปัจจุบัน
current_date = datetime.now()

# คำนวณจำนวนวันระหว่างวันที่ในคอลัมน์และวันที่ปัจจุบัน
df['dateSoon'] = (current_date - df['DateUs']).dt.days

# กรองและดึงข้อมูลตามเงื่อนไข
filtered_data = df[(df['dateSoon'] <= 7) & (df['dateSoon'] >= 0)]

# ถ้าคอลัมน์ C ถึง G เป็นคอลัมน์ที่ 2 ถึง 6 (ลำดับหมายเลข 1-5 ใน Python)
selected_columns = filtered_data.iloc[:, 2:7]
