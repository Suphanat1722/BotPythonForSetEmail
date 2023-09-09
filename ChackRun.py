#_______________________________________________________________________
#----------------------ส่วนของ ระบบเช็คให้ run วันละครั้ง---------------------
#_______________________________________________________________________
import os
import sys
import datetime

program_path = r"C:\Users\NewAng\Desktop\TestEXE\dist\BotMail\BotMail.exe"
marker_file = r"C:\Users\NewAng\Desktop\TestEXE\marker.txt"


# ตรวจสอบว่า marker.txt มีอยู่หรือไม่
if not os.path.exists(marker_file):
    # ถ้าไม่มี marker.txt ให้รันโปรแกรม
    os.system(program_path)
    
    # สร้าง marker.txt เพื่อบ่งชี้ว่าโปรแกรมได้รับการรันแล้ว
    with open(marker_file, "w") as file:
        file.write(str(datetime.date.today()))

# ถ้ามี marker.txt แล้ว ตรวจสอบว่าวันที่บน marker.txt ตรงกับวันนี้หรือไม่
else:
    with open(marker_file, "r") as file:
        last_run_date = file.read()
    
    today = str(datetime.date.today())
    
    if last_run_date != today:
        # ถ้าวันที่ไม่ตรงกัน ให้รันโปรแกรมและอัปเดต marker.txt
        os.system(program_path)
        with open(marker_file, "w") as file:
            file.write(today)


