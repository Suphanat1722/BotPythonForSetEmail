#_______________________________________________________________________
#----------------------ส่วนของ ระบบเช็คให้ run วันละครั้ง---------------------
#_______________________________________________________________________
import tkinter as tk
from tkinter import filedialog
import os
import datetime

def browse_button():
    global program_path, marker_file
    program_path = filedialog.askopenfilename()
    marker_file = os.path.join(os.path.dirname(program_path), "marker.txt")

def install_program():
    global program_path, marker_file

    if not program_path:
        tk.messagebox.showerror("ข้อผิดพลาด", "กรุณาเลือกโปรแกรม")
        return

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

    # ปิดหน้าต่าง GUI
    root.destroy()

# สร้างหน้าต่าง GUI
root = tk.Tk()
root.title("โปรแกรมติดตั้ง")

# กำหนดขนาดเริ่มต้นของหน้าต่างเป็น 500x300 pixels
root.geometry("500x300")

# สร้างปุ่มสำหรับเลือกโปรแกรม
browse_program_button = tk.Button(root, text="เลือกโปรแกรม", command=browse_button)
browse_program_button.pack()

# สร้างปุ่มสำหรับติดตั้ง
install_button = tk.Button(root, text="ติดตั้งโปรแกรม", command=install_program)
install_button.pack()

# โปรแกรมติดตั้งรันตลอดเวลา
root.mainloop()

# หลังจากปิด GUI โค้ดอื่นๆ จะถูกรันต่อทันที



