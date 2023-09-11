import tkinter as tk
from tkinter import filedialog
import os
import datetime

# สร้างตัวแปร marker_file เพื่อเก็บตำแหน่งของ marker.txt
marker_file = "marker.txt"

# ฟังก์ชันเมื่อกดปุ่ม "เลือกไฟล์"
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Executable files", "*.exe")])
    program_path_entry.delete(0, tk.END)
    program_path_entry.insert(0, file_path)

# ฟังก์ชันเมื่อกดปุ่ม "บันทึก"
def save_path():
    program_path = program_path_entry.get()
    
    # สร้างไฟล์ marker.txt
    with open(marker_file, "w") as f:
        f.write("This is a marker file.")
    
    # เก็บตำแหน่งของ marker.txt ในตัวแปร marker_file
    with open("program_path.txt", "w") as f:
        f.write(program_path)
        f.write("\n")  # เพิ่มบรรทัดใหม่
        f.write(marker_file)  # เพิ่มตำแหน่งของ marker.txt
    
    window.destroy()

# ตรวจสอบว่าไฟล์ program_path.txt มีข้อมูลหรือไม่
if os.path.exists("program_path.txt") and os.path.getsize("program_path.txt") > 0:
    with open("program_path.txt", "r") as f:
        program_path = f.readline().strip()
    
    # ตรวจสอบว่า marker.txt มีอยู่หรือไม่
    if not os.path.exists(marker_file):
        # ถ้าไม่มี marker.txt ให้รันโปรแกรม
        os.system(program_path)
        
        # สร้าง marker.txt เพื่อบ่งชี้ว่าโปรแกรมได้รับการรันแล้ว
        with open(marker_file, "w") as file:
            file.write(str(datetime.date.today()))
    else:
        with open(marker_file, "r") as file:
            last_run_date = file.read()
        
        today = str(datetime.date.today())
        
        if last_run_date != today:
            # ถ้าวันที่ไม่ตรงกัน ให้รันโปรแกรมและอัปเดต marker.txt
            os.system(program_path)
            with open(marker_file, "w") as file:
                file.write(today)

    print("สำเร็จ")
else:
    # สร้างหน้าต่างหลัก
    window = tk.Tk()
    window.title("เลือกตำแหน่งไฟล์")
    window.geometry("400x150")

    # สร้างเล้าอินพุทสำหรับระบุตำแหน่งไฟล์
    program_path_label = tk.Label(window, text="ตำแหน่งไฟล์:")
    program_path_label.pack()

    program_path_entry = tk.Entry(window, width=40)
    program_path_entry.pack()
    
    browse_button = tk.Button(window, text="เลือกไฟล์", command=browse_file)
    browse_button.pack()

    # สร้างปุ่ม "บันทึก"
    save_button = tk.Button(window, text="บันทึก", command=save_path)
    save_button.pack()

    window.mainloop()
