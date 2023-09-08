#_______________________________________________________________________
#----------------------ส่วนของ GUI---------------------------------------
#_______________________________________________________________________
import tkinter as tk
from tkinter import filedialog

# อ่านที่อยู่ไฟล์ล่าสุดจากไฟล์ข้อความ (PathData.txt) ถ้ามี
try:
    with open("PathData.txt", "r", encoding="utf-8") as file:
        Path = file.read()
except FileNotFoundError:
    Path = ""

# ถ้ามีข้อมูลใน "PathData.txt" ไม่ต้องแสดง GUI
if not Path:
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

    # เริ่มการรัน GUI
    root.mainloop()