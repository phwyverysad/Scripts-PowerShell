import os
import urllib.request
import zipfile
import shutil
import win32com.client # ต้องติดตั้งผ่าน pip install pywin32

# กำหนดตัวแปรต่างๆ
p = r"C:\phwyverysad"
z = os.path.join(p, "Mem.Reduct.zip")
e = os.path.join(p, "Ext")
d = r"C:\Program Files"
u = "https://github.com/phwyverysad/-/releases/download/%E0%B8%88%E0%B8%B9%E0%B8%99%E0%B8%84%E0%B8%AD%E0%B8%A1/Mem.Reduct.zip"

# สร้างโฟลเดอร์ชั่วคราวถ้ายังไม่มี
if not os.path.exists(p):
    os.makedirs(p)

try:
    print("Downloading...")
    # ดาวน์โหลดไฟล์ zip (Python urllib จัดการเรื่อง TLS ให้โดยอัตโนมัติ)
    urllib.request.urlretrieve(u, z)

    if os.path.exists(z):
        print("Extracting...")
        # แตกไฟล์ zip ไปที่โฟลเดอร์ Ext
        with zipfile.ZipFile(z, 'r') as zip_ref:
            zip_ref.extractall(e)

        # หาชื่อโฟลเดอร์แรกที่ได้จากการแตกไฟล์ (เทียบเท่า Select -First 1)
        extracted_items = [os.path.join(e, item) for item in os.listdir(e)]
        extracted_dirs = [item for item in extracted_items if os.path.isdir(item)]

        if extracted_dirs:
            in_dir = extracted_dirs[0]
            fd = os.path.join(d, "Mem.Reduct")

            # ถ้ามีโฟลเดอร์เดิมอยู่ใน Program Files ให้ลบทิ้งก่อน
            if os.path.exists(fd):
                shutil.rmtree(fd)

            # ย้ายโฟลเดอร์ที่แตกไฟล์แล้วไปที่ Program Files
            shutil.move(in_dir, fd)

            # สร้าง Shortcut
            shell = win32com.client.Dispatch("WScript.Shell")
            user_profile = os.environ.get("USERPROFILE")
            
            # กำหนด Path สำหรับสร้าง Shortcut (Desktop ปกติ และ OneDrive Desktop)
            shortcuts = [
                os.path.join(user_profile, "Desktop", "Mem Reduct.lnk"),
                os.path.join(user_profile, "OneDrive", "Desktop", "Mem Reduct.lnk")
            ]

            for t in shortcuts:
                target_dir = os.path.dirname(t)
                # ตรวจสอบว่ามีโฟลเดอร์ Desktop ปลายทางอยู่จริงหรือไม่
                if os.path.exists(target_dir):
                    shortcut = shell.CreateShortCut(t)
                    shortcut.Targetpath = os.path.join(fd, "memreduct.exe")
                    shortcut.WorkingDirectory = fd
                    shortcut.Save()
                    print(f"Shortcut Created: {t}")

except Exception as ex:
    print(f"Error: {ex}")

finally:
    print("Cleaning up...")
    # ลบโฟลเดอร์ชั่วคราวทิ้ง (เทียบเท่า Remove-Item -Recurse -Force)
    if os.path.exists(p):
        try:
            shutil.rmtree(p)
        except Exception as cleanup_err:
            pass # ซ่อน Error กรณีลบไฟล์ไม่ได้ เหมือนตอนใช้ SilentlyContinue
    
    print("Done.")
