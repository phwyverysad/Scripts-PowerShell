import ctypes
import sys
import subprocess
import os

def run_admin():
    # คำสั่ง PowerShell ของคุณ
    powershell_command = r"""
    $p="C:\phwyverysad"; $z="$p\Mem.Reduct.zip"; $e="$p\Ext"; $d="C:\Program Files"; $u="https://github.com/phwyverysad/-/releases/download/%E0%B8%88%E0%B8%B9%E0%B8%99%E0%B8%84%E0%B8%AD%E0%B8%A1/Mem.Reduct.zip"; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $oldP=$ProgressPreference; $ProgressPreference='SilentlyContinue'; if(!(Test-Path $p)){New-Item $p -Type Directory -Force | Out-Null}; try{Write-Host "Downloading..."; (New-Object System.Net.WebClient).DownloadFile($u,$z); if(Test-Path $z){Write-Host "Extracting..."; Expand-Archive $z $e -Force; $in=Get-ChildItem $e -Directory | Select -First 1; $fd=Join-Path $d "Mem.Reduct"; if(Test-Path $fd){Remove-Item $fd -Recurse -Force}; Move-Item $in.FullName $fd -Force; $sh=New-Object -ComObject WScript.Shell; $ts=@("$env:USERPROFILE\Desktop\Mem Reduct.lnk","$env:USERPROFILE\OneDrive\Desktop\Mem Reduct.lnk"); foreach($t in $ts){if(Test-Path (Split-Path $t)){$sc=$sh.CreateShortcut($t); $sc.TargetPath="$fd\memreduct.exe"; $sc.WorkingDirectory=$fd; $sc.Save();}}}catch{}finally{if(Test-Path $p){Remove-Item $p -Recurse -Force}; $ProgressPreference=$oldP}
    """
    
    # ตรวจสอบสิทธิ์ Admin
    if ctypes.windll.shell32.IsUserAnAdmin():
        # ถ้าเป็น Admin แล้ว ให้รันคำสั่งทันทีแบบ Hidden
        subprocess.run(["powershell", "-WindowStyle", "Hidden", "-Command", powershell_command], 
                       creationflags=subprocess.CREATE_NO_WINDOW)
    else:
        # ถ้ายังไม่เป็น ให้ขอสิทธิ์ Admin โดยไม่มีหน้าต่างโชว์
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{os.path.abspath(sys.argv[0])}"', None, 0)

if __name__ == "__main__":
    run_admin()
