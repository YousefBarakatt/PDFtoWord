import os
import sys
import winreg as reg
import shutil
import tkinter as tk
from tkinter import messagebox

def setup():
    # 1. Define safe installation path in AppData
    appdata_path = os.path.join(os.environ['LOCALAPPDATA'], 'PDFtoWord')
    
    # 2. Paths
    # When running as a PyInstaller bundle, bundled files are in sys._MEIPASS
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    src_convert_exe = os.path.join(base_path, 'convert.exe')
    src_uninstall_exe = os.path.join(base_path, 'uninstall.exe')
    
    dest_convert_exe = os.path.join(appdata_path, 'convert.exe')
    dest_uninstall_exe = os.path.join(appdata_path, 'uninstall.exe')

    try:
        # Create directory if it doesn't exist
        if not os.path.exists(appdata_path):
            os.makedirs(appdata_path)

        # Copy the engine and uninstaller to the permanent location
        if os.path.exists(src_convert_exe):
            shutil.copy2(src_convert_exe, dest_convert_exe)
        if os.path.exists(src_uninstall_exe):
            shutil.copy2(src_uninstall_exe, dest_uninstall_exe)

        # 3. Add to Registry
        key_path = r'Software\Classes\SystemFileAssociations\.pdf\shell\PDFtoWord'
        key = reg.CreateKeyEx(reg.HKEY_CURRENT_USER, key_path, 0, reg.KEY_SET_VALUE)
        reg.SetValueEx(key, '', 0, reg.REG_SZ, 'Convert to Word')
        
        # Add Icon (try to find Word)
        word_paths = [
            r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
            r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"
        ]
        
        icon_path = None
        for path in word_paths:
            if os.path.exists(path):
                icon_path = path
                break
        
        if icon_path:
            reg.SetValueEx(key, 'Icon', 0, reg.REG_SZ, f'"{icon_path}",0')

        # Create the command subkey
        command_key = reg.CreateKeyEx(key, 'command', 0, reg.KEY_SET_VALUE)
        # Use the permanent AppData path
        command = f'"{dest_convert_exe}" "%1"'
        reg.SetValueEx(command_key, '', 0, reg.REG_SZ, command)

        # 4. Success Message
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Success", "PDF to Word Converter installed successfully!\n\nYou can now right-click any PDF in File Explorer and select 'Convert to Word'.")
        root.destroy()

    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Installation Error", f"Failed to install:\n{str(e)}")
        root.destroy()

if __name__ == "__main__":
    setup()
