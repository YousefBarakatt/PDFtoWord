import winreg as reg
import os
import shutil
import tkinter as tk
from tkinter import messagebox

def remove_context_menu():
    key_path = r'Software\Classes\SystemFileAssociations\.pdf\shell\PDFtoWord'
    appdata_path = os.path.join(os.environ['LOCALAPPDATA'], 'PDFtoWord')
    
    try:
        # Delete subkeys first
        reg.DeleteKey(reg.HKEY_CURRENT_USER, f"{key_path}\\command")
        # Then delete the main key
        reg.DeleteKey(reg.HKEY_CURRENT_USER, key_path)
        
        # 2. Remove application files
        if os.path.exists(appdata_path):
            # We use rmtree with caution
            shutil.rmtree(appdata_path)
        
        # 3. Message for user
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Uninstallation Successful", "PDF to Word Converter has been completely removed from your system.")
        root.destroy()
        
    except FileNotFoundError:
        pass
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Uninstallation Error", f"An error occurred during cleanup:\n{str(e)}")
        root.destroy()

if __name__ == "__main__":
    remove_context_menu()
