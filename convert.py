import sys
import os
import subprocess
from pdf2docx import Converter
import tkinter as tk
from tkinter import messagebox

def convert_pdf_to_docx(pdf_file):
    if not pdf_file.lower().endswith('.pdf'):
        return

    docx_file = pdf_file.rsplit('.', 1)[0] + '.docx'
    
    try:
        # Perform conversion
        cv = Converter(pdf_file)
        cv.convert(docx_file, start=0, end=None)
        cv.close()
        
        # Create a hidden root window for the popup
        root = tk.Tk()
        root.withdraw()
        
        # Completion popup with "Open File" option
        response = messagebox.askyesno(
            "Conversion Complete", 
            f"Successfully converted:\n{os.path.basename(docx_file)}\n\nWould you like to open it now?"
        )
        
        if response:
            os.startfile(docx_file)
            
        root.destroy()

    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Conversion Error", f"An error occurred while converting:\n{str(e)}")
        root.destroy()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(1)
    
    for arg in sys.argv[1:]:
        convert_pdf_to_docx(arg)
