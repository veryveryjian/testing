import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import os

def on_drop(event):
    file_path = event.data
    if file_path:
        file_path = file_path.replace('{', '').replace('}', '')  # Windows 경로에서 '{}' 제거
        if os.path.isfile(file_path) and (file_path.endswith('.xlsx') or file_path.endswith('.xls') or file_path.endswith('.csv')):
            label.config(text=f"File Dropped: {file_path}")
            entry.delete(0, tk.END)
            entry.insert(0, file_path)
        else:
            label.config(text="Please drop an Excel or CSV file.")
    else:
        label.config(text="Failed to get the file path.")

root = TkinterDnD.Tk()
root.title("Drag and Drop Excel File")
root.geometry("400x200")

label = tk.Label(root, text="Drag and drop an Excel or CSV file here")
label.pack(pady=20)

entry = tk.Entry(root, width=50)
entry.pack(pady=10)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)

root.mainloop()
