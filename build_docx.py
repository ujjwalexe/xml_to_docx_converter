import os
import zipfile
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import sys

#Check if folder has the required Word XML structure
def is_valid_docx_structure(folder_path):
    required_files = [
        "[Content_Types].xml",
        os.path.join("_rels", ".rels"),
        os.path.join("word", "document.xml")
    ]
    for file in required_files:
        if not os.path.exists(os.path.join(folder_path, file)):
            print(f" Missing: {file}")
            return False
    return True

#Zipping the folder and renaming it to .docx
def create_docx_from_xml(source_folder, output_docx_path):
    zip_path = output_docx_path.replace('.docx', '.zip')

# Zip the folder
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                full_path = os.path.join(root, file)
                arcname = os.path.relpath(full_path, source_folder)
                zipf.write(full_path, arcname)
                print(f" Zipping: {arcname}")

# Rename .zip to .docx
    if os.path.exists(output_docx_path):
        os.remove(output_docx_path)
    shutil.move(zip_path, output_docx_path)
    print(f" DOCX created: {output_docx_path}")
# Open the DOCX file 
    print(" Opening the DOCX file...")
    os.startfile(output_docx_path)

# ------------------ GUI ------------------

def browse_folder():
    folder_selected = filedialog.askdirectory()
    folder_path_var.set(folder_selected)

def convert():
    folder_path = folder_path_var.get()
    if not folder_path:
        messagebox.showerror("Error", "Please select a folder.")
        return

    if not os.path.exists(folder_path):
        messagebox.showerror("Error", "Folder does not exist.")
        return

    if not is_valid_docx_structure(folder_path):
        messagebox.showerror("Invalid", "Folder is missing required Word XML files.")
        return

    output_path = os.path.join(folder_path, "converted.docx")
    try:
        create_docx_from_xml(folder_path, output_path)
        messagebox.showinfo("Success", f"DOCX created:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# UI Setup
root = tk.Tk()
root.title("XML → DOCX Converter")
root.geometry("500x200")
root.resizable(False, False)

folder_path_var = tk.StringVar()

label = tk.Label(root, text="Select Folder with Word XML Structure:", font=("Arial", 12))
label.pack(pady=10)

entry = tk.Entry(root, textvariable=folder_path_var, width=60)
entry.pack(pady=5)

browse_btn = tk.Button(root, text="Browse", command=browse_folder)
browse_btn.pack(pady=5)

convert_btn = tk.Button(root, text="Convert to .docx", bg="green", fg="white", font=("Arial", 11, "bold"), command=convert)
convert_btn.pack(pady=10)

root.mainloop()
