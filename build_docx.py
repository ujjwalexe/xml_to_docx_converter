import os
import zipfile
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import sys

#Core logic

# Checking if folder has the required Word XML structure
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

# Zipping the folder and renaming it to .docx
def create_docx_from_xml(source_folder, output_docx_path):
    zip_path = output_docx_path.replace('.docx', '.zip')

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                full_path = os.path.join(root, file)
                arcname = os.path.relpath(full_path, source_folder)
                zipf.write(full_path, arcname)
                print(f" Zipping: {arcname}")

    if os.path.exists(output_docx_path):
        os.remove(output_docx_path)
    shutil.move(zip_path, output_docx_path)
    print(f" DOCX created: {output_docx_path}")

    try:
        os.startfile(output_docx_path)
    except Exception:
        print(" Open manually:", output_docx_path)

# Extracting .docx back to folder
def extract_docx_to_folder(docx_path, output_folder):
    if not zipfile.is_zipfile(docx_path):
        raise ValueError("Selected file is not a valid DOCX (ZIP) file.")
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(output_folder)
    print(f" Extracted to: {output_folder}")



# GUI starts here

def browse_folder():
    folder_selected = filedialog.askdirectory()
    folder_path_var.set(folder_selected)

def browse_docx_file():
    file_selected = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    docx_file_path_var.set(file_selected)

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

def extract():
    docx_path = docx_file_path_var.get()
    if not docx_path or not os.path.exists(docx_path):
        messagebox.showerror("Error", "Please select a valid DOCX file.")
        return

    output_folder = filedialog.askdirectory(title="Select Output Folder")
    if not output_folder:
        return

    output_path = os.path.join(
        output_folder, os.path.splitext(os.path.basename(docx_path))[0] + "_unzipped"
    )
    os.makedirs(output_path, exist_ok=True)

    try:
        extract_docx_to_folder(docx_path, output_path)
        messagebox.showinfo("Success", f"Extracted to:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# UI setup 

root = tk.Tk()
root.title("XML ↔ DOCX Converter")
root.geometry("550x400")
root.resizable(False, False)

# Section 1 – Convert Folder to DOCX
folder_path_var = tk.StringVar()

label1 = tk.Label(root, text="Select Folder with Word XML Structure:", font=("Arial", 12))
label1.pack(pady=10)

entry1 = tk.Entry(root, textvariable=folder_path_var, width=60)
entry1.pack(pady=5)

browse_btn = tk.Button(root, text="Browse Folder", command=browse_folder)
browse_btn.pack(pady=5)

convert_btn = tk.Button(root, text="Convert to .docx", bg="green", fg="white", font=("Arial", 11, "bold"), command=convert)
convert_btn.pack(pady=10)


separator = tk.Label(root, text="—" * 100)
separator.pack(pady=10)

# Section 2 – Extract DOCX to Folder
docx_file_path_var = tk.StringVar()

label2 = tk.Label(root, text="Select a .docx file to extract:", font=("Arial", 12))
label2.pack(pady=5)

entry2 = tk.Entry(root, textvariable=docx_file_path_var, width=60)
entry2.pack(pady=5)

browse_docx_btn = tk.Button(root, text="Browse .docx", command=browse_docx_file)
browse_docx_btn.pack(pady=5)

extract_btn = tk.Button(root, text="Extract to Folder", bg="blue", fg="white", font=("Arial", 11, "bold"), command=extract)
extract_btn.pack(pady=10)

root.mainloop()
