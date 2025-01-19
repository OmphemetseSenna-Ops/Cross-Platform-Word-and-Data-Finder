import os
import shutil
import glob
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import time
import re
import pandas as pd
import docx

def search_and_copy_files(source_folder, keyword, file_type):
    # Map file type to appropriate glob pattern
    file_patterns = {
        'Text Files': '*.txt',
        'Word Documents': '*.docx',
        'Excel Spreadsheets': '*.xlsx'
    }
    
    search_pattern = os.path.join(source_folder, file_patterns[file_type])
    files = glob.glob(search_pattern)
    
    # Counter for files containing the keyword
    count = 0
    
    # Compile the regex for whole word matching
    regex = re.compile(r'\b' + re.escape(keyword) + r'\b')
    
    # Files to be copied
    files_to_copy = []
    
    # Iterate through each file
    for file_path in files:
        print(f"Processing file: {file_path}")
        try:
            if file_type == 'Text Files':
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
            elif file_type == 'Word Documents':
                doc = docx.Document(file_path)
                content = '\n'.join([para.text for para in doc.paragraphs])
            elif file_type == 'Excel Spreadsheets':
                content = ''
                excel_data = pd.ExcelFile(file_path)
                for sheet_name in excel_data.sheet_names:
                    sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
                    content += sheet_data.to_string()
            
            print(f"File content: {content}")  # Debugging: Print file content
            
            if regex.search(content):
                print(f"Keyword '{keyword}' found in file: {file_path}")  # Debugging: Print keyword detection
                files_to_copy.append(file_path)
                count += 1
        except Exception as e:
            print(f"Error reading file {file_path}: {e}")
    
    # If files containing the keyword are found, create the destination folder and copy the files
    if count > 0:
        destination_folder = os.path.join(source_folder, keyword)
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
        
        for file_path in files_to_copy:
            file_name = os.path.basename(file_path)
            destination_path = os.path.join(destination_folder, file_name)
            for attempt in range(5):  # Try up to 5 times to copy the file
                try:
                    shutil.copy(file_path, destination_path)
                    print(f"Copied file: {file_path} to {destination_path}")
                    break
                except PermissionError:
                    print(f"PermissionError: Attempt {attempt + 1} to copy file: {file_path}")  # Debugging: Log each attempt
                    time.sleep(1)  # Wait a bit before retrying
                    continue
            else:
                print(f"Failed to copy file after several attempts: {file_path}")
    
    return count

def browse_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, folder_path)

def start_search():
    folder_path = folder_entry.get()
    keyword = keyword_entry.get()
    file_type = file_type_var.get()
    
    if not folder_path:
        messagebox.showwarning("Input Error", "Please select a folder.")
        return
    if not keyword:
        messagebox.showwarning("Input Error", "Please enter a keyword to search.")
        return
    if not file_type:
        messagebox.showwarning("Input Error", "Please select a file type.")
        return

    num_files_found = search_and_copy_files(folder_path, keyword, file_type)
    
    if num_files_found > 0:
        messagebox.showinfo("Search Complete", f"Number of files containing '{keyword}': {num_files_found}")
    else:
        messagebox.showinfo("Search Complete", f"No files containing '{keyword}' found.")

# Create the main window
root = tk.Tk()
root.title("File Search and Copy")

# Folder selection
folder_label = tk.Label(root, text="Select Folder:")
folder_label.grid(row=0, column=0, padx=10, pady=10)

folder_entry = tk.Entry(root, width=50)
folder_entry.grid(row=0, column=1, padx=10, pady=10)

browse_button = tk.Button(root, text="Browse", command=browse_folder)
browse_button.grid(row=0, column=2, padx=10, pady=10)

# Keyword entry
keyword_label = tk.Label(root, text="Enter Keyword:")
keyword_label.grid(row=1, column=0, padx=10, pady=10)

keyword_entry = tk.Entry(root, width=50)
keyword_entry.grid(row=1, column=1, padx=10, pady=10)

# File type selection
file_type_label = tk.Label(root, text="Select File Type:")
file_type_label.grid(row=2, column=0, padx=10, pady=10)

file_type_var = tk.StringVar(root)
file_type_var.set("Text Files")  # default value
file_type_dropdown = tk.OptionMenu(root, file_type_var, "Text Files", "Word Documents", "Excel Spreadsheets")
file_type_dropdown.grid(row=2, column=1, padx=10, pady=10)

# Search button
search_button = tk.Button(root, text="Search and Copy", command=start_search)
search_button.grid(row=3, column=1, padx=10, pady=10)

# Start the GUI event loop
root.mainloop()
