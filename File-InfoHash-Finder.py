# ONLY ISSUE WITH THE PDF EXPORT - CONTENT OVERFLOW
import tkinter as tk
from tkinter import ttk, filedialog
import hashlib
import os
import pandas as pd
from tkinter import messagebox
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

def calculate_hashes():
    file_path = file_path_entry.get()
    if os.path.exists(file_path):
        clear_treeview()
        if os.path.isdir(file_path):
            calculate_directory_hash(file_path)
        elif os.path.isfile(file_path):
            calculate_file_hashes(file_path)
    else:
        clear_treeview()
        result_tree.insert("", "end", values=("Invalid path", "", "", "", "", "", ""))

def calculate_file_hashes(file_path):
    md5_hash = hashlib.md5()
    sha1_hash = hashlib.sha1()
    sha256_hash = hashlib.sha256()

    file_size_bytes = os.path.getsize(file_path)
    file_size_mb = file_size_bytes / (1024 * 1024)  # Convert to MB
    file_size_hr = f"{file_size_mb:.2f} MB"  # Human-readable file size
    _, file_extension = os.path.splitext(file_path)

    with open(file_path, "rb") as file:
        while True:
            data = file.read(8192)
            if not data:
                break
            md5_hash.update(data)
            sha1_hash.update(data)
            sha256_hash.update(data)

    result_tree.insert("", "end", values=(os.path.basename(file_path), file_size_hr, file_path, md5_hash.hexdigest(), sha1_hash.hexdigest(), sha256_hash.hexdigest(), file_extension))

def calculate_directory_hash(directory_path):
    for root, _, files in os.walk(directory_path):
        for filename in files:
            file_path = os.path.join(root, filename)
            calculate_file_hashes(file_path)

def clear_treeview():
    for row in result_tree.get_children():
        result_tree.delete(row)

def export_results(format):
    items = result_tree.get_children()
    if not items:
        messagebox.showinfo("Export", "No results to export.")
        return

    data = []
    for item in items:
        data.append(result_tree.item(item)['values'])

    if format == 'CSV':
        export_to_csv(data)
    elif format == 'XLSX':
        export_to_xlsx(data)
    elif format == 'PDF':
        export_to_pdf(data)

def export_to_csv(data):
    df = pd.DataFrame(data, columns=["Filename", "Filesize", "Filepath", "MD5", "SHA1", "SHA256", "Filetype"])
    df.to_csv("file_hashes.csv", index=False)
    messagebox.showinfo("Export", "CSV file saved as 'file_hashes.csv'.")

def export_to_xlsx(data):
    df = pd.DataFrame(data, columns=["Filename", "Filesize", "Filepath", "MD5", "SHA1", "SHA256", "Filetype"])
    df.to_excel("file_hashes.xlsx", index=False)
    messagebox.showinfo("Export", "XLSX file saved as 'file_hashes.xlsx'.")

def export_to_pdf(data):
    doc = SimpleDocTemplate("file_hashes.pdf", pagesize=letter)
    elements = []

    table_data = [["Filename", "Filesize", "Filepath", "MD5", "SHA1", "SHA256", "Filetype"]] + data

    # Adjust the colWidths to fit your content
    col_widths = [3 * cm] * 7  # You can adjust the column widths as needed

    table = Table(table_data, colWidths=col_widths, rowHeights=[1.5 * cm] * (len(data) + 1))
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('WORDWRAP', (0, 1), (-1, -1), True),  # Enable word wrap for content cells
    ]))

    elements.append(table)
    doc.build(elements)
    messagebox.showinfo("Export", "PDF file saved as 'file_hashes.pdf'.")

# Create the main window
root = tk.Tk()
root.title("File Hash Calculator")
root.geometry("1366x768")  # Set the initial window size

# Create and arrange GUI components
file_path_label = tk.Label(root, text="Enter File/Directory Path:")
file_path_label.pack()

file_path_entry = tk.Entry(root, width=100)
file_path_entry.pack()

browse_button = tk.Button(root, text="Browse", command=lambda: file_path_entry.insert(0, filedialog.askdirectory()))
browse_button.pack()

calculate_button = tk.Button(root, text="Calculate Hashes", command=calculate_hashes)
calculate_button.pack()

# Create a StringVar to store the selected format
selected_format = tk.StringVar()
selected_format.set("CSV")  # Set a default format

# Create a label for the export format
format_label = tk.Label(root, text="Select Export Format:")
format_label.pack()

# Create a combobox to select the export format
format_combobox = ttk.Combobox(root, textvariable=selected_format, values=["CSV", "XLSX", "PDF"])
format_combobox.pack()

# Create the Treeview widget for displaying results in a table
result_tree = ttk.Treeview(root, columns=("Filename", "Filesize", "Filepath", "MD5", "SHA1", "SHA256", "Filetype"),show="headings", height=25)

# Define column headers
result_tree.heading("Filename", text="Filename")
result_tree.heading("Filesize", text="Filesize")
result_tree.heading("Filepath", text="Filepath")
result_tree.heading("MD5", text="MD5")
result_tree.heading("SHA1", text="SHA1")
result_tree.heading("SHA256", text="SHA256")
result_tree.heading("Filetype", text="Filetype")

# Set column widths
result_tree.column("Filename", width=200)
result_tree.column("Filesize", width=200)
result_tree.column("Filepath", width=200)
result_tree.column("MD5", width=200)
result_tree.column("SHA1", width=200)
result_tree.column("SHA256", width=200)
result_tree.column("Filetype", width=200)

result_tree.pack()

# Add a button to export results
export_button = tk.Button(root, text="Export Results", command=lambda: export_results(selected_format.get()))
export_button.pack()

root.mainloop()