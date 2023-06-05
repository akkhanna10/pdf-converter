import os
import comtypes.client
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

def ppt_to_pdf(input_file, output_file):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    presentation = powerpoint.Presentations.Open(str(input_file))
    presentation.ExportAsFixedFormat(str(output_file), 2)  # 2 for PDF format
    presentation.Close()

    powerpoint.Quit()

def excel_to_pdf(input_file, output_file):
    excel = comtypes.client.CreateObject("Excel.Application")
    excel.Visible = 0

    workbook = excel.Workbooks.Open(str(input_file))
    workbook.ExportAsFixedFormat(0, str(output_file))  # 0 for PDF format
    workbook.Close()

    excel.Quit()

def word_to_pdf(input_file, output_file):
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = 0

    doc = word.Documents.Open(str(input_file))
    doc.SaveAs(str(output_file), FileFormat=17)  # 17 for PDF format
    doc.Close()

    word.Quit()

def convert_to_pdf(input_path, output_path):
    input_file = Path(input_path)
    output_file = Path(output_path)

    if input_file.suffix in ['.ppt', '.pptx']:
        ppt_to_pdf(input_file, output_file)
    elif input_file.suffix in ['.xls', '.xlsx']:
        excel_to_pdf(input_file, output_file)
    elif input_file.suffix in ['.doc', '.docx']:
        word_to_pdf(input_file, output_file)
    else:
        messagebox.showerror("Error", "Unsupported file format.")

def browse_input_file():
    file_path = filedialog.askopenfilename()
    input_entry.delete(0, tk.END)
    input_entry.insert(tk.END, file_path)

def browse_output_folder():
    folder_path = filedialog.askdirectory()
    output_entry.delete(0, tk.END)
    output_entry.insert(tk.END, folder_path)

def convert_file():
    input_file = input_entry.get()
    output_folder = output_entry.get()

    if not input_file or not output_folder:
        messagebox.showerror("Error", "Please select input file and output folder.")
        return

    output_file = os.path.join(output_folder, Path(input_file).stem + ".pdf")
    
    try:
        convert_to_pdf(input_file, output_file)
        messagebox.showinfo("Success", "File converted to PDF successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Create the main window
window = tk.Tk()
window.title("File Converter")

# Input File
input_label = tk.Label(window, text="Input File:")
input_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
input_entry = tk.Entry(window, width=50)
input_entry.grid(row=0, column=1, padx=5, pady=5)
input_button = tk.Button(window, text="Browse", command=browse_input_file)
input_button.grid(row=0, column=2, padx=5, pady=5)

# Output Folder
output_label = tk.Label(window, text="Output Folder:")
output_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
output_entry = tk.Entry(window, width=50)
output_entry.grid(row=1, column=1, padx=5, pady=5)
output_button = tk.Button(window, text="Browse", command=browse_output_folder)
output_button.grid(row=1, column=2, padx=5, pady=5)

# Convert Button
convert_button = tk.Button(window, text="Convert", command=convert_file)
convert_button.grid(row=2, column=1, padx=5, pady=10)

# Start the main loop
window.mainloop()
