import PyPDF2
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
from PIL import Image, ImageTk

# --- GUI setup ---
# Configure the main application window
root = tk.Tk()
root.title("PDF Converter Tool - KTS")

# Function to change the window icon
def set_window_icon(icon_path):
    """
    Sets the window icon.
    Requires an .ico file for the title bar.
    """
    try:
        # Get the absolute path of the script's directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        absolute_path = os.path.join(script_dir, icon_path)
        root.iconbitmap(absolute_path)
    except tk.TclError:
        # Fails silently if the icon file is not found or is in the wrong format
        pass

# Set the .ico file for the title bar
set_window_icon('img/icon.ico')

root.geometry("500x300")
root.resizable(False, False)

# Style for the widgets
style = ttk.Style()
style.configure("TButton", font=("Helvetica", 12))
style.configure("TLabel", font=("Helvetica", 12))
style.configure("TCombobox", font=("Helvetica", 12))

# Use a frame to place widgets at the top
top_frame = tk.Frame(root)
top_frame.pack(fill='x', padx=10, pady=5)

# Load the icon for the title label using Pillow
try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(script_dir, 'img/icon.png')
    app_icon = ImageTk.PhotoImage(Image.open(image_path).resize((40, 40)))
except FileNotFoundError:
    app_icon = None

# Application title with an icon
title_label = ttk.Label(top_frame, text=" PDF Converter Tool - KTS", font=("Helvetica", 18, "bold"))
if app_icon:
    title_label.config(image=app_icon, compound="left")
title_label.pack(side='left', padx=10)

# New Info button at the top right
info_button = ttk.Button(top_frame, text="Info", command=lambda: show_info())
info_button.pack(side='right', padx=10)

# Main content frame
main_frame = tk.Frame(root)
main_frame.pack(pady=20)

# Button to select a file
select_button = ttk.Button(main_frame, text="Select PDF", command=lambda: select_file())
select_button.pack(pady=5)

# Label to show the selected file path
file_label = ttk.Label(main_frame, text="No file selected")
file_label.pack(pady=5)

# Dropdown to select the output format
format_label = ttk.Label(main_frame, text="Choose output format:")
format_label.pack(pady=5)
output_formats = ["TXT", "Word (text)", "Excel (text)", "PowerPoint (text)"]
format_combobox = ttk.Combobox(main_frame, values=output_formats, state="readonly")
format_combobox.set("TXT")
format_combobox.pack(pady=5)

# Button to start conversion
convert_button = ttk.Button(main_frame, text="Convert", state="disabled", command=lambda: convert_file_wrapper())
convert_button.pack(pady=10)

# --- GUI functions ---
def select_file():
    """
    Opens a file dialog for file selection and updates the GUI.
    """
    file_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")]
    )
    if file_path:
        file_label.config(text=f"Selected File: {file_path}")
        convert_button.config(state="normal")
    return file_path

def show_info():
    """
    Displays a message box with software information and contact details.
    """
    info_text = (
        "PDF Converter Tool - KTS\n\n"
        "Version: 0.0.3\n"
        "Developed by: KekkoTech Softwares\n\n"
        "This is an open-source tool for basic PDF conversion.\n"
        "It extracts text from PDFs and saves it to other file formats.\n\n"
        "Contact:\n"
        "For feedback: feedback@kekkotech.it\n"
        "For support: support@kekkotech.it\n"
        "Find more on www.kekkotech.com and www.kekkotech.it"
    )
    messagebox.showinfo("About PDF Converter - KTS", info_text)

# --- Conversion logic ---
def extract_text_from_pdf(pdf_path):
    """
    Extracts all text from a given PDF file.
    """
    try:
        with open(pdf_path, 'rb') as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            full_text = ""
            for page in reader.pages:
                full_text += page.extract_text()
        return full_text
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while extracting text: {e}")
        return None

def save_as_txt(text, file_name, save_dir, extension=".txt"):
    """
    Saves the text to a file with the specified extension.
    """
    output_path = os.path.join(save_dir, f"{file_name}{extension}")
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        messagebox.showinfo("Success", f"File successfully converted to {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while writing the file: {e}")

def save_as_docx(text, file_name, save_dir):
    """
    Saves the text to a .docx file.
    """
    output_path = os.path.join(save_dir, f"{file_name}.docx")
    try:
        document = Document()
        document.add_paragraph(text)
        document.save(output_path)
        messagebox.showinfo("Success", f"File successfully converted to {output_path}")
    except ImportError:
        messagebox.showwarning("Warning", "The 'python-docx' library is not installed. The file will be saved as a plain text file instead.")
        save_as_txt(text, file_name, save_dir, extension=".txt")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while writing the file: {e}")

def convert_file(file_path, output_format):
    """
    Handles the file conversion process based on user selection.
    """
    if not file_path:
        messagebox.showwarning("Warning", "Please select a PDF file first.")
        return

    extracted_text = extract_text_from_pdf(file_path)
    if not extracted_text:
        return

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_directory = filedialog.askdirectory(title="Select a destination folder")
    if not output_directory:
        return

    if output_format in ["Excel (text)", "PowerPoint (text)"]:
        response = messagebox.askyesno(
            "Text-Only Conversion",
            f"The '{output_format}' format only supports text extraction. All formatting, tables, and images will be lost. Do you want to continue?",
            icon='warning'
        )
        if not response:
            return

    if output_format == "TXT":
        save_as_txt(extracted_text, base_name, output_directory)
    elif output_format == "Word (text)":
        save_as_docx(extracted_text, base_name, output_directory)
    elif output_format == "Excel (text)":
        save_as_txt(extracted_text, base_name, output_directory, extension=".xlsx")
    elif output_format == "PowerPoint (text)":
        save_as_txt(extracted_text, base_name, output_directory, extension=".pptx")

def convert_file_wrapper():
    """
    A wrapper function to get the selected file path and format from the GUI widgets.
    """
    file_path = file_label.cget("text").replace("Selected File: ", "")
    output_format = format_combobox.get()
    convert_file(file_path, output_format)

# --- Start the application ---
root.mainloop()