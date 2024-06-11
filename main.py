import tkinter as tk
import pathlib
import os
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from PIL import Image, ImageTk

# Function to change button background color on hover
def on_enter(event):
    event.widget.config(bg='darkorange')

def on_leave(event):
    event.widget.config(bg='orange1')

# Function to convert PDF to Word
def convert_pdf_to_word(pdf_file, word_file):
    try:
        cv = Converter(pdf_file)
        cv.convert(word_file)
        cv.close()
    except Exception as e:
        print(f"Conversion Error: {e}")
        raise e

# Function to select a PDF file
def select_pdf_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    pdf_entry.delete(0, tk.END)
    pdf_entry.insert(0, file_path)

# Function to handle the conversion process
def convert():
    pdf_file = pdf_entry.get()

    if not pdf_file:
        messagebox.showwarning("Input Error", "Please select a PDF file.")
        return

    # Automatically set the Word file path to the desktop
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    word_file = os.path.join(desktop_path, os.path.splitext(os.path.basename(pdf_file))[0] + ".docx")

    try:
        convert_pdf_to_word(pdf_file, word_file)
        messagebox.showinfo("Success", f"Converted {pdf_file} to {word_file} successfully!")
    except Exception as e:
        messagebox.showerror("Conversion Error", f"An error occurred: {e}")

# Function to load the application logo
def load_logo(image_path):
    try:
        if image_path.exists():
            image = Image.open(image_path)
            icon_image = ImageTk.PhotoImage(image)
            root.wm_iconphoto(False, icon_image)
        else:
            print("Image Error: logo.png not found")
    except Exception as e:
        print("Error occurred while loading logo:", e)

# Main application setup
def setup_main_frame(root):
    root.title("PDF to Word Converter")
    root.geometry("860x350")
    root.resizable(False, False)
    root.configure(bg="#242124")

    # Application logo
    image_path = pathlib.Path(r"logo.png")
    load_logo(image_path)

    # Title Label
    tk.Label(root, text="PDF to Word Converter", font="Helvetica 30 bold", bg="#242124", fg="orange1").place(relx=0.05, rely=0.05)

    # File Path Label
    tk.Label(root, text="File Path (.pdf) ➡️", font="Helvetica 17", bg="#242124", fg="Ghost White").place(relx=0.05, rely=0.40)

    # File path entry for PDF file
    global pdf_entry
    pdf_entry = tk.Entry(root, width=45)
    pdf_entry.place(relx=0.30, rely=0.41)

    # Browse button for PDF file
    browse_pdf_btn = tk.Button(root, text="◀ BROWSE ▶", font="Helvetica 13 bold", relief="raised", bg="orange1", command=select_pdf_file)
    browse_pdf_btn.place(relx=0.80, rely=0.39)
    browse_pdf_btn.bind("<Enter>", on_enter)
    browse_pdf_btn.bind("<Leave>", on_leave)

    # Convert button
    convert_btn = tk.Button(root, text="CONVERT ✅", font="Helvetica 16 bold", relief="raised", bg="orange1", command=convert)
    convert_btn.place(relx=0.40, rely=0.70)
    convert_btn.bind("<Enter>", on_enter)
    convert_btn.bind("<Leave>", on_leave)

# Initialize the main frame
if __name__ == "__main__":
    root = tk.Tk()
    setup_main_frame(root)
    root.mainloop()  