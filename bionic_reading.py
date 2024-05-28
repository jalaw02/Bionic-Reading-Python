from tkinter import Tk, Text, Button, END, messagebox
from tkinter import filedialog
from docx import Document
from docx.shared import Pt
import subprocess

def bionic_reading(text):
    lines = text.splitlines()
    formatted_lines = []
    for line in lines:
        words = line.split()
        formatted_words = []
        for word in words:
            if len(word) > 1:
                mid = len(word) // 2
                formatted_word = word[:mid] + word[mid:]
                formatted_words.append((word[:mid], word[mid:]))
            else:
                formatted_words.append((word, ''))
        formatted_lines.append(formatted_words)
    return formatted_lines

def export_to_word(formatted_lines, filename):
    document = Document()
    style = document.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)

    for formatted_words in formatted_lines:
        paragraph = document.add_paragraph()
        for bold_part, normal_part in formatted_words:
            run = paragraph.add_run(bold_part)
            run.bold = True
            run = paragraph.add_run(normal_part)
            run.font.name = 'Candara'
            paragraph.add_run(' ')

    document.save(filename)
    messagebox.showinfo("Success", "The Bionic Reading formatted document has been created successfully.")
    open_file(filename)

def on_convert_click():
    plain_text = text_box.get("1.0", END).strip()
    formatted_lines = bionic_reading(plain_text)
    root.withdraw()  # Hide the main window
    filename = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")], title="Save Document As")
    if filename:
        export_to_word(formatted_lines, filename)
        root.destroy()  # Close the window after conversion

def open_file(filename):
    subprocess.call(["open", filename])

# Create the main window
root = Tk()
root.title("Bionic Reading Converter")

# Create a Text widget for input
text_box = Text(root, height=15, width=50)
text_box.pack(pady=20)

# Create a Button to trigger the conversion
convert_button = Button(root, text="Convert to Bionic Reading", command=on_convert_click)
convert_button.pack(pady=10)

# Run the application
root.mainloop()
