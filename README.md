import tkinter as tk
from tkinter import filedialog
from xml.etree import ElementTree as ET
from docx import Document
import openpyxl
from pptx import Presentation
from tkinter import ttk
import csv

def upload_file():
    global file_path
    file_types = [
        ("XML files", "*.xml"),
        ("Word files", "*.docx"),
        ("Excel files", "*.xlsx"),
        ("PowerPoint files", "*.pptx")
    ]
    file_path = filedialog.askopenfilename(filetypes=file_types)
    if file_path:
        label.config(text=f"File Uploaded: {file_path}")
        file_type_var.set(".xml")  # Reset file type to XML after uploading a file

def convert_and_save():
    selected_format = file_type_var.get()
    if not selected_format or not file_path:
        print("Please upload a file and select a file format before converting.")
        return

    try:
        tree = ET.parse(file_path)
        root = tree.getroot()

        if selected_format == ".docx":
            document = Document()
            # Convert XML content to paragraphs in the Word document
            for element in root.iter():
                document.add_paragraph(element.text)
            
            # Save the Word document
            save_path = filedialog.asksaveasfilename(
                defaultextension=selected_format,
                initialfile=f"{file_path.split('/')[-1].split('.')[0]}_{selected_format}",
                filetypes=[(f"{selected_format.upper()} files", f"*{selected_format}")]
            )
            if save_path:
                document.save(save_path)
                print(f"File converted and saved successfully: {save_path}")
                label.config(text=f"File Converted and Saved: {save_path}")
                display_data(save_path)
            else:
                print("File save operation canceled.")

        elif selected_format == ".xlsx":
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            # Convert XML content to cells in the Excel sheet
            for element in root.iter():
                sheet.append([element.text])

            # Save the Excel workbook
            save_path = filedialog.asksaveasfilename(
                defaultextension=selected_format,
                initialfile=f"{file_path.split('/')[-1].split('.')[0]}_{selected_format}",
                filetypes=[(f"{selected_format.upper()} files", f"*{selected_format}")]
            )
            if save_path:
                workbook.save(save_path)
                print(f"File converted and saved successfully: {save_path}")
                label.config(text=f"File Converted and Saved: {save_path}")
                display_data(save_path)
            else:
                print("File save operation canceled.")

        elif selected_format == ".pptx":
            presentation = Presentation()
            # Convert XML content to slides in the PowerPoint presentation
            for element in root.iter():
                slide = presentation.slides.add_slide(presentation.slide_layouts[0])
                slide.shapes.title.text = element.text

            # Save the PowerPoint presentation
            save_path = filedialog.asksaveasfilename(
                defaultextension=selected_format,
                initialfile=f"{file_path.split('/')[-1].split('.')[0]}_{selected_format}",
                filetypes=[(f"{selected_format.upper()} files", f"*{selected_format}")]
            )
            if save_path:
                presentation.save(save_path)
                print(f"File converted and saved successfully: {save_path}")
                label.config(text=f"File Converted and Saved: {save_path}")
                display_data(save_path)
            else:
                print("File save operation canceled.")

        else:
            print("Unsupported file format.")

    except Exception as e:
        print(f"Error converting and saving file: {e}")

def display_data(file_path):
    data_window = tk.Toplevel(window)
    data_window.title("Data Display")

    tree = ttk.Treeview(data_window)
    tree["columns"] = ("value")
    tree.heading("#0", text="Element")
    tree.heading("value", text="Value")

    try:
        tree_data(file_path, tree)
    except Exception as e:
        print(f"Error displaying data: {e}")

    tree.pack(expand=True, fill=tk.BOTH)

    save_button = tk.Button(data_window, text="Save Data", command=lambda: save_data(tree, file_path), font=("Arial", 14))
    save_button.pack(pady=10)

def save_data(tree, file_path):
    save_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")]
    )
    if save_path:
        with open(save_path, 'w', newline='', encoding='utf-8') as file:
            csv_writer = csv.writer(file)
            write_tree_to_csv(tree, csv_writer)

        print(f"Data saved to CSV file: {save_path}")

def write_tree_to_csv(tree, csv_writer, parent=""):
    csv_writer.writerow(["Element", "Value"])
    for child_id in tree.get_children(parent):
        values = tree.item(child_id, 'values')
        csv_writer.writerow([tree.item(child_id, 'text'), values[0]])
        write_tree_to_csv(tree, csv_writer, child_id)

def tree_data(file_path, tree, parent=""):
    tree.delete(*tree.get_children(parent))
    tree.insert(parent, "end", text=file_path, values=(""))

    if file_path.endswith(".xml"):
        tree_data_xml(file_path, tree, parent)
    elif file_path.endswith(".docx"):
        tree_data_docx(file_path, tree, parent)
    elif file_path.endswith(".xlsx"):
        tree_data_xlsx(file_path, tree, parent)
    elif file_path.endswith(".pptx"):
        tree_data_pptx(file_path, tree, parent)

def tree_data_xml(file_path, tree, parent=""):
    tree.delete(*tree.get_children(parent))
    tree.insert(parent, "end", text=file_path, values=(""))

    tree_data_xml_recursive(file_path, tree, parent)

def tree_data_xml_recursive(file_path, tree, parent):
    tree_data = ET.parse(file_path)
    root = tree_data.getroot()

    for element in root.iter():
        tree.insert(parent, "end", text=element.tag, values=(element.text,))
        if len(element) > 0:
            tree_data_xml_recursive(file_path, tree, element.tag)

def tree_data_docx(file_path, tree, parent=""):
    tree.delete(*tree.get_children(parent))
    tree.insert(parent, "end", text=file_path, values=(""))

    doc = Document(file_path)

    for paragraph in doc.paragraphs:
        tree.insert(parent, "end", text="Paragraph", values=(paragraph.text,))

def tree_data_xlsx(file_path, tree, parent=""):
    tree.delete(*tree.get_children(parent))
    tree.insert(parent, "end", text=file_path, values=(""))

    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    for row in sheet.iter_rows(min_row=1, values_only=True):
        for value in row:
            tree.insert(parent, "end", text="Cell", values=(value,))

def tree_data_pptx(file_path, tree, parent=""):
    tree.delete(*tree.get_children(parent))
    tree.insert(parent, "end", text=file_path, values=(""))

    presentation = Presentation(file_path)

    for i, slide in enumerate(presentation.slides):
        tree.insert(parent, "end", text=f"Slide {i + 1}", values=("",))
        for shape in slide.shapes:
            if shape.has_text_frame:
                tree.insert(parent, "end", text="Shape", values=(shape.text,))

# Create the main window
window = tk.Tk()
window.title("File Upload and Convert UI")

# Create a label
label = tk.Label(window, text="No file uploaded", font=("Arial", 18))
label.pack(pady=20)

# Create a button for file upload
upload_button = tk.Button(window, text="Upload File", command=upload_file, font=("Arial", 16))
upload_button.pack(pady=10)

# Create a combo box for file types
file_type_var = tk.StringVar()
file_type_var.set(".xml")  # Default selection
file_types = [".docx", ".xlsx", ".pptx"]
file_type_combo = tk.OptionMenu(window, file_type_var, *file_types)
file_type_combo.pack(pady=10)
file_type_label = tk.Label(window, text="Select File Format:", font=("Arial", 14))
file_type_label.pack()

# Create a button for converting and saving the file
convert_button = tk.Button(window, text="Convert and Save", command=convert_and_save, font=("Arial", 16))
convert_button.pack(pady=10)

# Global variable to store the uploaded file path
file_path = None

# Start the main event loop
window.mainloop()
