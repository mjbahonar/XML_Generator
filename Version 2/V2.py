import tkinter as tk
from tkinter import filedialog
import pandas as pd
import xml.etree.ElementTree as ET
import os

def update_status(message):
    status_label.config(text=message)

def open_file():
    global excel_file_path, default_xml_name
    excel_file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")]
    )
    if excel_file_path:
        # Extract the base name for the default XML file name
        default_xml_name = os.path.splitext(os.path.basename(excel_file_path))[0] + ".xml"
        update_status(f"Selected file: {excel_file_path}")
    else:
        update_status("No file selected. Please try again.")

def generate_xml():
    if not excel_file_path:
        update_status("Error: Please select an Excel file first!")
        return
    
    save_path = filedialog.asksaveasfilename(
        initialfile=default_xml_name,
        defaultextension=".xml",
        filetypes=[("XML files", "*.xml")]
    )
    
    if save_path:
        try:
            # Read the Excel file
            xl = pd.ExcelFile(excel_file_path)
            
            # Parse the sheets
            journal_df = xl.parse('Journal')
            article_df = xl.parse('Article')
            author_df = xl.parse('Author(s)')

            # Create the root XML element
            root = ET.Element("article")

            # Create journal element
            journal_element = ET.SubElement(root, "journal")
            for _, row in journal_df.iterrows():
                ET.SubElement(journal_element, row[0]).text = str(row[1]) if pd.notna(row[1]) else ""

            # Create article element
            article_element = ET.SubElement(root, "article_info")
            for _, row in article_df.iterrows():
                ET.SubElement(article_element, row[0]).text = str(row[1]) if pd.notna(row[1]) else ""

            # Create authors list element
            authors_element = ET.SubElement(root, "author_list")
            for _, row in author_df.iterrows():
                author_element = ET.SubElement(authors_element, "author")
                ET.SubElement(author_element, row[0]).text = str(row[1]) if pd.notna(row[1]) else ""

            # Write to XML file
            tree = ET.ElementTree(root)
            tree.write(save_path, encoding='utf-8', xml_declaration=True)
            update_status(f"Success: XML file generated and saved to: {save_path}")
        except Exception as e:
            update_status(f"Error: Failed to generate XML: {e}")

# Setup the main window
root = tk.Tk()
root.title("Excel to XML Converter")
root.geometry("300x300")  # Set the initial size to 300x300

# Menu setup
menu = tk.Menu(root)
root.config(menu=menu)
file_menu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Open", command=open_file)
file_menu.add_command(label="Generate", command=generate_xml)

# Button setup
btn_open = tk.Button(root, text="Open", command=open_file)
btn_open.pack(pady=10)

btn_generate = tk.Button(root, text="Generate", command=generate_xml)
btn_generate.pack(pady=10)

# Status label to show messages
status_label = tk.Label(root, text="Please open an Excel file", fg="blue")
status_label.pack(pady=10)

# Initialize global variables
excel_file_path = None
default_xml_name = "output.xml"

# Run the application
root.mainloop()
