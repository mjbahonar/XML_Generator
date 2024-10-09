import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import xml.etree.ElementTree as ET
import os

def open_file():
    global excel_file_path
    excel_file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")])
    if excel_file_path:
        messagebox.showinfo("File Selected", f"Selected file: {excel_file_path}")

def generate_xml():
    if not excel_file_path:
        messagebox.showerror("Error", "Please select an Excel file first!")
        return
    
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xml",
        filetypes=[("XML files", "*.xml")])
    
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
            messagebox.showinfo("Success", f"XML file generated and saved to: {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate XML: {e}")

# Setup the main window
root = tk.Tk()
root.title("Excel to XML Converter")

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

excel_file_path = None

# Run the application
root.mainloop()
