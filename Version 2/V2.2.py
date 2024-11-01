import tkinter as tk
from tkinter import filedialog
import pandas as pd
from lxml import etree  # Using lxml for better XML handling
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
            root = etree.Element("article")

            # Create journal element
            journal_element = etree.SubElement(root, "journal")
            for _, row in journal_df.iterrows():
                # Assuming column A is attributes and column B is values
                sub_element = etree.SubElement(journal_element, row[0])  # Attribute name
                # Create empty tags if value is NaN
                sub_element.text = str(row[1]) if pd.notna(row[1]) else " "  # Ensure tags are empty

            # Create article element
            article_element = etree.SubElement(root, "article_info")
            for _, row in article_df.iterrows():
                # Assuming column A is attributes and column B is values
                sub_element = etree.SubElement(article_element, row[0])  # Attribute name
                # Create empty tags if value is NaN
                sub_element.text = str(row[1]) if pd.notna(row[1]) else " "  # Ensure tags are empty

            # Create authors list element
            authors_element = etree.SubElement(root, "author_list")

            # Iterate through the authors
            for col in author_df.columns[1:]:  # Skip the first column (attributes)
                author_element = etree.SubElement(authors_element, "author")
                for index, row in author_df.iterrows():
                    # Create a sub-element for each attribute in the first column
                    attribute = author_df.iloc[index, 0]  # Get the attribute name from the first column
                    sub_element = etree.SubElement(author_element, attribute)
                    # Create empty tags if value is NaN
                    sub_element.text = row[col] if pd.notna(row[col]) else " "  # Ensure tags are empty

            # Write to XML file with pretty print
            xml_str = etree.tostring(root, pretty_print=True, xml_declaration=True, encoding='UTF-8')
            with open(save_path, 'wb') as file:
                file.write(xml_str)

            update_status(f"Success: XML file generated and saved to: {save_path}")
        except Exception as e:
            update_status(f"Error: Failed to generate XML: {e}")

# Setup the main window
root = tk.Tk()
root.title("Excel to XML Converter")
root.geometry("600x300")  # Set the initial size to 600x400

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