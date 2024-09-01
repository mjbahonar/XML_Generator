import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
from xml.dom import minidom

class XMLGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XML File Generator")

        # Notebook for Tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Create Tabs
        self.create_journal_tab()
        self.create_article_tab()
        self.create_authors_tab()

        # Generate Button
        self.generate_button = ttk.Button(root, text="Generate XML", command=self.generate_xml)
        self.generate_button.grid(row=1, column=0, pady=10)

    def create_journal_tab(self):
        # Journal Information Tab
        self.journal_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.journal_tab, text="Journal Information")

        fields = [
            ("Journal Title", "title"),
            ("Journal Title (FA)", "title_fa"),
            ("Short Title", "short_title"),
            ("Subject", "subject"),
            ("Web URL", "web_url"),
            ("Journal HBI System ID", "journal_hbi_system_id"),
            ("Journal HBI System User", "journal_hbi_system_user"),
            ("Journal ISSN", "journal_id_issn"),
            ("Journal ISSN Online", "journal_id_issn_online"),
            ("Journal ID PII", "journal_id_pii"),
            ("Journal DOI", "journal_id_doi"),
            ("Journal ID IranMedex", "journal_id_iranmedex"),
            ("Journal ID Magiran", "journal_id_magiran"),
            ("Journal ID SID", "journal_id_sid"),
            ("Journal ID NLAI", "journal_id_nlai"),
            ("Journal ID Science", "journal_id_science"),
            ("Language", "language"),
            ("Volume", "volume"),
            ("Number", "number")
        ]

        self.journal_entries = {}
        for i, (label_text, key) in enumerate(fields):
            ttk.Label(self.journal_tab, text=label_text + ":").grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
            entry = ttk.Entry(self.journal_tab, width=50)
            entry.grid(row=i, column=1, padx=5, pady=2)
            self.journal_entries[key] = entry

        # Publication Dates Section
        self.date_frame = ttk.LabelFrame(self.journal_tab, text="Publication Dates", padding="10")
        self.date_frame.grid(row=len(fields), column=0, columnspan=2, pady=10, padx=5)

        self.pub_dates = []
        ttk.Button(self.date_frame, text="Add Date", command=self.add_date).grid(row=0, column=0, sticky=(tk.W, tk.E))

    def create_article_tab(self):
        # Article Information Tab
        self.article_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.article_tab, text="Article Information")

        fields = [
            ("Article Title", "article_title"),
            ("Article Title (FA)", "article_title_fa"),
            ("Subject (FA)", "subject_fa"),
            ("Subject", "subject"),
            ("Content Type (FA)", "content_type_fa"),
            ("Content Type", "content_type"),
            ("Start Page", "start_page"),
            ("End Page", "end_page"),
            ("Web URL", "web_url"),
            ("Keywords", "keywords"),
            ("Keywords (FA)", "keywords_fa")
        ]

        self.article_entries = {}
        for i, (label_text, key) in enumerate(fields):
            ttk.Label(self.article_tab, text=label_text + ":").grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
            entry = ttk.Entry(self.article_tab, width=50)
            entry.grid(row=i, column=1, padx=5, pady=2)
            self.article_entries[key] = entry

        # Abstract as Rich Text Box
        ttk.Label(self.article_tab, text="Abstract:").grid(row=len(fields), column=0, sticky=tk.W, padx=5, pady=2)
        self.abstract_text = tk.Text(self.article_tab, height=5, width=50)
        self.abstract_text.grid(row=len(fields), column=1, padx=5, pady=2)

        ttk.Label(self.article_tab, text="Abstract (FA):").grid(row=len(fields) + 1, column=0, sticky=tk.W, padx=5, pady=2)
        self.abstract_text_fa = tk.Text(self.article_tab, height=5, width=50)
        self.abstract_text_fa.grid(row=len(fields) + 1, column=1, padx=5, pady=2)

    def create_authors_tab(self):
        # Authors Management Tab
        self.authors_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.authors_tab, text="Author(s)")

        self.authors_frame = ttk.LabelFrame(self.authors_tab, text="Authors", padding="10")
        self.authors_frame.grid(row=0, column=0, columnspan=2, pady=10, padx=5)

        self.authors = []
        ttk.Button(self.authors_frame, text="Add Author", command=self.add_author).grid(row=0, column=0, sticky=(tk.W, tk.E))

    def add_date(self):
        date_frame = ttk.Frame(self.date_frame)
        date_frame.grid(pady=5)

        ttk.Label(date_frame, text="Type:").grid(row=0, column=0)
        type_combobox = ttk.Combobox(date_frame, values=["jalali", "gregorian"], width=10)
        type_combobox.grid(row=0, column=1)

        ttk.Label(date_frame, text="Year:").grid(row=0, column=2)
        year_entry = ttk.Entry(date_frame, width=5)
        year_entry.grid(row=0, column=3)

        ttk.Label(date_frame, text="Month:").grid(row=0, column=4)
        month_entry = ttk.Entry(date_frame, width=5)
        month_entry.grid(row=0, column=5)

        ttk.Label(date_frame, text="Day:").grid(row=0, column=6)
        day_entry = ttk.Entry(date_frame, width=5)
        day_entry.grid(row=0, column=7)

        self.pub_dates.append((type_combobox, year_entry, month_entry, day_entry))

    def add_author(self):
        author_frame = ttk.Frame(self.authors_frame)
        author_frame.grid(pady=5)

        fields = [
            ("First Name", 15), ("Middle Name", 15), ("Last Name", 15), ("Suffix", 10),
            ("First Name (FA)", 15), ("Middle Name (FA)", 15), ("Last Name (FA)", 15), ("Suffix (FA)", 10),
            ("Email", 25), ("Code", 10), ("ORCID", 20), ("Core Author (Yes/No)", 10),
            ("Affiliation", 30), ("Affiliation (FA)", 30)
        ]

        author_entries = []
        for i, (label, width) in enumerate(fields):
            ttk.Label(author_frame, text=label + ":").grid(row=i//4, column=(i % 4) * 2, sticky=tk.W)
            entry = ttk.Entry(author_frame, width=width)
            entry.grid(row=i//4, column=(i % 4) * 2 + 1)
            author_entries.append(entry)

        remove_button = ttk.Button(author_frame, text="Remove", command=lambda: self.remove_author(author_frame))
        remove_button.grid(row=len(fields)//4 + 1, column=0, columnspan=8)

        self.authors.append((author_frame, author_entries))

    def remove_author(self, frame):
        frame.destroy()
        self.authors = [author for author in self.authors if author[0] != frame]

    def generate_xml(self):
        root = ET.Element("journal")

        # Add journal details
        for key, entry in self.journal_entries.items():
            element = ET.SubElement(root, key)
            element.text = entry.get()

        # Add publication dates
        for type_entry, year_entry, month_entry, day_entry in self.pub_dates:
            pubdate = ET.SubElement(root, "pubdate")
            type_elem = ET.SubElement(pubdate, "type")
            type_elem.text = type_entry.get()
            year_elem = ET.SubElement(pubdate, "year")
            year_elem.text = year_entry.get()
            month_elem = ET.SubElement(pubdate, "month")
            month_elem.text = month_entry.get()
            day_elem = ET.SubElement(pubdate, "day")
            day_elem.text = day_entry.get()

        # Add article details
        article = ET.SubElement(root, "article")
        for key, entry in self.article_entries.items():
            element = ET.SubElement(article, key)
            element.text = entry.get()

        # Add abstracts
        abstract = ET.SubElement(article, "abstract")
        abstract.text = self.abstract_text.get("1.0", tk.END).strip()

        abstract_fa = ET.SubElement(article, "abstract_fa")
        abstract_fa.text = self.abstract_text_fa.get("1.0", tk.END).strip()

        # Add authors
        author_list = ET.SubElement(root, "author_list")
        for author_frame, entries in self.authors:
            author = ET.SubElement(author_list, "author")
            fields = ["first_name", "middle_name", "last_name", "suffix",
                      "first_name_fa", "middle_name_fa", "last_name_fa", "suffix_fa",
                      "email", "code", "orcid", "coreauthor", "affiliation", "affiliation_fa"]

            for field, entry in zip(fields, entries):
                elem = ET.SubElement(author, field)
                elem.text = entry.get() if field != "coreauthor" else "Yes" if entry.get().lower() == "yes" else "No"

        # Format XML to be readable
        xmlstr = minidom.parseString(ET.tostring(root)).toprettyxml(indent="   ")

        # Save File Dialog
        file_path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml")])
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(xmlstr)
            messagebox.showinfo("Success", "XML file generated successfully!")

def main():
    root = tk.Tk()
    app = XMLGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
