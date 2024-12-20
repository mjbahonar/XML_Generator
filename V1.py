import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
from xml.dom import minidom

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

class XMLGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XML File Generator")
        self.root.geometry("1500x800")  # Set the default window size here

        # Set the program icon using the resource_path method
        icon_path = self.resource_path("logo.ico")
        self.root.iconbitmap(icon_path)

        # Create Menu
        self.create_menu()

        # Notebook for Tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)

        # Create Tabs
        self.create_journal_tab()
        self.create_article_tab()
        self.create_authors_tab()

        # Generate Button
        self.generate_button = ttk.Button(root, text="Generate XML", command=self.generate_xml)
        self.generate_button.pack(pady=10)

    def resource_path(self, relative_path):
        """ Get the absolute path to the resource, works for dev and PyInstaller """
        try:
            # PyInstaller creates a temporary folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def create_menu(self):
        # Creating the menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New", command=self.new_file)
        file_menu.add_command(label="Save", command=self.generate_xml)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

        # Edit menu for Copy, Cut, Paste
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Select All", command=self.edit_select_all)
        edit_menu.add_command(label="Copy", command=self.edit_copy)
        edit_menu.add_command(label="Cut", command=self.edit_cut)
        edit_menu.add_command(label="Paste", command=self.edit_paste)
        file_menu.add_separator()
        file_menu.add_command(label="Clear All", command=self.clear_fields)

        # Default menu for applying predefined values (JIDS and JZMS)
        default_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Defaults", menu=default_menu)
        default_menu.add_command(label="JIDS", command=self.apply_default_1)
        default_menu.add_command(label="JZMS", command=self.apply_default_2)

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Contact Us", command=self.show_contact_us)



    def edit_copy(self):
        widget = self.root.focus_get()
        try:
            widget.event_generate("<<Copy>>")
        except:
            pass  # Handle if the widget doesn't support copy

    def edit_cut(self):
        widget = self.root.focus_get()
        try:
            widget.event_generate("<<Cut>>")
        except:
            pass  # Handle if the widget doesn't support cut

    def edit_paste(self):
        widget = self.root.focus_get()
        try:
            widget.event_generate("<<Paste>>")
        except:
            pass  # Handle if the widget doesn't support paste

    def edit_select_all(self):
        widget = self.root.focus_get()
        try:
            widget.event_generate("<<SelectAll>>")
        except:
            pass  # Handle if the widget doesn't support select all


    def new_file(self):
        self.clear_fields()
        messagebox.showinfo("New File", "All fields have been reset.")

    def show_about(self):
        # Editable help content
        messagebox.showinfo("About", "XML File Generator\nVersion 1.1\n\nThis tool helps generate XML files for journal articles. Use the tabs to input journal, article, and author details.")

    def show_contact_us(self):
        # Contact us information
        messagebox.showinfo("Contact Us", "Contact Us\n\nPlease contact me in Telegram with https://t.me/thegrupmy")

    def create_journal_tab(self):
        # Journal Information Tab
        self.journal_tab = ScrollableFrame(self.notebook)
        self.notebook.add(self.journal_tab, text="Journal Information")

        # Frame for Journal Information
        journal_frame = ttk.Frame(self.journal_tab.scrollable_frame, padding="10")
        journal_frame.pack(fill='both', expand=True)

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
            label = ttk.Label(journal_frame, text=label_text + ":", anchor='w')
            label.grid(row=i, column=0, sticky=tk.W, padx=(10, 5), pady=2)  # Consistent padding
            entry = ttk.Entry(journal_frame, width=50)
            entry.grid(row=i, column=1, padx=(5, 10), pady=2, sticky=tk.W)  # Align to start of column
            self.bind_copy_paste(entry)  # Enable copy-paste operations
            self.journal_entries[key] = entry

        # Publication Dates Section
        self.date_frame = ttk.LabelFrame(journal_frame, text="Publication Dates", padding="10")
        self.date_frame.grid(row=len(fields), column=0, columnspan=2, pady=10, padx=5, sticky='ew')

        self.pub_dates = []
        ttk.Button(self.date_frame, text="Add Date", command=self.add_date).grid(row=0, column=0, sticky=(tk.W, tk.E))

        # Default Buttons and Clear Button
        buttons_frame = ttk.Frame(journal_frame)
        buttons_frame.grid(row=len(fields) + 1, column=0, columnspan=2, pady=10)

        ttk.Button(buttons_frame, text="JIDS", command=self.apply_default_1).grid(row=0, column=0, padx=5)
        ttk.Button(buttons_frame, text="JZMS", command=self.apply_default_2).grid(row=0, column=1, padx=5)
        ttk.Button(buttons_frame, text="Clear", command=self.clear_journal_fields).grid(row=0, column=2, padx=5)

    def create_article_tab(self):
        # Article Information Tab
        self.article_tab = ScrollableFrame(self.notebook)
        self.notebook.add(self.article_tab, text="Article Information")

        # Frame for Article Information
        article_frame = ttk.Frame(self.article_tab.scrollable_frame, padding="10")
        article_frame.pack(fill='both', expand=True)

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
            label = ttk.Label(article_frame, text=label_text + ":", anchor='w')
            label.grid(row=i, column=0, sticky=tk.W, padx=(10, 5), pady=2)  # Consistent padding
            entry = ttk.Entry(article_frame, width=50)
            entry.grid(row=i, column=1, padx=(5, 10), pady=2, sticky=tk.W)  # Align to start of column
            self.bind_copy_paste(entry)  # Enable copy-paste operations
            self.article_entries[key] = entry

        # Abstract as Rich Text Box
        ttk.Label(article_frame, text="Abstract:", anchor='w').grid(row=len(fields), column=0, sticky=tk.W, padx=(10, 5), pady=2)
        self.abstract_text = tk.Text(article_frame, height=5, width=50, wrap="word")
        self.abstract_text.grid(row=len(fields), column=1, padx=(5, 10), pady=2, sticky=tk.W)
        self.bind_copy_paste(self.abstract_text)  # Enable copy-paste operations

        ttk.Label(article_frame, text="Abstract (FA):", anchor='w').grid(row=len(fields) + 1, column=0, sticky=tk.W, padx=(10, 5), pady=2)
        self.abstract_text_fa = tk.Text(article_frame, height=5, width=50, wrap="word")
        self.abstract_text_fa.grid(row=len(fields) + 1, column=1, padx=(5, 10), pady=2, sticky=tk.W)
        self.bind_copy_paste(self.abstract_text_fa)  # Enable copy-paste operations

        # Clear Button
        ttk.Button(article_frame, text="Clear", command=self.clear_article_fields).grid(row=len(fields) + 2, column=0, columnspan=2, pady=10)

    def create_authors_tab(self):
        # Authors Management Tab
        self.authors_tab = ScrollableFrame(self.notebook)
        self.notebook.add(self.authors_tab, text="Author(s)")

        self.authors_frame = ttk.LabelFrame(self.authors_tab.scrollable_frame, text="Authors", padding="10")
        self.authors_frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.authors = []
        ttk.Button(self.authors_frame, text="Add Author", command=self.add_author).pack(pady=5)
        ttk.Button(self.authors_frame, text="Clear", command=self.clear_authors).pack(pady=5)

    def add_date(self):
        date_frame = ttk.Frame(self.date_frame)
        date_frame.grid(pady=5, sticky='ew')

        ttk.Label(date_frame, text="Type:").grid(row=0, column=0, sticky=tk.W, padx=(10, 5))
        type_combobox = ttk.Combobox(date_frame, values=["jalali", "gregorian"], width=10)
        type_combobox.grid(row=0, column=1, padx=(5, 10))

        ttk.Label(date_frame, text="Year:").grid(row=0, column=2, sticky=tk.W, padx=(10, 5))
        year_entry = ttk.Entry(date_frame, width=5)
        year_entry.grid(row=0, column=3, padx=(5, 10))
        self.bind_copy_paste(year_entry)  # Enable copy-paste operations

        ttk.Label(date_frame, text="Month:").grid(row=0, column=4, sticky=tk.W, padx=(10, 5))
        month_entry = ttk.Entry(date_frame, width=5)
        month_entry.grid(row=0, column=5, padx=(5, 10))
        self.bind_copy_paste(month_entry)  # Enable copy-paste operations

        ttk.Label(date_frame, text="Day:").grid(row=0, column=6, sticky=tk.W, padx=(10, 5))
        day_entry = ttk.Entry(date_frame, width=5)
        day_entry.grid(row=0, column=7, padx=(5, 10))
        self.bind_copy_paste(day_entry)  # Enable copy-paste operations

        self.pub_dates.append((type_combobox, year_entry, month_entry, day_entry))

    def add_author(self):
        author_frame = ttk.Frame(self.authors_frame, padding="5")
        author_frame.pack(pady=5, fill='x')

        fields = [
            ("First Name", 15), ("Middle Name", 15), ("Last Name", 15), ("Suffix", 10),
            ("First Name (FA)", 15), ("Middle Name (FA)", 15), ("Last Name (FA)", 15), ("Suffix (FA)", 10),
            ("Email", 25), ("Code", 10), ("ORCID", 20), ("Core Author (Yes/No)", 10),
            ("Affiliation", 30), ("Affiliation (FA)", 30)
        ]

        author_entries = []
        for i, (label, width) in enumerate(fields):
            ttk.Label(author_frame, text=label + ":", anchor='w').grid(row=i//4, column=(i % 4) * 2, sticky=tk.W, padx=(10, 5))
            if label == "Core Author (Yes/No)":  # Core author as a boolean checkbox
                var = tk.BooleanVar()
                checkbox = ttk.Checkbutton(author_frame, variable=var)
                checkbox.grid(row=i//4, column=(i % 4) * 2 + 1, padx=(5, 10), sticky=tk.W)
                author_entries.append(var)
            else:
                entry = ttk.Entry(author_frame, width=width)
                entry.grid(row=i//4, column=(i % 4) * 2 + 1, padx=(5, 10), sticky=tk.W)
                self.bind_copy_paste(entry)  # Enable copy-paste operations
                author_entries.append(entry)

        remove_button = ttk.Button(author_frame, text="Remove", command=lambda: self.remove_author(author_frame))
        remove_button.grid(row=len(fields)//4 + 1, column=0, columnspan=8, pady=5)

        self.authors.append((author_frame, author_entries))

    def remove_author(self, frame):
        frame.destroy()
        self.authors = [author for author in self.authors if author[0] != frame]

    def clear_fields(self):
        self.clear_journal_fields()
        self.clear_article_fields()
        self.clear_authors()

    def clear_journal_fields(self):
        for entry in self.journal_entries.values():
            entry.delete(0, tk.END)

    def clear_article_fields(self):
        for entry in self.article_entries.values():
            entry.delete(0, tk.END)
        self.abstract_text.delete("1.0", tk.END)
        self.abstract_text_fa.delete("1.0", tk.END)

    def clear_authors(self):
        for author_frame, _ in list(self.authors):  # Convert to list to avoid modification during iteration
            self.remove_author(author_frame)

    def apply_default_1(self):
        # Predefined values for Default 1
        defaults = {
            "title": "Journal of Isfahan Dental School",
            "title_fa": "مجله دانکشده دندانپزشکی",
            "short_title": "JIDS",
            "subject": "Medical Sciences",
            "web_url": "http://jids.ir",
            "journal_hbi_system_id": "",
            "journal_hbi_system_user": "",
            "journal_id_issn": "1735-255X",
            "journal_id_issn_online": "",
            "journal_id_pii": "",
            "journal_id_doi": "",
            "journal_id_iranmedex": "",
            "journal_id_magiran": "",
            "journal_id_sid": "",
            "journal_id_nlai": "",
            "journal_id_science": "",
            "language": "fa",
            "volume": "",
            "number": ""
        }
        for key, value in defaults.items():
            self.journal_entries[key].delete(0, tk.END)
            self.journal_entries[key].insert(0, value)

    def apply_default_2(self):
        # Predefined values for Default 2
        defaults = {
            "title": "Journal of Zabol Medical school",
            "title_fa": "مجمله دانشگاه زابل",
            "short_title": "ZJMS",
            "subject": "Engineering",
            "web_url": "http://jzms.ir",
            "journal_hbi_system_id": "",
            "journal_hbi_system_user": "",
            "journal_id_issn": "2645-880X",
            "journal_id_issn_online": "2645-7180",
            "journal_id_pii": "",
            "journal_id_doi": "",
            "journal_id_iranmedex": "",
            "journal_id_magiran": "",
            "journal_id_sid": "",
            "journal_id_nlai": "",
            "journal_id_science": "",
            "language": "en",
            "volume": "20",
            "number": "2"
        }
        for key, value in defaults.items():
            self.journal_entries[key].delete(0, tk.END)
            self.journal_entries[key].insert(0, value)

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
                if field == "coreauthor":
                    elem = ET.SubElement(author, field)
                    elem.text = "Yes" if entry.get() else "No"  # Set as Yes/No based on boolean checkbox
                else:
                    elem = ET.SubElement(author, field)
                    elem.text = entry.get()

        # Format XML to be readable
        xmlstr = minidom.parseString(ET.tostring(root)).toprettyxml(indent="   ")

        # Get default name for save dialog
        default_name = self.article_entries.get("article_title").get() or "Untitled_Article"

        # Save File Dialog
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xml",
            filetypes=[("XML files", "*.xml")],
            initialfile=default_name
        )
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(xmlstr)
            messagebox.showinfo("Success", "XML file generated successfully!")

    def bind_copy_paste(self, widget):
        """Enable copy, cut, paste, and select all for entry and text widgets with custom paste handling"""
        widget.bind("<Control-c>", lambda e: widget.event_generate("<<Copy>>"))
        widget.bind("<Control-x>", lambda e: widget.event_generate("<<Cut>>"))
        
        # Custom paste handler to prevent pasting twice
        def custom_paste(event):
            try:
                # Get clipboard content
                clipboard = widget.clipboard_get()
                # Insert clipboard content at the current position
                widget.insert(tk.INSERT, clipboard)
            except tk.TclError:
                pass  # Handle cases where clipboard is empty or unavailable

            return "break"  # Prevent the default paste behavior

        widget.bind("<Control-v>", custom_paste)
        widget.bind("<Control-a>", lambda e: widget.event_generate("<<SelectAll>>"))

def main():
    root = tk.Tk()
    app = XMLGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
