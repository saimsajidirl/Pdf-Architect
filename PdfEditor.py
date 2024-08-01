import tkinter as tk
from tkinter import filedialog, ttk
from PyPDF2 import PdfReader, PdfWriter
import docx

class PDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Tool")

        self.file_list = []
        self.current_page = 1
        self.num_pages = 0
        self.text = ""

        self.create_widgets()

    def create_widgets(self):
        # Main Frame
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Toolbar
        self.toolbar = ttk.Frame(self.main_frame)
        self.toolbar.pack(side=tk.TOP, fill=tk.X)

        # Merge Button
        self.merge_button = ttk.Button(self.toolbar, text="Merge PDFs", command=self.merge_pdfs)
        self.merge_button.pack(side=tk.LEFT, padx=10, pady=5)

        # Edit Button
        self.edit_button = ttk.Button(self.toolbar, text="Edit PDF", command=self.edit_pdf)
        self.edit_button.pack(side=tk.LEFT, padx=10, pady=5)

        # Convert Button
        self.convert_button = ttk.Button(self.toolbar, text="Convert to Word", command=self.convert_to_word)
        self.convert_button.pack(side=tk.LEFT, padx=10, pady=5)

        # File Listbox
        self.listbox = tk.Listbox(self.main_frame, selectmode=tk.MULTIPLE)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Add Button
        self.add_button = ttk.Button(self.main_frame, text="Add PDFs", command=self.add_pdfs)
        self.add_button.pack(side=tk.LEFT, padx=10, pady=5)

        # Output Directory Label and Entry
        self.output_label = ttk.Label(self.main_frame, text="Save as:")
        self.output_label.pack(side=tk.LEFT, padx=10, pady=5)

        self.output_entry = ttk.Entry(self.main_frame)
        self.output_entry.pack(side=tk.LEFT, padx=10, pady=5)

        # Browse Button
        self.browse_button = ttk.Button(self.main_frame, text="Browse", command=self.browse_directory)
        self.browse_button.pack(side=tk.LEFT, padx=10, pady=5)

        # Text Editor Frame (for editing PDFs)
        self.text_frame = ttk.Frame(self.main_frame)
        self.text_frame.pack(fill=tk.BOTH, expand=True)

        # Text Editor
        self.text_editor = tk.Text(self.text_frame)
        self.text_editor.pack(fill=tk.BOTH, expand=True)

        # Navigation Toolbar (for editing PDFs)
        self.nav_toolbar = ttk.Frame(self.text_frame)
        self.nav_toolbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Previous Page Button
        self.prev_page_button = ttk.Button(self.nav_toolbar, text="Previous Page", command=self.prev_page)
        self.prev_page_button.pack(side=tk.LEFT, padx=10, pady=5)

        # Page Number Label
        self.page_number_label = ttk.Label(self.nav_toolbar, text=f"Page {self.current_page}/{self.num_pages}")
        self.page_number_label.pack(side=tk.LEFT, padx=10, pady=5)

        # Next Page Button
        self.next_page_button = ttk.Button(self.nav_toolbar, text="Next Page", command=self.next_page)
        self.next_page_button.pack(side=tk.LEFT, padx=10, pady=5)

        # Save Button (for editing PDFs)
        self.save_button = ttk.Button(self.nav_toolbar, text="Save", command=self.save_pdf)
        self.save_button.pack(side=tk.RIGHT, padx=10, pady=5)

    def merge_pdfs(self):
        if not self.file_list:
            tk.messagebox.showwarning("Warning", "No PDFs selected.")
            return

        merger = PdfMerger()
        for pdf_file in self.file_list:
            merger.append(pdf_file)

        output_file = self.output_entry.get()
        if not output_file:
            tk.messagebox.showwarning("Warning", "Please specify an output file path.")
            return

        merger.write(output_file)
        tk.messagebox.showinfo("Success", "PDFs merged successfully.")

        # Clear the file list and listbox
        self.file_list = []
        self.listbox.delete(0, tk.END)

    def edit_pdf(self):
        if not self.file_list:
            tk.messagebox.showwarning("Warning", "No PDFs selected.")
            return

        # Open the first PDF file for editing
        pdf_file = PdfReader(open(self.file_list[0], "rb"))
        self.num_pages = len(pdf_file.pages)

        # Extract and display the text from the current page
        self.text = pdf_file.pages[self.current_page - 1].extract_text()
        self.text_editor.delete("1.0", tk.END)
        self.text_editor.insert("1.0", self.text)

        # Update the page number label
        self.page_number_label.config(text=f"Page {self.current_page}/{self.num_pages}")

        # Show the text editor and navigation toolbar
        self.text_frame.pack(fill=tk.BOTH, expand=True)
        self.nav_toolbar.pack(side=tk.BOTTOM, fill=tk.X)

    def add_pdfs(self):
        files = filedialog.askopenfilenames(title="Select PDF files", filetypes=[("PDF files", "*.pdf")])
        if files:
            for file in files:
                self.listbox.insert(tk.END, file)
            self.file_list.extend(files)

    def browse_directory(self):
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, directory)

    def prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1

            # Extract and display the text from the previous page
            pdf_file = PdfReader(open(self.file_list[0], "rb"))
            self.text = pdf_file.pages[self.current_page - 1].extract_text()
            self.text_editor.delete("1.0", tk.END)
            self.text_editor.insert("1.0", self.text)

            # Update the page number label
            self.page_number_label.config(text=f"Page {self.current_page}/{self.num_pages}")

    def next_page(self):
        if self.current_page < self.num_pages:
            self.current_page += 1

            # Extract and display the text from the next page
            pdf_file = PdfReader(open(self.file_list[0], "rb"))
            self.text = pdf_file.pages[self.current_page - 1].extract_text()
            self.text_editor.delete("1.0", tk.END)
            self.text_editor.insert("1.0", self.text)

            # Update the page number label
            self.page_number_label.config(text=f"Page {self.current_page}/{self.num_pages}")

    def save_pdf(self):
        # Create a new PDF writer
        writer = PdfWriter()

        # Add the current page to the writer
        writer.addPage(PdfReader(open(self.file_list[0], "rb")).pages[self.current_page - 1])

        # Write the updated text to the PDF
        writer.write(open(self.file_list[0], "wb"))

        tk.messagebox.showinfo("Success", "PDF saved successfully.")

    def convert_to_word(self):
        if not self.file_list:
            tk.messagebox.showwarning("Warning", "No PDFs selected.")
            return

        file_path = self.file_list[0]
        output_file = filedialog.asksaveasfilename(defaultextension=".docx",
                                                 filetypes=[("Word Documents", "*.docx")])
        if not output_file:
            return

        try:
            pdf_reader = PdfReader(open(file_path, "rb"))
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()

            doc = docx.Document()
            doc.add_paragraph(text)
            doc.save(output_file)

            tk.messagebox.showinfo("Success", f"PDF converted to {output_file}")

        except Exception as e:
            tk.messagebox.showerror("Error", f"Conversion failed: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFApp(root)
    root.mainloop()