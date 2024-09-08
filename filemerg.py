import tkinter as tk
from tkinter import filedialog, messagebox
import PyPDF2
from PIL import Image
import os
import win32com.client  # Ensure you have pywin32 installed for this


class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Document Merger with Reorder Option")
        self.root.geometry("600x700")

        self.file_paths = []  # Stores paths of selected files

        # Initialize all instance attributes within __init__
        self.label_project = tk.Label(self.root, text="Enter Project Number:")
        self.entry_project = tk.Entry(self.root, width=50)
        self.label_activity = tk.Label(self.root, text="Enter Activity Number:")
        self.entry_activity = tk.Entry(self.root, width=50)
        self.label_mt = tk.Label(self.root, text="Enter MT Number:")
        self.entry_mt = tk.Entry(self.root, width=50)
        self.label_lead_trade = tk.Label(self.root, text="Enter Lead Trade:")
        self.entry_lead_trade = tk.Entry(self.root, width=50)
        self.label_files = tk.Label(self.root, text="Select files (PDF, Word, Excel, Images):")
        self.button_browse_files = tk.Button(self.root, text="Browse", command=self.browse_files)
        self.files_frame = tk.Frame(self.root)
        self.buttons_frame = tk.Frame(self.root)
        self.button_merge = tk.Button(self.buttons_frame, text="Merge Files", command=self.merge_files)
        self.button_clear = tk.Button(self.buttons_frame, text="Clear", command=self.clear_inputs)
        self.button_close = tk.Button(self.buttons_frame, text="Close", command=self.root.quit)

        self.create_widgets()

    def create_widgets(self):
        # Layout for input fields
        self.label_project.pack(pady=5)
        self.entry_project.pack(pady=5)
        self.label_activity.pack(pady=5)
        self.entry_activity.pack(pady=5)
        self.label_mt.pack(pady=5)
        self.entry_mt.pack(pady=5)
        self.label_lead_trade.pack(pady=5)
        self.entry_lead_trade.pack(pady=5)

        # File selection
        self.label_files.pack(pady=5)
        self.button_browse_files.pack(pady=5)
        self.files_frame.pack(pady=5, fill='both', expand=True)

        # Buttons frame
        self.buttons_frame.pack(pady=20)
        self.button_merge.grid(row=0, column=0, padx=5)
        self.button_clear.grid(row=0, column=1, padx=5)
        self.button_close.grid(row=0, column=2, padx=5)

    def browse_files(self):
        file_paths = filedialog.askopenfilenames(
            filetypes=[("All files", "*.pdf;*.docx;*.xlsx;*.png;*.jpg")]
        )
        if file_paths:
            for file_path in file_paths:
                if file_path not in self.file_paths:  # Avoid duplicate files
                    self.file_paths.append(file_path)
                    self.display_file(file_path)

    def display_file(self, file_path):
        file_frame = tk.Frame(self.files_frame)
        file_frame.pack(fill='x', pady=2)

        file_label = tk.Label(file_frame, text=os.path.basename(file_path), width=50, anchor='w')
        file_label.pack(side='left', padx=10)

        # Move Up button
        move_up_button = tk.Button(
            file_frame, text="Up", command=lambda: self.move_file(file_path, -1), width=5
        )
        move_up_button.pack(side='left', padx=5)

        # Move Down button
        move_down_button = tk.Button(
            file_frame, text="Down", command=lambda: self.move_file(file_path, 1), width=5
        )
        move_down_button.pack(side='left', padx=5)

        # Delete button
        delete_button = tk.Button(
            file_frame, text="Delete", command=lambda: self.delete_file(file_frame, file_path), width=10
        )
        delete_button.pack(side='right', padx=10)

    def move_file(self, file_path, direction):
        try:
            index = self.file_paths.index(file_path)
            new_index = index + direction
            if 0 <= new_index < len(self.file_paths):
                self.file_paths[index], self.file_paths[new_index] = (
                    self.file_paths[new_index], self.file_paths[index]
                )
                self.update_file_display()
        except Exception as e:
            messagebox.showerror("Reorder Error", f"An error occurred while reordering files: {e}")

    def update_file_display(self):
        # Clear the current file display
        for widget in self.files_frame.winfo_children():
            widget.destroy()

        # Recreate the file display with updated order
        for file_path in self.file_paths:
            self.display_file(file_path)

    def delete_file(self, file_frame, file_path):
        """Remove the file from the list and the UI"""
        self.file_paths.remove(file_path)
        file_frame.destroy()  # Remove the file's frame from the UI

    def merge_files(self):
        if not self.file_paths:
            messagebox.showerror("Error", "Please select at least one file.")
            return

        project = self.entry_project.get().strip()
        activity = self.entry_activity.get().strip()
        mt = self.entry_mt.get().strip()
        lead_trade = self.entry_lead_trade.get().strip()

        if not project or not activity or not mt or not lead_trade:
            messagebox.showerror("Error", "Please fill in all the fields.")
            return

        output_filename = f"{project}-{activity}-{mt}-{lead_trade}.pdf"
        output_path = filedialog.asksaveasfilename(
            initialfile=output_filename, defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not output_path:
            return

        try:
            pdf_writer = PyPDF2.PdfWriter()

            for file_path in self.file_paths:
                ext = os.path.splitext(file_path)[1].lower()

                if ext == ".pdf":
                    self.merge_pdf(file_path, pdf_writer)
                elif ext == ".docx":
                    self.convert_word_to_pdf(file_path, pdf_writer)
                elif ext == ".xlsx":
                    self.convert_excel_to_pdf(file_path, pdf_writer)
                elif ext in [".png", ".jpg", ".jpeg"]:
                    self.convert_image_to_pdf(file_path, pdf_writer)

            with open(output_path, "wb") as output_file:
                pdf_writer.write(output_file)

            messagebox.showinfo("Success", "Files have been merged successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    @staticmethod
    def merge_pdf(file_path, pdf_writer):
        try:
            pdf_reader = PyPDF2.PdfReader(file_path)
            for page in range(len(pdf_reader.pages)):
                pdf_writer.add_page(pdf_reader.pages[page])
        except Exception as e:
            messagebox.showerror("PDF Error", f"Error reading {file_path}: {e}")

    @staticmethod
    def convert_word_to_pdf(file_path, pdf_writer):
        try:
            word = win32com.client.Dispatch("Word.Application")
            doc = word.Documents.Open(file_path)
            temp_pdf = file_path.replace(".docx", ".pdf")
            doc.SaveAs(temp_pdf, FileFormat=17)  # 17 = PDF format
            doc.Close(False)
            word.Quit()
            PDFMergerApp.merge_pdf(temp_pdf, pdf_writer)
            os.remove(temp_pdf)
        except Exception as e:
            messagebox.showerror("Word Error", f"Error converting Word document {file_path}: {e}")

    @staticmethod
    def convert_excel_to_pdf(file_path, pdf_writer):
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(file_path)
            temp_pdf = file_path.replace(".xlsx", ".pdf")
            wb.ExportAsFixedFormat(0, temp_pdf)  # 0 = PDF format
            wb.Close(False)
            excel.Quit()
            PDFMergerApp.merge_pdf(temp_pdf, pdf_writer)
            os.remove(temp_pdf)
        except Exception as e:
            messagebox.showerror("Excel Error", f"Error converting Excel document {file_path}: {e}")

    @staticmethod
    def convert_image_to_pdf(file_path, pdf_writer):
        try:
            img = Image.open(file_path)
            temp_pdf = file_path.replace(os.path.splitext(file_path)[1], ".pdf")
            img.convert("RGB").save(temp_pdf)
            PDFMergerApp.merge_pdf(temp_pdf, pdf_writer)
            os.remove(temp_pdf)
        except Exception as e:
            messagebox.showerror("Image Error", f"Error converting image {file_path}: {e}")

    def clear_inputs(self):
        self.entry_project.delete(0, tk.END)
        self.entry_activity.delete(0, tk.END)
        self.entry_mt.delete(0, tk.END)
        self.entry_lead_trade.delete(0, tk.END)
        for widget in self.files_frame.winfo_children():
            widget.destroy()  # Remove all file displays from the frame
        self.file_paths = []  # Clear the file paths


if __name__ == "__main__":
    app_root = tk.Tk()  # Renamed to avoid shadowing
    app = PDFMergerApp(app_root)
    app_root.mainloop()
