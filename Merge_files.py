from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from PIL import Image
import os
import comtypes
import comtypes.client
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
import time
import threading

# Initialize COM for Word with detailed error handling
com_available = False
try:
    comtypes.CoInitialize()
    word_test = comtypes.client.CreateObject('Word.Application', dynamic=True)
    # Try to load type library for constants
    try:
        comtypes.client.GetModule(['{000209FF-0000-0000-C000-000000000046}'])
        from comtypes.gen import Word
        wdFormatPDF = Word.wdFormatPDF
    except Exception:
        wdFormatPDF = 17  # Fallback to hardcoded constant
    word_test.Quit()
    com_available = True
    print("COM module for Word initialized successfully")
except Exception as e:
    messagebox.showwarning("Warning", f"Microsoft Word COM initialization failed: {e}. DOCX conversion will be disabled. Ensure Word is installed, run as administrator, and check 32/64-bit compatibility.")
    print(f"Failed to initialize COM module for Word: {e}")
    wdFormatPDF = 17

# Function to convert an image (JPG or PNG) to PDF
def image_to_pdf(image_path, output_pdf):
    start_time = time.time()
    try:
        img = Image.open(image_path)
        img = img.convert('RGB')
        img_width, img_height = img.size
        page_width, page_height = letter
        scale = min(page_width / img_width, page_height / img_height)
        new_width, new_height = img_width * scale, img_height * scale
        x_offset, y_offset = (page_width - new_width) / 2, (page_height - new_height) / 2
        c = canvas.Canvas(output_pdf, pagesize=letter)
        c.drawImage(ImageReader(image_path), x_offset, y_offset, width=new_width, height=new_height)
        c.save()
        print(f"Converted {image_path} in {time.time() - start_time:.2f}s")
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert {image_path}: {e}")
        print(f"Error converting {image_path}: {e} in {time.time() - start_time:.2f}s")
        return False

# Function to convert a DOCX file to PDF using comtypes
def docx_to_pdf(docx_path, output_pdf):
    if not com_available:
        messagebox.showerror("Error", "DOCX conversion requires Microsoft Word, which is not available.")
        print(f"DOCX conversion disabled: Microsoft Word not available")
        return False
    start_time = time.time()
    word = None
    doc = None
    try:
        word = comtypes.client.CreateObject('Word.Application', dynamic=True)
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(os.path.abspath(docx_path))
        abs_output = os.path.abspath(output_pdf)
        # Try multiple save methods to handle version differences
        try:
            doc.SaveAs2(abs_output, FileFormat=wdFormatPDF)
            print("Used SaveAs2 with FileFormat")
        except (AttributeError, TypeError):
            try:
                doc.SaveAs(abs_output, FileFormat=wdFormatPDF)
                print("Used SaveAs with FileFormat")
            except (AttributeError, TypeError):
                try:
                    doc.SaveAs(abs_output, wdFormatPDF)  # Positional argument
                    print("Used SaveAs with positional FileFormat")
                except (AttributeError, TypeError):
                    try:
                        doc.ExportAsFixedFormat(abs_output, OutputFormat=wdFormatPDF)
                        print("Used ExportAsFixedFormat")
                    except (AttributeError, TypeError):
                        doc.SaveAs(abs_output)  # Try without format parameter
                        print("Used SaveAs without format parameter")
        print(f"Converted {docx_path} to {output_pdf} in {time.time() - start_time:.2f}s")
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert {docx_path}: {e}")
        print(f"Error converting {docx_path}: {e} in {time.time() - start_time:.2f}s")
        return False
    finally:
        try:
            if doc:
                doc.Close()
        except:
            pass
        try:
            if word:
                word.Quit()
        except:
            pass

# Function to convert a TXT file to PDF
def txt_to_pdf(txt_path, output_pdf):
    start_time = time.time()
    try:
        c = canvas.Canvas(output_pdf, pagesize=letter)
        c.setFont("Helvetica", 12)
        with open(txt_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        y = 750
        for line in lines:
            if y < 50:
                c.showPage()
                y = 750
            c.drawString(50, y, line.strip())
            y -= 15
        c.save()
        print(f"Converted {txt_path} to {output_pdf} in {time.time() - start_time:.2f}s")
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert {txt_path}: {e}")
        print(f"Error converting {txt_path}: {e} in {time.time() - start_time:.2f}s")
        return False

# Function to extract specific pages from a PDF or DOCX
def extract_pages(input_file, output_file, pages):
    start_time = time.time()
    temp_pdf = None
    try:
        if input_file.endswith('.docx'):
            if not com_available:
                messagebox.showerror("Error", "DOCX page extraction requires Microsoft Word, which is not available.")
                print(f"DOCX page extraction disabled: Microsoft Word not available")
                return None
            temp_pdf = os.path.join("temp_split", f"temp_{os.path.basename(input_file)}.pdf")
            if not docx_to_pdf(input_file, temp_pdf):
                return None
            input_file = temp_pdf

        reader = PdfReader(input_file)
        total_pages = len(reader.pages)
        valid_pages = [p for p in pages if 0 <= p < total_pages]
        if not valid_pages:
            messagebox.showerror("Error", f"No valid pages selected for {input_file}. Using all pages.")
            valid_pages = list(range(total_pages))  # Use all pages if invalid

        writer = PdfWriter()
        for page_num in valid_pages:
            writer.add_page(reader.pages[page_num])
        with open(output_file, 'wb') as f:
            writer.write(f)
        print(f"Extracted pages {valid_pages} from {input_file} to {output_file} in {time.time() - start_time:.2f}s")
        return output_file, temp_pdf
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract pages from {input_file}: {e}")
        print(f"Error extracting pages from {input_file}: {e} in {time.time() - start_time:.2f}s")
        return None
    finally:
        if temp_pdf and os.path.exists(temp_pdf):
            for _ in range(5):
                try:
                    os.remove(temp_pdf)
                    print(f"Deleted temp file: {temp_pdf}")
                    break
                except PermissionError:
                    time.sleep(0.1)
                except Exception as e:
                    print(f"Error deleting {temp_pdf}: {e}")

# Function to merge files into a single PDF
def merge_files(input_files, output_file, progress_callback=None, file_pages=None):
    start_time = time.time()
    merger = PdfMerger()
    temp_files = []
    output_dir = "temp_split"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    print(f"merge_files: input_files={[os.path.abspath(f) for f in input_files]}")
    print(f"merge_files: file_pages={file_pages}")

    total_files = len(input_files)
    for i, file in enumerate(input_files):
        if not os.path.exists(file):
            messagebox.showerror("Error", f"File {file} does not exist")
            continue
        file = os.path.abspath(file)
        pages = file_pages.get(file, []) if file_pages else []
        print(f"Processing {file}: requested pages={pages}")
        if progress_callback:
            progress_callback(i / total_files * 100)

        file_start_time = time.time()
        if file.endswith(('.pdf', '.docx')) and pages:
            temp_file, extra_temp = extract_pages(file, os.path.join(output_dir, f"temp_extract_{i}.pdf"), pages)
            if temp_file:
                temp_files.append(temp_file)
                if extra_temp:
                    temp_files.append(extra_temp)
                try:
                    merger.append(temp_file)
                    print(f"Appended {temp_file} with selected pages in {time.time() - file_start_time:.2f}s")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to append {temp_file}: {e}")
                    print(f"Error appending {temp_file}: {e}")
            else:
                continue
        elif file.endswith('.pdf'):
            try:
                merger.append(file)
                print(f"Appended {file} with all pages in {time.time() - file_start_time:.2f}s")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to append {file}: {e}")
                print(f"Error appending {file}: {e}")
        elif file.endswith('.docx'):
            if not com_available:
                messagebox.showerror("Error", f"Skipping {file}: DOCX conversion requires Microsoft Word")
                print(f"Skipped {file}: DOCX conversion disabled")
                continue
            temp_pdf = os.path.join(output_dir, f"temp_docx_{i}.pdf")
            if docx_to_pdf(file, temp_pdf):
                temp_files.append(temp_pdf)
                try:
                    merger.append(temp_pdf)
                    print(f"Appended {temp_pdf} with all pages in {time.time() - file_start_time:.2f}s")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to append {temp_pdf}: {e}")
                    print(f"Error appending {temp_pdf}: {e}")
        elif file.endswith(('.jpg', '.png')):
            temp_pdf = os.path.join(output_dir, f"temp_image_{i}.pdf")
            if image_to_pdf(file, temp_pdf):
                temp_files.append(temp_pdf)
                try:
                    merger.append(temp_pdf)
                    print(f"Appended {temp_pdf} in {time.time() - file_start_time:.2f}s")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to append {temp_pdf}: {e}")
                    print(f"Error appending {temp_pdf}: {e}")
        elif file.endswith('.txt'):
            temp_pdf = os.path.join(output_dir, f"temp_txt_{i}.pdf")
            if txt_to_pdf(file, temp_pdf):
                temp_files.append(temp_pdf)
                try:
                    merger.append(temp_pdf)
                    print(f"Appended {temp_pdf} in {time.time() - file_start_time:.2f}s")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to append {temp_pdf}: {e}")
                    print(f"Error appending {temp_pdf}: {e}")
        else:
            messagebox.showerror("Error", f"Skipping {file}: Unsupported file type")
            print(f"Skipped {file}: Unsupported file type")

        if progress_callback:
            progress_callback((i + 1) / total_files * 100)

    if not merger.pages:
        messagebox.showerror("Error", "No pages were merged. Check your input files.")
        print("No pages merged")
        return

    try:
        with open(output_file, 'wb') as f:
            merger.write(f)
        messagebox.showinfo("Success", f"Files merged into {output_file}")
        print(f"Successfully merged into {output_file} in {time.time() - start_time:.2f}s")
    except PermissionError:
        messagebox.showerror("Error", f"Cannot write to {output_file}. Ensure it’s not open and you have permissions.")
        print(f"Permission error writing to {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Error merging files: {e}")
        print(f"Error merging files: {e}")
    finally:
        merger.close()

    for temp_file in temp_files:
        for _ in range(5):
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                    print(f"Deleted temp file: {temp_file}")
                break
            except PermissionError:
                time.sleep(0.1)
            except Exception as e:
                messagebox.showerror("Error", f"Error deleting {temp_file}: {e}")
                print(f"Error deleting {temp_file}: {e}")

# Dialog for setting page ranges per file
class PageSelectionDialog(tk.Toplevel):
    def __init__(self, parent, files):
        start_time = time.time()
        super().__init__(parent)
        self.title("Set Page Ranges")
        self.geometry("500x400")
        self.parent = parent
        self.file_pages = {os.path.abspath(file): [] for file in files}
        self.entries = {}
        print("Opening Set Pages dialog with files:", [os.path.abspath(f) for f in files])
        self.transient(parent)
        self.grab_set()

        try:
            self.main_frame = ttk.Frame(self)
            self.main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
            self.grid_rowconfigure(0, weight=1)
            self.grid_columnconfigure(0, weight=1)

            ttk.Label(self.main_frame, text="Enter page ranges (e.g., 1,3-5). Leave blank for all pages.").grid(row=0, column=0, columnspan=2, pady=5, sticky="w")

            canvas = tk.Canvas(self.main_frame)
            scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            canvas.grid(row=1, column=0, sticky="nsew")
            scrollbar.grid(row=1, column=1, sticky="ns")
            self.main_frame.grid_rowconfigure(1, weight=1)
            self.main_frame.grid_columnconfigure(0, weight=1)

            for i, file in enumerate(files):
                abs_file = os.path.abspath(file)
                if abs_file.endswith('.pdf') or (abs_file.endswith('.docx') and com_available):
                    frame = ttk.Frame(scrollable_frame)
                    frame.grid(row=i, column=0, sticky="ew", padx=5, pady=2)
                    label_text = os.path.basename(abs_file)
                    ttk.Label(frame, text=label_text, width=40).grid(row=0, column=0, sticky="w")
                    entry = ttk.Entry(frame)
                    entry.grid(row=0, column=1, sticky="ew")
                    frame.grid_columnconfigure(1, weight=1)
                    self.entries[abs_file] = entry
                    print(f"Added {label_text} to dialog")

            button_frame = ttk.Frame(self.main_frame)
            button_frame.grid(row=2, column=0, columnspan=2, pady=10)
            ttk.Button(button_frame, text="OK", command=self.save_pages).grid(row=0, column=0, padx=5)
            ttk.Button(button_frame, text="Cancel", command=self.destroy).grid(row=0, column=1, padx=5)
            print(f"Dialog initialized in {time.time() - start_time:.2f}s")
        except Exception as e:
            print(f"Dialog initialization error: {e}")
            messagebox.showerror("Error", f"Failed to initialize dialog: {e}")

    def save_pages(self):
        print("save_pages called")
        for file, entry in self.entries.items():
            print(f"Entry for {file}: {entry.get()}")
            pages = []
            try:
                if entry.get().strip():
                    for part in entry.get().split(','):
                        part = part.strip()
                        if '-' in part:
                            start, end = map(int, part.split('-'))
                            if start < 1 or end < start:
                                raise ValueError(f"Invalid page range for {os.path.basename(file)}")
                            pages.extend(range(start-1, end))
                        else:
                            page = int(part)
                            if page < 1:
                                raise ValueError(f"Invalid page number for {os.path.basename(file)}")
                            pages.append(page-1)
                self.file_pages[file] = pages
                print(f"Set pages for {file}: {pages}")
            except Exception as e:
                print(f"Error processing {file}: {e}")
                messagebox.showerror("Error", f"Invalid page format for {os.path.basename(file)}: {e}")
                return
        self.parent.file_pages = dict(self.file_pages)
        print(f"save_pages: file_pages={self.parent.file_pages}")
        self.destroy()

# GUI App with drag-and-drop and file selection
class FileMergerApp(TkinterDnD.Tk):
    def __init__(self):
        start_time = time.time()
        super().__init__()
        self.title("FileMerger - Offline File Combiner")
        self.geometry("600x400")
        self.files = []
        self.file_pages = {}
        style = ttk.Style()
        style.configure("TButton", padding=5)
        label_text = "Drag files here or click 'Add Files' to merge PDFs, DOCX, TXT, JPG, PNG"
        if not com_available:
            label_text += " (DOCX conversion disabled due to COM initialization failure)"
        ttk.Label(self, text=label_text).pack(pady=10)
        self.file_list = tk.Listbox(self, width=80, height=10)
        self.file_list.pack(pady=10)
        self.file_list.drop_target_register(DND_FILES)
        self.file_list.dnd_bind('<<Drop>>', self.drop_files)
        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Add Files", command=self.add_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Remove Selected", command=self.remove_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Move Up", command=self.move_up).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Move Down", command=self.move_down).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Set Pages", command=self.set_pages).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Merge Files", command=self.merge).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Help", command=self.show_help).pack(side=tk.LEFT, padx=5)
        self.progress = ttk.Progressbar(self, length=400, mode='determinate')
        self.progress.pack(pady=10)
        print(f"GUI initialized in {time.time() - start_time:.2f}s")

    def drop_files(self, event):
        start_time = time.time()
        files = self.split_files(event.data)
        for file in files:
            abs_file = os.path.abspath(file)
            if abs_file not in self.files and abs_file.endswith(('.pdf', '.docx', '.txt', '.jpg', '.png')):
                if abs_file.endswith('.docx') and not com_available:
                    messagebox.showwarning("Warning", f"Skipping {os.path.basename(abs_file)}: DOCX conversion disabled")
                    continue
                self.files.append(abs_file)
                self.file_list.insert(tk.END, os.path.basename(abs_file))
        print(f"Files after drop: {self.files} in {time.time() - start_time:.2f}s")

    def split_files(self, data):
        files = []
        if '{' in data:
            current = ""
            in_braces = False
            for char in data:
                if char == '{':
                    in_braces = True
                elif char == '}':
                    in_braces = False
                elif char == ' ' and not in_braces:
                    files.append(current)
                    current = ""
                else:
                    current += char
            if current:
                files.append(current)
        else:
            files = data.split()
        return [f.strip() for f in files if os.path.exists(f.strip())]

    def add_files(self):
        start_time = time.time()
        files = filedialog.askopenfilenames(filetypes=[("Supported Files", "*.pdf *.docx *.txt *.jpg *.png"), ("All Files", "*.*")])
        for file in files:
            abs_file = os.path.abspath(file)
            if abs_file not in self.files:
                if abs_file.endswith('.docx') and not com_available:
                    messagebox.showwarning("Warning", f"Skipping {os.path.basename(abs_file)}: DOCX conversion disabled")
                    continue
                self.files.append(abs_file)
                self.file_list.insert(tk.END, os.path.basename(abs_file))
        print(f"Files after add: {self.files} in {time.time() - start_time:.2f}s")

    def remove_file(self):
        start_time = time.time()
        selected = self.file_list.curselection()
        if selected:
            index = selected[0]
            self.file_list.delete(index)
            file = self.files.pop(index)
            self.file_pages.pop(file, None)
            print(f"Removed file: {file}, updated file_pages: {self.file_pages} in {time.time() - start_time:.2f}s")

    def move_up(self):
        start_time = time.time()
        selected = self.file_list.curselection()
        if selected and selected[0] > 0:
            index = selected[0]
            self.files.insert(index - 1, self.files.pop(index))
            self.file_list.delete(index)
            self.file_list.insert(index - 1, os.path.basename(self.files[index - 1]))
            self.file_list.select_set(index - 1)
            print(f"Files after move up: {self.files} in {time.time() - start_time:.2f}s")

    def move_down(self):
        start_time = time.time()
        selected = self.file_list.curselection()
        if selected and selected[0] < len(self.files) - 1:
            index = selected[0]
            self.files.insert(index + 1, self.files.pop(index))
            self.file_list.delete(index)
            self.file_list.insert(index + 1, os.path.basename(self.files[index]))
            self.file_list.select_set(index + 1)
        print(f"Files after move down: {self.files} in {time.time() - start_time:.2f}s")

    def set_pages(self):
        start_time = time.time()
        if not self.files:
            messagebox.showwarning("Warning", "No files selected")
            return
        print("Calling set_pages, files:", self.files)
        print("file_pages before dialog:", self.file_pages)
        PageSelectionDialog(self, self.files)
        print(f"Set pages dialog opened in {time.time() - start_time:.2f}s")

    def merge(self):
        start_time = time.time()
        if not self.files:
            messagebox.showwarning("Warning", "No files selected")
            return
        output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_file:
            print(f"merge: file_pages={self.file_pages}")
            self.progress['value'] = 0
            threading.Thread(target=merge_files, args=(self.files, output_file, self.update_progress, self.file_pages), daemon=True).start()
        print(f"Merge initiated in {time.time() - start_time:.2f}s")

    def update_progress(self, value):
        if int(value) % 10 == 0:  # Update every 10% to reduce GUI lag
            self.progress['value'] = value
            self.update_idletasks()

    def show_help(self):
        help_text = "1. Drag files or click 'Add Files' to add PDFs, DOCX, TXT, JPG, PNG.\n2. Reorder with 'Move Up'/'Move Down'.\n3. Click 'Set Pages' to select pages for each PDF/DOCX.\n4. Click 'Merge Files' to create a PDF."
        if not com_available:
            help_text += "\nNote: DOCX conversion is disabled due to Microsoft Word COM initialization failure. Ensure Word is installed and run as administrator."
        else:
            help_text += "\nNote: DOCX conversion requires Microsoft Word."
        messagebox.showinfo("Help", help_text)

if __name__ == "__main__":
    app = FileMergerApp()
    app.mainloop()
    
    # Clean up COM resources
    try:
        comtypes.CoUninitialize()
    except:
        pass
    