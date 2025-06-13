import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfMerger, PdfReader
import os

class BookmarkSequencePreview(tk.Toplevel):
    def __init__(self, parent, sequence, missing_files, current_files):
        super().__init__(parent)
        self.title("Bookmark Sequence Preview")
        self.geometry("600x500")
        
        # Main frames
        top_frame = tk.Frame(self)
        top_frame.pack(fill="x", padx=10, pady=10)
        
        bottom_frame = tk.Frame(self)
        bottom_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Title and description
        tk.Label(top_frame, text="Detected Bookmark Sequence", font=('Arial', 12, 'bold')).pack(anchor="w")
        tk.Label(top_frame, text="Files will be merged in the order shown below:").pack(anchor="w", pady=(0, 10))
        
        # Sequence list
        seq_frame = tk.Frame(bottom_frame)
        seq_frame.pack(fill="both", expand=True)
        
        # Listbox with scrollbar
        scrollbar = tk.Scrollbar(seq_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.listbox = tk.Listbox(
            seq_frame, 
            yscrollcommand=scrollbar.set,
            selectmode=tk.SINGLE,
            font=('Courier', 10)
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)
        
        # Add items to listbox with color coding
        for i, (item, status) in enumerate(sequence, 1):
            prefix = "✓ " if status == "found" else "✗ "
            self.listbox.insert(tk.END, f"{i:2d}. {prefix}{os.path.basename(item)}")
            if status == "missing":
                self.listbox.itemconfig(i-1, {'fg': 'red'})
            elif status == "extra":
                self.listbox.itemconfig(i-1, {'fg': 'blue'})
        
        # Missing files section
        if missing_files:
            missing_frame = tk.LabelFrame(bottom_frame, text="Files Not Found in Current Selection", padx=5, pady=5)
            missing_frame.pack(fill="x", pady=(10, 0))
            
            for i, item in enumerate(missing_files[:5], 1):
                tk.Label(missing_frame, text=f"{i}. {item}", fg="red", anchor="w").pack(fill="x")
            if len(missing_files) > 5:
                tk.Label(missing_frame, text=f"...and {len(missing_files)-5} more", fg="red").pack()
        
        # Current files not in sequence
        current_basenames = [os.path.basename(f) for f in current_files]
        extra_files = [f for f in current_files 
                      if os.path.basename(f) not in [os.path.basename(s[0]) for s in sequence]]
        
        if extra_files:
            extra_frame = tk.LabelFrame(bottom_frame, text="Current Files Not in Sequence", padx=5, pady=5)
            extra_frame.pack(fill="x", pady=(10, 0))
            
            for i, item in enumerate(extra_files[:3], 1):
                tk.Label(extra_frame, text=f"{i}. {os.path.basename(item)}", fg="blue", anchor="w").pack(fill="x")
            if len(extra_files) > 3:
                tk.Label(extra_frame, text=f"...and {len(extra_files)-3} more", fg="blue").pack()
        
        # Buttons
        btn_frame = tk.Frame(bottom_frame)
        btn_frame.pack(pady=10)
        
        tk.Button(
            btn_frame, 
            text="Apply This Sequence", 
            command=lambda: self.apply_sequence(True),
            bg="#4CAF50", 
            fg="white"
        ).pack(side="left", padx=5)
        
        tk.Button(
            btn_frame, 
            text="Apply and Append Others", 
            command=lambda: self.apply_sequence(False),
            bg="#2196F3", 
            fg="white"
        ).pack(side="left", padx=5)
        
        tk.Button(
            btn_frame, 
            text="Cancel", 
            command=self.destroy
        ).pack(side="left", padx=5)
        
        self.sequence = [s[0] for s in sequence if s[1] == "found"]
        self.missing_files = missing_files
        self.extra_files = extra_files
        self.apply_callback = None
    
    def set_apply_callback(self, callback):
        self.apply_callback = callback
    
    def apply_sequence(self, strict_mode):
        if self.apply_callback:
            self.apply_callback(self.sequence, self.missing_files, self.extra_files, strict_mode)
        self.destroy()

class PDFMergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger Tool (Bookmark Sequence)")
        self.root.geometry("800x650")
        
        # Variables
        self.folder_path = tk.StringVar()
        self.output_filename = tk.StringVar(value="merged.pdf")
        self.include_subfolders = tk.BooleanVar()
        self.sort_by = tk.StringVar(value="name")
        self.pdf_files = []
        self.sequence_pdf = tk.StringVar()
        
        # Create UI
        self.create_widgets()
    
    def create_widgets(self):
        # Configure grid weights
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        
        # Folder selection
        tk.Label(self.root, text="Source Folder:").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))
        
        folder_frame = tk.Frame(self.root)
        folder_frame.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=(10, 0))
        
        tk.Entry(folder_frame, textvariable=self.folder_path).pack(side="left", fill="x", expand=True)
        tk.Button(folder_frame, text="Browse...", command=self.browse_folder).pack(side="left", padx=5)
        
        # Sequence PDF
        tk.Label(self.root, text="Get Sequence From PDF:").grid(row=1, column=0, sticky="w", padx=10, pady=(10, 0))
        
        seq_frame = tk.Frame(self.root)
        seq_frame.grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=(10, 0))
        
        tk.Entry(seq_frame, textvariable=self.sequence_pdf).pack(side="left", fill="x", expand=True)
        tk.Button(seq_frame, text="Browse...", command=self.browse_sequence_pdf).pack(side="left", padx=5)
        tk.Button(self.root, text="Extract Bookmark Sequence", command=self.preview_bookmark_sequence,
                bg="#FF9800", fg="white").grid(row=2, column=0, columnspan=2, pady=5)
        
        # Options frame
        options_frame = tk.LabelFrame(self.root, text="Options", padx=10, pady=10)
        options_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10, pady=10)
        
        # Include subfolders
        tk.Checkbutton(options_frame, text="Include subfolders", variable=self.include_subfolders,
                     command=self.update_file_list).grid(row=0, column=0, sticky="w")
        
        # Sort by
        tk.Label(options_frame, text="Sort by:").grid(row=0, column=1, sticky="w", padx=(20, 5))
        tk.Radiobutton(options_frame, text="Name", variable=self.sort_by, value="name",
                      command=self.update_file_list).grid(row=0, column=2, sticky="w")
        tk.Radiobutton(options_frame, text="Date", variable=self.sort_by, value="date",
                      command=self.update_file_list).grid(row=0, column=3, sticky="w")
        tk.Radiobutton(options_frame, text="Size", variable=self.sort_by, value="size",
                      command=self.update_file_list).grid(row=0, column=4, sticky="w")
        
        # Output filename
        tk.Label(self.root, text="Output Filename:").grid(row=4, column=0, sticky="w", padx=10, pady=(0, 5))
        tk.Entry(self.root, textvariable=self.output_filename).grid(row=4, column=1, sticky="ew", padx=(0, 10), pady=(0, 5))
        
        # File list with scrollbar
        list_frame = tk.Frame(self.root)
        list_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=10, pady=(0, 10))
        
        # Configure grid weights for list frame
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.listbox = tk.Listbox(
            list_frame, 
            yscrollcommand=scrollbar.set,
            selectmode=tk.EXTENDED,
            font=('Courier', 10)
        )
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.config(command=self.listbox.yview)
        
        # Control buttons
        control_frame = tk.Frame(self.root)
        control_frame.grid(row=6, column=0, columnspan=2, pady=(0, 10))
        
        tk.Button(control_frame, text="Move Up", command=self.move_up).pack(side="left", padx=2)
        tk.Button(control_frame, text="Move Down", command=self.move_down).pack(side="left", padx=2)
        tk.Button(control_frame, text="Reverse", command=self.reverse_order).pack(side="left", padx=2)
        tk.Button(control_frame, text="Reset Order", command=self.update_file_list).pack(side="left", padx=2)
        
        # Progress and status
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=100, mode="determinate")
        self.progress.grid(row=7, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 5))
        
        self.status_label = tk.Label(self.root, text="Ready", fg="gray")
        self.status_label.grid(row=8, column=0, columnspan=2, sticky="w", padx=10)
        
        # Merge button
        tk.Button(self.root, text="Merge PDFs", command=self.merge_pdfs, bg="#4CAF50", fg="white",
                font=('Arial', 10, 'bold')).grid(row=9, column=0, columnspan=2, pady=10)
    
    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.update_file_list()
    
    def browse_sequence_pdf(self):
        file_selected = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_selected:
            self.sequence_pdf.set(file_selected)
    
    def preview_bookmark_sequence(self):
        if not self.sequence_pdf.get() or not os.path.exists(self.sequence_pdf.get()):
            messagebox.showerror("Error", "Please select a valid PDF file first")
            return
        
        if not self.pdf_files:
            messagebox.showerror("Error", "No PDF files loaded to reorder")
            return
        
        try:
            # Extract sequence from bookmarks
            sequence, missing_files = self.extract_bookmark_sequence(self.sequence_pdf.get())
            
            if not sequence:
                messagebox.showwarning("Warning", "No bookmarks found in PDF or could not determine sequence")
                return
            
            # Show preview window
            preview = BookmarkSequencePreview(
                self.root, 
                sequence, 
                missing_files,
                self.pdf_files
            )
            preview.set_apply_callback(self.apply_bookmark_sequence)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract bookmarks: {str(e)}")
    
    def extract_bookmark_sequence(self, pdf_path):
        """Extract sequence from PDF bookmarks"""
        sequence = []
        found_titles = set()
        
        with open(pdf_path, 'rb') as f:
            reader = PdfReader(f)
            
            # Get all bookmarks (outlines)
            if not reader.outline:
                return [], []
            
            # Extract titles from bookmarks
            bookmarks = self.flatten_bookmarks(reader.outline)
            bookmark_titles = [bm.title for bm in bookmarks if hasattr(bm, 'title')]
            
            # Create mapping of basenames to full paths
            basename_to_path = {os.path.basename(p): p for p in self.pdf_files}
            
            # Match bookmark titles to files
            for title in bookmark_titles:
                # Clean the title (remove common extensions if present)
                clean_title = os.path.splitext(title)[0]
                
                # Try to match with and without .pdf extension
                matches = []
                for ext in ['', '.pdf', '.PDF']:
                    potential_match = clean_title + ext
                    if potential_match in basename_to_path:
                        matches.append(basename_to_path[potential_match])
                
                if matches:
                    # Use the first match
                    matched_file = matches[0]
                    sequence.append((matched_file, "found"))
                    found_titles.add(title)
                else:
                    sequence.append((title, "missing"))
        
        # Identify missing files (bookmarks without matches)
        missing_files = [title for title in bookmark_titles if title not in found_titles]
        
        return sequence, missing_files
    
    def flatten_bookmarks(self, outlines, level=0):
        """Flatten nested bookmarks structure"""
        flat = []
        for item in outlines:
            if isinstance(item, list):
                flat.extend(self.flatten_bookmarks(item, level+1))
            else:
                flat.append(item)
        return flat
    
    def apply_bookmark_sequence(self, sequence, missing_files, extra_files, strict_mode):
        """Apply the sequence extracted from bookmarks"""
        if not sequence:
            messagebox.showwarning("Warning", "No valid sequence to apply")
            return
        
        # In strict mode, use only the sequence
        if strict_mode:
            new_order = sequence
        else:
            # In non-strict mode, append extra files
            new_order = sequence + extra_files
        
        # Update the display
        self.pdf_files = new_order
        self.listbox.delete(0, tk.END)
        for pdf in self.pdf_files:
            self.listbox.insert(tk.END, os.path.basename(pdf))
        
        # Show result message
        msg = f"Applied sequence from {len(sequence)} bookmarks"
        if missing_files:
            msg += f"\n{len(missing_files)} bookmarks couldn't be matched"
        if not strict_mode and extra_files:
            msg += f"\nAdded {len(extra_files)} extra files at the end"
        
        messagebox.showinfo("Sequence Applied", msg)
    
    def update_file_list(self):
        self.listbox.delete(0, tk.END)
        self.pdf_files = self.get_pdf_files()
        
        if not self.pdf_files:
            self.listbox.insert(tk.END, "No PDF files found in selected folder")
            return
        
        for pdf in self.pdf_files:
            self.listbox.insert(tk.END, os.path.basename(pdf))
    
    def get_pdf_files(self):
        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            return []
        
        pdf_files = []
        if self.include_subfolders.get():
            for root, _, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith('.pdf'):
                        pdf_files.append(os.path.join(root, file))
        else:
            pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) 
                         if f.lower().endswith('.pdf')]
        
        # Sort files
        if self.sort_by.get() == "name":
            pdf_files.sort()
        elif self.sort_by.get() == "date":
            pdf_files.sort(key=lambda x: os.path.getmtime(x))
        elif self.sort_by.get() == "size":
            pdf_files.sort(key=lambda x: os.path.getsize(x))
        
        return pdf_files
    
    def move_up(self):
        selected = self.listbox.curselection()
        if not selected or selected[0] == 0:
            return
        
        for index in sorted(selected):
            if index == 0:
                continue
            # Swap in listbox
            item = self.listbox.get(index)
            self.listbox.delete(index)
            self.listbox.insert(index-1, item)
            # Swap in pdf_files list
            self.pdf_files[index], self.pdf_files[index-1] = self.pdf_files[index-1], self.pdf_files[index]
        
        # Maintain selection
        new_selection = [i-1 for i in selected if i > 0]
        self.listbox.selection_clear(0, tk.END)
        for i in new_selection:
            self.listbox.selection_set(i)
    
    def move_down(self):
        selected = self.listbox.curselection()
        if not selected or selected[-1] == len(self.pdf_files)-1:
            return
        
        for index in sorted(selected, reverse=True):
            if index == len(self.pdf_files)-1:
                continue
            # Swap in listbox
            item = self.listbox.get(index)
            self.listbox.delete(index)
            self.listbox.insert(index+1, item)
            # Swap in pdf_files list
            self.pdf_files[index], self.pdf_files[index+1] = self.pdf_files[index+1], self.pdf_files[index]
        
        # Maintain selection
        new_selection = [i+1 for i in selected if i < len(self.pdf_files)-1]
        self.listbox.selection_clear(0, tk.END)
        for i in new_selection:
            self.listbox.selection_set(i)
    
    def reverse_order(self):
        self.pdf_files.reverse()
        self.listbox.delete(0, tk.END)
        for pdf in self.pdf_files:
            self.listbox.insert(tk.END, os.path.basename(pdf))
    
    def merge_pdfs(self):
        if not self.pdf_files:
            messagebox.showerror("Error", "No PDF files to merge")
            return
        
        output_file = self.output_filename.get()
        if not output_file:
            messagebox.showerror("Error", "Please enter an output filename")
            return
        
        if not output_file.lower().endswith('.pdf'):
            output_file += '.pdf'
        
        output_path = os.path.join(self.folder_path.get(), output_file)
        if os.path.exists(output_path):
            if not messagebox.askyesno("Confirm", f"{output_file} already exists. Overwrite?"):
                return
        
        try:
            merger = PdfMerger()
            total_files = len(self.pdf_files)
            
            # Store source files in metadata
            merger.add_metadata({
                '/SourceFiles': "; ".join([os.path.basename(p) for p in self.pdf_files])
            })
            
            for i, pdf in enumerate(self.pdf_files, 1):
                try:
                    merger.append(pdf)
                    self.status_label.config(text=f"Adding {os.path.basename(pdf)}...")
                    self.progress['value'] = (i/total_files)*100
                    self.root.update_idletasks()
                except Exception as e:
                    self.listbox.insert(tk.END, f"Error: {os.path.basename(pdf)} - {str(e)}")
                    self.listbox.itemconfig(tk.END, {'fg': 'red'})
            
            with open(output_path, 'wb') as f:
                merger.write(f)
            
            messagebox.showinfo("Success", 
                f"Successfully merged {len(merger.pages)} pages from {total_files} PDFs\n"
                f"Saved to: {output_path}")
            
            self.progress['value'] = 0
            self.status_label.config(text="Ready")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to merge PDFs: {str(e)}")
        finally:
            if 'merger' in locals():
                merger.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFMergerGUI(root)
    root.mainloop()