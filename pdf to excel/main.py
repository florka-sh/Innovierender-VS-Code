import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
from extractor import InvoiceExtractor
import sys

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Fix for ZeroDivisionError on some Windows systems
ctk.set_widget_scaling(1.0)
ctk.set_window_scaling(1.0)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PDF Invoice to Excel Extractor")
        self.geometry("800x600")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --- Header ---
        self.header_frame = ctk.CTkFrame(self)
        self.header_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        
        self.label_title = ctk.CTkLabel(self.header_frame, text="PDF Invoice Extractor", font=("Roboto", 24, "bold"))
        self.label_title.pack(pady=10)

        # --- Controls ---
        self.controls_frame = ctk.CTkFrame(self)
        self.controls_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        self.btn_select_files = ctk.CTkButton(self.controls_frame, text="Select PDFs", command=self.select_files)
        self.btn_select_files.grid(row=0, column=0, padx=10, pady=10)

        self.label_files_count = ctk.CTkLabel(self.controls_frame, text="No files selected")
        self.label_files_count.grid(row=0, column=1, padx=10, pady=10)

        self.btn_process = ctk.CTkButton(self.controls_frame, text="Start Processing", command=self.start_processing_thread, state="disabled", fg_color="green")
        self.btn_process.grid(row=0, column=2, padx=10, pady=10)

        # --- Log / Output ---
        self.log_textbox = ctk.CTkTextbox(self, width=760, height=300)
        self.log_textbox.grid(row=2, column=0, padx=20, pady=20, sticky="nsew")
        self.log_textbox.insert("0.0", "Ready. Please select PDF files to begin.\n")

        # --- State ---
        self.selected_files = []
        self.extractor = InvoiceExtractor() 

    def log(self, message):
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")

    def select_files(self):
        filetypes = (("PDF files", "*.pdf"), ("All files", "*.*"))
        filenames = filedialog.askopenfilenames(title="Select PDF Invoices", filetypes=filetypes)
        if filenames:
            self.selected_files = filenames
            self.label_files_count.configure(text=f"{len(filenames)} files selected")
            self.btn_process.configure(state="normal")
            self.log(f"Selected {len(filenames)} files.")

    def start_processing_thread(self):
        self.btn_process.configure(state="disabled")
        self.btn_select_files.configure(state="disabled")
        
        thread = threading.Thread(target=self.process_files)
        thread.start()

    def process_files(self):
        extracted_data = []
        
        self.log("-" * 30)
        self.log("Starting extraction process...")

        # Check for local Poppler
        poppler_path = None
        
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        local_poppler = os.path.join(base_dir, "poppler", "Library", "bin")
        if os.path.exists(local_poppler):
            self.log(f"Using local Poppler: {local_poppler}")
            poppler_path = local_poppler
        else:
            self.log("Using system Poppler (if available)...")

        # Update extractor with path
        self.extractor.poppler_path = poppler_path

        for i, pdf_path in enumerate(self.selected_files):
            filename = os.path.basename(pdf_path)
            self.log(f"Processing [{i+1}/{len(self.selected_files)}]: {filename}...")
            
            try:
                # Now returns a list of dicts (one per page)
                invoices = self.extractor.extract_invoices_from_pdf(pdf_path)
                
                # Check for fatal file-level error
                if len(invoices) == 1 and "error" in invoices[0]:
                    error_msg = invoices[0]["error"]
                    self.log(f"FAILED: {error_msg}")
                    if "poppler" in error_msg.lower():
                        self.log("HINT: Is Poppler installed and in your PATH?")
                    continue

                for inv_data in invoices:
                    inv_data["filename"] = filename
                    extracted_data.append(inv_data)
                    
                    page_info = f" (Page {inv_data.get('page_number')})" if 'page_number' in inv_data else ""
                    self.log(f"  -> Found Invoice{page_info}:")
                    self.log(f"     Date: {inv_data.get('date', 'N/A')}")
                    self.log(f"     Total: {inv_data.get('total_amount', 'N/A')}")
                    self.log(f"     Invoice #: {inv_data.get('invoice_number', 'N/A')}")
                
            except Exception as e:
                self.log(f"ERROR processing {filename}: {e}")

        if extracted_data:
            base_filename = "extracted_invoices"
            output_file = f"{base_filename}.xlsx"
            
            # Try to find a writable filename
            counter = 1
            while True:
                try:
                    # Check if we can open it for writing (exclusive access check is hard in cross-platform way, 
                    # but we can just try to save and catch the error, or pick a new name if it exists)
                    # Simpler approach: If file exists, check if we can write to it. 
                    # If permission error, increment counter.
                    
                    # Actually, let's just try to save. If it fails due to permission, try next name.
                    # But pandas to_excel truncates by default.
                    # If we want to avoid overwriting *any* existing file, we should check existence.
                    # If we want to overwrite but handle locks, we catch PermissionError.
                    
                    self.extractor.save_to_excel(extracted_data, output_file)
                    break
                except PermissionError:
                    self.log(f"File {output_file} is open. Trying new name...")
                    output_file = f"{base_filename}_{counter}.xlsx"
                    counter += 1
                except Exception as e:
                    self.log(f"Error saving Excel: {e}")
                    return

            self.log("-" * 30)
            self.log(f"SUCCESS! Data saved to {os.path.abspath(output_file)}")
            messagebox.showinfo("Done", f"Extraction complete!\nSaved to {output_file}")
        else:
            self.log("No data extracted.")

        self.btn_process.configure(state="normal")
        self.btn_select_files.configure(state="normal")

if __name__ == "__main__":
    app = App()
    app.mainloop()
