"""
PDF Invoice Extractor - Desktop Application
Modern GUI application for extracting invoice data from PDFs and generating Excel files.
Features editable grid, template-based text generation, and persistent mapping database.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont
import threading
from pathlib import Path
from pdf_extractor import extract_invoices
from excel_generator import generate_excel
import pandas as pd
import shutil
import os


class ModernStyle:
    """Color scheme and styling constants"""
    BG_DARK = "#0f0f23"
    BG_CARD = "#1a1a2e"
    BG_INPUT = "#252541"
    PRIMARY = "#667eea"
    PRIMARY_HOVER = "#5a67d8"
    SUCCESS = "#38ef7d"
    TEXT_PRIMARY = "#ffffff"
    TEXT_SECONDARY = "#a0aec0"
    BORDER = "#2d2d44"


class EditableTreeview(ttk.Treeview):
    """Treeview with editable cells"""
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.bind("<Double-1>", self.on_double_click)
        self.root = master.winfo_toplevel()

    def on_double_click(self, event):
        """Handle double click to edit cell"""
        region = self.identify("region", event.x, event.y)
        if region != "cell":
            return
            
        column = self.identify_column(event.x)
        row_id = self.identify_row(event.y)
        
        if not row_id:
            return
            
        # Get column index (e.g., "#1" -> 0)
        col_idx = int(column[1:]) - 1
        
        # Get current value
        values = self.item(row_id, "values")
        current_value = values[col_idx]
        
        # Get cell coordinates
        x, y, width, height = self.bbox(row_id, column)
        
        # Create entry widget
        entry = tk.Entry(self, bg=ModernStyle.BG_INPUT, fg=ModernStyle.TEXT_PRIMARY, 
                        insertbackground=ModernStyle.TEXT_PRIMARY, relief=tk.FLAT)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus()
        
        def save_edit(event=None):
            new_value = entry.get()
            current_values = list(self.item(row_id, "values"))
            current_values[col_idx] = new_value
            self.item(row_id, values=current_values)
            entry.destroy()
            
            # Trigger event for app to update data model
            self.event_generate("<<TreeviewEdit>>")
            
        def cancel_edit(event=None):
            entry.destroy()
            
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)


class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Invoice Extractor - Lernförderung")
        
        # Make window larger and centered
        window_width = 1400
        window_height = 900
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Make window resizable
        self.root.resizable(True, True)
        self.root.minsize(1000, 700)
        
        self.root.configure(bg=ModernStyle.BG_DARK)
        
        # Data storage
        self.raw_data = []      # Original extracted data
        self.processed_data = [] # Data with accounting columns applied
        self.pdf_path = None
        self.last_saved_excel = None
        self.mapping_data = None # DataFrame for mapping
        
        # Persistent storage setup
        self.data_dir = Path("data")
        self.data_dir.mkdir(exist_ok=True)
        self.mapping_file_path = self.data_dir / "mapping_db.xlsx"
        
        # Configure custom fonts
        self.title_font = tkfont.Font(family="Segoe UI", size=24, weight="bold")
        self.header_font = tkfont.Font(family="Segoe UI", size=14, weight="bold")
        self.normal_font = tkfont.Font(family="Segoe UI", size=10)
        
        # Configure style
        self.setup_styles()
        
        # Create UI
        self.create_widgets()
        
        # Auto-load mapping file if exists
        self.load_stored_mapping()
        
    def setup_styles(self):
        """Configure ttk styles for modern look"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure button style
        style.configure('Primary.TButton',
                       background=ModernStyle.PRIMARY,
                       foreground=ModernStyle.TEXT_PRIMARY,
                       borderwidth=0,
                       focuscolor='none',
                       font=('Segoe UI', 11, 'bold'),
                       padding=(15, 10))
        
        style.map('Primary.TButton',
                 background=[('active', ModernStyle.PRIMARY_HOVER)])
        
        # Configure label style
        style.configure('Title.TLabel',
                       background=ModernStyle.BG_DARK,
                       foreground=ModernStyle.TEXT_PRIMARY,
                       font=('Segoe UI', 24, 'bold'))
        
        style.configure('Header.TLabel',
                       background=ModernStyle.BG_CARD,
                       foreground=ModernStyle.TEXT_PRIMARY,
                       font=('Segoe UI', 12, 'bold'))
        
        # Configure treeview
        style.configure('Treeview',
                       background=ModernStyle.BG_INPUT,
                       foreground=ModernStyle.TEXT_PRIMARY,
                       fieldbackground=ModernStyle.BG_INPUT,
                       borderwidth=0,
                       font=('Segoe UI', 9),
                       rowheight=30)
        
        style.configure('Treeview.Heading',
                       background=ModernStyle.PRIMARY,
                       foreground=ModernStyle.TEXT_PRIMARY,
                       borderwidth=0,
                       font=('Segoe UI', 9, 'bold'))
        
        style.map('Treeview',
                 background=[('selected', ModernStyle.PRIMARY)])
    
    def create_widgets(self):
        """Create all UI widgets"""
        # Main container (Split View)
        main_container = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, bg=ModernStyle.BG_DARK, sashwidth=4)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # LEFT PANEL: Controls & Settings
        left_panel = tk.Frame(main_container, bg=ModernStyle.BG_DARK, width=400)
        main_container.add(left_panel, stretch="never")
        
        # RIGHT PANEL: Data Grid
        right_panel = tk.Frame(main_container, bg=ModernStyle.BG_DARK)
        main_container.add(right_panel, stretch="always")
        
        self.create_left_panel(left_panel)
        self.create_right_panel(right_panel)
        
    def create_left_panel(self, parent):
        """Create controls on the left side"""
        # Header
        header_frame = tk.Frame(parent, bg=ModernStyle.BG_DARK)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(header_frame, text="PDF Extractor", style='Title.TLabel').pack(anchor=tk.W)
        
        # 1. PDF Selection
        pdf_card = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        pdf_card.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(pdf_card, text="1. Datei auswählen", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        self.file_label = tk.Label(pdf_card, text="Keine Datei", bg=ModernStyle.BG_INPUT, fg=ModernStyle.TEXT_SECONDARY, anchor=tk.W)
        self.file_label.pack(fill=tk.X, pady=(0, 10), ipady=5)
        
        btn_frame = tk.Frame(pdf_card, bg=ModernStyle.BG_CARD)
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="📁 Öffnen", style='Primary.TButton', command=self.select_pdf).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.extract_btn = ttk.Button(btn_frame, text="🚀 Extrahieren", style='Primary.TButton', command=self.extract_data, state=tk.DISABLED)
        self.extract_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        self.status_label = tk.Label(pdf_card, text="", bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY)
        self.status_label.pack(pady=(10, 0))
        
        # 2. Global Settings (Scrollable)
        settings_card = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        settings_card.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        ttk.Label(settings_card, text="2. Globale Einstellungen", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        # Scrollable canvas for settings
        canvas = tk.Canvas(settings_card, bg=ModernStyle.BG_CARD, highlightthickness=0)
        scrollbar = ttk.Scrollbar(settings_card, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=ModernStyle.BG_CARD)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        
        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas.find_all()[0], width=event.width)
            
        scroll_frame.bind("<Configure>", on_configure)
        
        # Mapping File Section
        mapping_frame = tk.Frame(scroll_frame, bg=ModernStyle.BG_CARD, pady=10)
        mapping_frame.pack(fill=tk.X)
        
        tk.Label(mapping_frame, text="Mapping-Datenbank (Kundennummer -> Kostenträger)", bg=ModernStyle.BG_CARD, fg=ModernStyle.PRIMARY, font=('Segoe UI', 9, 'bold')).pack(anchor=tk.W)
        
        self.mapping_status = tk.Label(mapping_frame, text="Keine Datenbank geladen", bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY, anchor=tk.W)
        self.mapping_status.pack(fill=tk.X, pady=(5, 5))
        
        ttk.Button(mapping_frame, text="📂 Datenbank hochladen/aktualisieren", style='Primary.TButton', command=self.upload_mapping_file).pack(fill=tk.X)
        
        # Settings Fields
        self.config_entries = {}
        fields = [
            ('FIRMA', '9251'),
            ('SATZART', 'D'),
            ('SOLL_HABEN', 'H'),
            ('BUCH_KREIS', 'RA'),
            ('HABENKONTO', '42200'),
            ('KOSTSTELLE', '190'),
            ('KOSTTRAGER', '190111512110'),
            ('Kostenträgerbezeichnung', 'SPFH/HzE Siegen'),
            ('Bebuchbar', 'Ja'),
        ]
        
        tk.Label(scroll_frame, text="Standardwerte (werden durch Datenbank überschrieben)", bg=ModernStyle.BG_CARD, fg=ModernStyle.PRIMARY, font=('Segoe UI', 9, 'bold')).pack(anchor=tk.W, pady=(15, 5))
        
        for field, default in fields:
            f_frame = tk.Frame(scroll_frame, bg=ModernStyle.BG_CARD, pady=5)
            f_frame.pack(fill=tk.X)
            tk.Label(f_frame, text=field, bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY).pack(anchor=tk.W)
            entry = tk.Entry(f_frame, bg=ModernStyle.BG_INPUT, fg=ModernStyle.TEXT_PRIMARY, relief=tk.FLAT)
            entry.insert(0, default)
            entry.pack(fill=tk.X, ipady=5)
            self.config_entries[field] = entry
            
        # Templates Section
        tk.Label(scroll_frame, text="Templates (Platzhalter: {student}, {subject})", bg=ModernStyle.BG_CARD, fg=ModernStyle.PRIMARY, font=('Segoe UI', 9, 'bold')).pack(anchor=tk.W, pady=(15, 5))
        
        tk.Label(scroll_frame, text="BUCH_TEXT Template", bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY).pack(anchor=tk.W)
        self.buch_text_entry = tk.Entry(scroll_frame, bg=ModernStyle.BG_INPUT, fg=ModernStyle.TEXT_PRIMARY, relief=tk.FLAT)
        self.buch_text_entry.insert(0, "1025 {student} {subject}")
        self.buch_text_entry.pack(fill=tk.X, ipady=5)
        
        # Apply Button
        ttk.Button(settings_card, text="⚡ Auf alle anwenden", style='Primary.TButton', command=self.apply_global_settings).pack(fill=tk.X, pady=(10, 0))
        
        # 3. Export
        export_card = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        export_card.pack(fill=tk.X)
        
        ttk.Label(export_card, text="3. Exportieren", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        self.generate_btn = ttk.Button(export_card, text="💾 Excel speichern", style='Primary.TButton', command=self.generate_excel_new, state=tk.DISABLED)
        self.generate_btn.pack(fill=tk.X)
        
    def create_right_panel(self, parent):
        """Create data grid on the right side"""
        # Header
        header = tk.Frame(parent, bg=ModernStyle.BG_DARK)
        header.pack(fill=tk.X, pady=(0, 10), padx=10)
        
        ttk.Label(header, text="Datenvorschau & Bearbeitung", style='Header.TLabel').pack(side=tk.LEFT)
        tk.Label(header, text="(Doppelklick zum Bearbeiten)", bg=ModernStyle.BG_DARK, fg=ModernStyle.TEXT_SECONDARY).pack(side=tk.LEFT, padx=10)
        
        # Treeview Container
        tree_frame = tk.Frame(parent, bg=ModernStyle.BG_CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # Configure grid layout for tree_frame
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Columns
        self.columns = [
            'SATZART', 'FIRMA', 'BELEG_NR', 'BELEG_DAT', 'SOLL_HABEN', 'BUCH_KREIS', 
            'BUCH_JAHR', 'BUCH_MONAT', 'DEBI_KREDI', 'BETRAG', 'RECHNUNG', 
            'BUCH_TEXT', 'HABENKONTO', 'KOSTSTELLE', 'KOSTTRAGER', 
            'Kostenträgerbezeichnung', 'Bebuchbar'
        ]
        
        self.tree = EditableTreeview(tree_frame, columns=self.columns, show='headings',
                                    yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        # Configure columns
        for col in self.columns:
            self.tree.heading(col, text=col)
            width = 100
            if col in ['BUCH_TEXT', 'Kostenträgerbezeichnung']:
                width = 250
            if col in ['SATZART', 'SOLL_HABEN', 'BUCH_MONAT']:
                width = 60
            self.tree.column(col, width=width, minwidth=60, stretch=False) # stretch=False is key for horizontal scrolling
            
        # Grid layout
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        # Bind edit event
        self.tree.bind("<<TreeviewEdit>>", self.on_tree_edit)

    def select_pdf(self):
        filename = filedialog.askopenfilename(title="PDF auswählen", filetypes=[("PDF", "*.pdf")])
        if filename:
            self.pdf_path = filename
            self.file_label.config(text=Path(filename).name, fg=ModernStyle.TEXT_PRIMARY)
            self.extract_btn.config(state=tk.NORMAL)
            self.status_label.config(text="Bereit zum Extrahieren", fg=ModernStyle.TEXT_SECONDARY)

    def extract_data(self):
        if not self.pdf_path: return
        self.status_label.config(text="⏳ Extrahiere...", fg=ModernStyle.TEXT_SECONDARY)
        self.extract_btn.config(state=tk.DISABLED)
        threading.Thread(target=self._extract_thread).start()

    def _extract_thread(self):
        try:
            self.raw_data = extract_invoices(self.pdf_path)
            self.root.after(0, self._extraction_complete)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Fehler", str(e)))

    def _extraction_complete(self):
        self.status_label.config(text=f"✅ {len(self.raw_data)} Einträge", fg=ModernStyle.SUCCESS)
        self.extract_btn.config(state=tk.NORMAL)
        self.generate_btn.config(state=tk.NORMAL)
        
        # Initial processing with default settings
        self.apply_global_settings()

    def upload_mapping_file(self):
        """Upload and save mapping file"""
        filename = filedialog.askopenfilename(title="Mapping-Datei auswählen", filetypes=[("Excel", "*.xlsx")])
        if not filename: return
        
        try:
            # Verify file structure
            df = pd.read_excel(filename)
            
            # Column mapping (User File -> Internal Standard)
            column_map = {
                'Personenkonto': 'Kundennummer',
                'Kostt Hellern 2025': 'Kostenträger',
                'Kostenträger Bezeichnung': 'Kostenträgerbezeichnung'
            }
            
            # Rename columns if they exist
            df = df.rename(columns=column_map)
            
            # Check for required columns (Internal Standard)
            required_cols = ['Kundennummer', 'Kostenträger', 'Kostenträgerbezeichnung']
            missing = [col for col in required_cols if col not in df.columns]
            
            if missing:
                # Try fuzzy matching or flexible check if exact names don't match
                # For now, just show error with expected names
                messagebox.showerror("Fehler", 
                    f"Die Excel-Datei muss folgende Spalten enthalten (oder deren Äquivalente):\n"
                    f"- Personenkonto (oder Kundennummer)\n"
                    f"- Kostt Hellern 2025 (oder Kostenträger)\n"
                    f"- Kostenträger Bezeichnung\n\n"
                    f"Fehlende Spalten: {', '.join(missing)}")
                return
                
            # Keep only relevant columns
            df = df[required_cols]
            
            # Save to persistent storage
            df.to_excel(self.mapping_file_path, index=False)
            
            # Load it
            self.load_stored_mapping()
            messagebox.showinfo("Erfolg", "Datenbank erfolgreich aktualisiert!")
            
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden: {str(e)}")

    def load_stored_mapping(self):
        """Load mapping file from storage"""
        if self.mapping_file_path.exists():
            try:
                self.mapping_data = pd.read_excel(self.mapping_file_path)
                # Ensure Kundennummer is string for matching
                self.mapping_data['Kundennummer'] = self.mapping_data['Kundennummer'].astype(str).str.replace(' ', '')
                self.mapping_status.config(text="✅ Datenbank aktiv", fg=ModernStyle.SUCCESS)
            except Exception as e:
                self.mapping_status.config(text="❌ Fehler in Datenbank", fg='#f15e64')
        else:
            self.mapping_status.config(text="Keine Datenbank geladen", fg=ModernStyle.TEXT_SECONDARY)

    def apply_global_settings(self):
        """Apply global settings to all rows and update grid"""
        if not self.raw_data: return
        
        # Get current settings
        settings = {k: v.get() for k, v in self.config_entries.items()}
        buch_text_template = self.buch_text_entry.get()
        
        self.processed_data = []
        
        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Process each row
        for item in self.raw_data:
            # Format date
            beleg_dat = ""
            buch_jahr = ""
            buch_monat = ""
            if item.get('invoice_date'):
                try:
                    date_obj = pd.to_datetime(item['invoice_date'], dayfirst=True)
                    beleg_dat = date_obj.strftime('%Y%m%d')
                    buch_jahr = date_obj.year
                    buch_monat = date_obj.month
                except: pass
                
            # Format amount (cents)
            betrag = int(item.get('amount', 0) * 100)
            
            # Generate BUCH_TEXT from template
            buch_text = buch_text_template.format(
                student=item.get('student_name', ''),
                subject=item.get('subject', ''),
                school=item.get('school', ''),
                month=item.get('month_year', '')
            )
            
            # Determine Kostenträger info (Default vs Mapping)
            kosttrager = settings['KOSTTRAGER']
            kost_bez = settings['Kostenträgerbezeichnung']
            
            # Check mapping file
            if self.mapping_data is not None and item.get('customer_number'):
                cust_num = str(item.get('customer_number')).replace(' ', '')
                match = self.mapping_data[self.mapping_data['Kundennummer'] == cust_num]
                if not match.empty:
                    kosttrager = str(match.iloc[0]['Kostenträger'])
                    kost_bez = str(match.iloc[0]['Kostenträgerbezeichnung'])
            
            # Logic: Ensure Kostenträger starts with 0
            if kosttrager and not kosttrager.startswith('0'):
                kosttrager = '0' + kosttrager
                
            # Logic: Koststelle is first 4 digits of Kostenträger
            koststelle = kosttrager[:4] if kosttrager and len(kosttrager) >= 4 else settings['KOSTSTELLE']
            
            row = {
                'SATZART': settings['SATZART'],
                'FIRMA': settings['FIRMA'],
                'BELEG_NR': item.get('invoice_number', ''),
                'BELEG_DAT': beleg_dat,
                'SOLL_HABEN': settings['SOLL_HABEN'],
                'BUCH_KREIS': settings['BUCH_KREIS'],
                'BUCH_JAHR': buch_jahr,
                'BUCH_MONAT': buch_monat,
                'DEBI_KREDI': item.get('customer_number', ''),
                'BETRAG': betrag,
                'RECHNUNG': item.get('invoice_number', ''),
                'BUCH_TEXT': buch_text,
                'HABENKONTO': settings['HABENKONTO'],
                'KOSTSTELLE': koststelle,
                'KOSTTRAGER': kosttrager,
                'Kostenträgerbezeichnung': kost_bez,
                'Bebuchbar': settings['Bebuchbar']
            }
            
            self.processed_data.append(row)
            
            # Add to tree
            values = [row.get(col, '') for col in self.columns]
            self.tree.insert('', tk.END, values=values)
            
    def on_tree_edit(self, event):
        """Update processed_data when tree is edited"""
        # Rebuild processed_data from tree items
        self.processed_data = []
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id, "values")
            row = dict(zip(self.columns, values))
            self.processed_data.append(row)

    def generate_excel_new(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not filename: return
        
        try:
            # Create DataFrame from current grid data
            df = pd.DataFrame(self.processed_data)
            
            # Ensure all 23 columns exist (fill missing with None)
            all_columns = [
                'SATZART', 'FIRMA', 'BELEG_NR', 'BELEG_DAT', 'SOLL_HABEN', 'BUCH_KREIS',
                'BUCH_JAHR', 'BUCH_MONAT', 'DEBI_KREDI', 'BETRAG', 'RECHNUNG', 'leer',
                'BUCH_TEXT', 'HABENKONTO', 'SOLLKONTO', 'leer_1', 'KOSTSTELLE',
                'KOSTTRAGER', 'Kostenträgerbezeichnung', 'Bebuchbar',
                'Debitoren.Bezeichnung', 'Debitoren.Aktuelle Anschrift Anschrift-Zusatz',
                'AbgBenutzerdefiniert'
            ]
            
            for col in all_columns:
                if col not in df.columns:
                    df[col] = None
                    
            # Reorder columns
            df = df[all_columns]
            
            # Save
            df.to_excel(filename, index=False, engine='openpyxl')
            messagebox.showinfo("Erfolg", f"Datei gespeichert:\n{Path(filename).name}")
            
        except Exception as e:
            messagebox.showerror("Fehler", str(e))

def main():
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
