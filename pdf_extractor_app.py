"""
Multi-Project Invoice Processing Application
Supports: PDF Reader, Excel Transformer, and Scanned PDF (OCR - future)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont
import threading
from pathlib import Path
from pdf_extractor import extract_invoices
from excel_generator import generate_excel
from bereitspf_transformer import transform_excel as bereitspf_transform
import pandas as pd
import shutil


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

    def on_double_click(self, event):
        """Handle double click to edit cell"""
        region = self.identify("region", event.x, event.y)
        if region != "cell":
            return
            
        column = self.identify_column(event.x)
        row_id = self.identify_row(event.y)
        
        if not row_id:
            return
            
        col_idx = int(column[1:]) - 1
        values = self.item(row_id, "values")
        current_value = values[col_idx]
        
        x, y, width, height = self.bbox(row_id, column)
        
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
            self.event_generate("<<TreeviewEdit>>")
            
        def cancel_edit(event=None):
            entry.destroy()
            
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)


class MultiProjectApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Processing App")
        
        # Window setup
        window_width = 1400
        window_height = 900
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        self.root.resizable(True, True)
        self.root.minsize(1000, 700)
        self.root.configure(bg=ModernStyle.BG_DARK)
        
        # Data storage
        self.processed_data = []
        self.current_project = "PDF Reader"
        
        # PDF Reader specific
        self.raw_pdf_data = []
        self.pdf_path = None
        self.mapping_data = None
        # Extended mapping (with FIRMA & DEBI_KREDI details, e.g. kostentreäger_info_3.xlsx)
        self.extended_mapping = None
        self.data_dir = Path("data")
        self.data_dir.mkdir(exist_ok=True)
        self.mapping_file_path = self.data_dir / "mapping_db.xlsx"
        
        # Excel Transformer specific
        self.template_path = None
        self.source_path = None
        
        # Fonts
        self.title_font = tkfont.Font(family="Segoe UI", size=24, weight="bold")
        self.header_font = tkfont.Font(family="Segoe UI", size=14, weight="bold")
        
        self.setup_styles()
        self.create_widgets()
        
        # Auto-load PDF Reader mapping
        self.load_stored_mapping()
        self.load_extended_mapping()
        
    def setup_styles(self):
        """Configure ttk styles"""
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('Primary.TButton',
                       background=ModernStyle.PRIMARY,
                       foreground=ModernStyle.TEXT_PRIMARY,
                       borderwidth=0,
                       font=('Segoe UI', 11, 'bold'),
                       padding=(15, 10))
        
        style.map('Primary.TButton',
                 background=[('active', ModernStyle.PRIMARY_HOVER)])
        
        style.configure('Title.TLabel',
                       background=ModernStyle.BG_DARK,
                       foreground=ModernStyle.TEXT_PRIMARY,
                       font=('Segoe UI', 24, 'bold'))
        
        style.configure('Header.TLabel',
                       background=ModernStyle.BG_CARD,
                       foreground=ModernStyle.TEXT_PRIMARY,
                       font=('Segoe UI', 12, 'bold'))
        
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
        """Create main UI"""
        # Main container
        main_container = tk.Frame(self.root, bg=ModernStyle.BG_DARK)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header with project selector
        header_frame = tk.Frame(main_container, bg=ModernStyle.BG_DARK)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(header_frame, text="Invoice Processing", style='Title.TLabel').pack(side=tk.LEFT)
        
        # Project selector
        selector_frame = tk.Frame(header_frame, bg=ModernStyle.BG_DARK)
        selector_frame.pack(side=tk.RIGHT)
        
        tk.Label(selector_frame, text="Project:", bg=ModernStyle.BG_DARK, fg=ModernStyle.TEXT_SECONDARY, font=('Segoe UI', 11)).pack(side=tk.LEFT, padx=(0, 10))
        
        self.project_var = tk.StringVar(value="PDF Reader")
        project_selector = ttk.Combobox(selector_frame, textvariable=self.project_var, 
                                       values=["PDF Reader", "Excel Transformer", "Scanned PDF (OCR)"],
                                       state='readonly', width=20, font=('Segoe UI', 11))
        project_selector.pack(side=tk.LEFT)
        project_selector.bind('<<ComboboxSelected>>', self.on_project_change)
        
        # Content area (split view)
        content_container = tk.PanedWindow(main_container, orient=tk.HORIZONTAL, bg=ModernStyle.BG_DARK, sashwidth=4)
        content_container.pack(fill=tk.BOTH, expand=True)
        
        # Left panel
        left_panel = tk.Frame(content_container, bg=ModernStyle.BG_DARK, width=400)
        content_container.add(left_panel, stretch="never")
        
        # Right panel (data grid)
        right_panel = tk.Frame(content_container, bg=ModernStyle.BG_DARK)
        content_container.add(right_panel, stretch="always")
        
        # Create mode-specific panels
        self.pdf_reader_panel = tk.Frame(left_panel, bg=ModernStyle.BG_DARK)
        self.excel_transformer_panel = tk.Frame(left_panel, bg=ModernStyle.BG_DARK)
        self.ocr_panel = tk.Frame(left_panel, bg=ModernStyle.BG_DARK)
        
        self.create_pdf_reader_ui(self.pdf_reader_panel)
        self.create_excel_transformer_ui(self.excel_transformer_panel)
        self.create_ocr_ui(self.ocr_panel)
        self.create_data_grid(right_panel)
        
        # Show PDF Reader by default
        self.pdf_reader_panel.pack(fill=tk.BOTH, expand=True)
    
    def on_project_change(self, event=None):
        """Handle project selection change"""
        new_project = self.project_var.get()
        
        # Hide all panels
        self.pdf_reader_panel.pack_forget()
        self.excel_transformer_panel.pack_forget()
        self.ocr_panel.pack_forget()
        
        # Show selected panel
        if new_project == "PDF Reader":
            self.pdf_reader_panel.pack(fill=tk.BOTH, expand=True)
        elif new_project == "Excel Transformer":
            self.excel_transformer_panel.pack(fill=tk.BOTH, expand=True)
        elif new_project == "Scanned PDF (OCR)":
            self.ocr_panel.pack(fill=tk.BOTH, expand=True)
        
        self.current_project = new_project
        
        # Clear data grid
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.processed_data = []
    
    def create_pdf_reader_ui(self, parent):
        """Create PDF Reader mode UI"""
        # Card style
        card1 = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        card1.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(card1, text="1. PDF-Datei", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        self.pdf_file_label = tk.Label(card1, text="Keine Datei", bg=ModernStyle.BG_INPUT, fg=ModernStyle.TEXT_SECONDARY, anchor=tk.W)
        self.pdf_file_label.pack(fill=tk.X, pady=(0, 10), ipady=5)
        
        btn_frame = tk.Frame(card1, bg=ModernStyle.BG_CARD)
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="📁 Öffnen", style='Primary.TButton', command=self.select_pdf).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.pdf_extract_btn = ttk.Button(btn_frame, text="🚀 Extrahieren", style='Primary.TButton', command=self.extract_pdf_data, state=tk.DISABLED)
        self.pdf_extract_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        self.pdf_status_label = tk.Label(card1, text="", bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY)
        self.pdf_status_label.pack(pady=(10, 0))
        
        # Mapping DB card
        card2 = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        card2.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(card2, text="Mapping-Datenbank", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        self.pdf_mapping_status = tk.Label(card2, text="Keine Datenbank", bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY)
        self.pdf_mapping_status.pack(fill=tk.X, pady=(0,  10))
        
        ttk.Button(card2, text="📂 Datenbank hochladen", style='Primary.TButton', command=self.upload_pdf_mapping).pack(fill=tk.X)
        
        # Settings card with scrollable area
        card3 = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        card3.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        ttk.Label(card3, text="2. Einstellungen", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        canvas = tk.Canvas(card3, bg=ModernStyle.BG_CARD, highlightthickness=0)
        scrollbar = ttk.Scrollbar(card3, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=ModernStyle.BG_CARD)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        
        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas.find_all()[0], width=event.width)
            
        scroll_frame.bind("<Configure>", on_configure)
        
        self.pdf_config_entries = {}
        fields = [
            ('FIRMA', '9251'), ('SATZART', 'D'), ('SOLL_HABEN', 'H'),
            ('BUCH_KREIS', 'RA'), ('HABENKONTO', '42200'), ('KOSTSTELLE', '190'),
            ('KOSTTRAGER', '190111512110'), ('Kostenträgerbezeichnung', 'SPFH/HzE Siegen'),
            ('Bebuchbar', 'Ja'),
        ]
        
        for field, default in fields:
            f_frame = tk.Frame(scroll_frame, bg=ModernStyle.BG_CARD, pady=5)
            f_frame.pack(fill=tk.X)
            tk.Label(
                f_frame, text=field, bg=ModernStyle.BG_CARD,
                fg=ModernStyle.TEXT_SECONDARY
            ).pack(anchor=tk.W)
            # Use Combobox for FIRMA to pick from mapping DB
            if field == 'FIRMA':
                self.firma_var = tk.StringVar(value=default)
                try:
                    options = []
                    # Primary source: mapping DB if it has FIRMA
                    if (
                        self.mapping_data is not None and
                        'FIRMA' in self.mapping_data.columns
                    ):
                        options = sorted(
                            set(
                                self.mapping_data['FIRMA']
                                .astype(str).str.strip()
                            )
                        )
                    else:
                        # Fallback: use kostentreäger_info_3.xlsx if present
                        try:
                            from pathlib import Path
                            alt_path = Path('kostentreäger_info_3.xlsx')
                            if alt_path.exists():
                                import pandas as pd
                                alt_df = pd.read_excel(alt_path)
                                if 'FIRMA' in alt_df.columns:
                                    options = sorted(
                                        set(
                                            alt_df['FIRMA']
                                            .astype(str).str.strip()
                                        )
                                    )
                        except Exception:
                            pass
                    self.firma_combo = ttk.Combobox(
                        f_frame, textvariable=self.firma_var,
                        values=options, state='readonly'
                    )
                    self.firma_combo.pack(fill=tk.X, ipady=0)
                    entry_widget = self.firma_combo
                except Exception:
                    entry = tk.Entry(
                        f_frame, bg=ModernStyle.BG_INPUT,
                        fg=ModernStyle.TEXT_PRIMARY, relief=tk.FLAT
                    )
                    entry.insert(0, default)
                    entry.pack(fill=tk.X, ipady=5)
                    entry_widget = entry
            else:
                entry = tk.Entry(
                    f_frame, bg=ModernStyle.BG_INPUT,
                    fg=ModernStyle.TEXT_PRIMARY, relief=tk.FLAT
                )
                entry.insert(0, default)
                entry.pack(fill=tk.X, ipady=5)
                entry_widget = entry
            self.pdf_config_entries[field] = entry_widget
        
        tk.Label(scroll_frame, text="BUCH_TEXT Template", bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY).pack(anchor=tk.W, pady=(15, 0))
        self.pdf_buch_text_entry = tk.Entry(scroll_frame, bg=ModernStyle.BG_INPUT, fg=ModernStyle.TEXT_PRIMARY, relief=tk.FLAT)
        self.pdf_buch_text_entry.insert(0, "1025 {student} {subject}")
        self.pdf_buch_text_entry.pack(fill=tk.X, ipady=5)
        
        # Buttons: apply and check matches
        btns = tk.Frame(card3, bg=ModernStyle.BG_CARD)
        btns.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(
            btns, text="⚡ Auf alle anwenden", style='Primary.TButton',
            command=self.apply_pdf_settings
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(
            btns, text="🔎 Prüfe Matches", style='Secondary.TButton',
            command=self.preview_pdf_matches
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5,0))
        
        # Export card
        card4 = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        card4.pack(fill=tk.X)
        
        ttk.Label(card4, text="3. Exportieren", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        self.pdf_export_btn = ttk.Button(card4, text="💾 Excel speichern", style='Primary.TButton', command=self.export_excel, state=tk.DISABLED)
        self.pdf_export_btn.pack(fill=tk.X)
    
    def create_excel_transformer_ui(self, parent):
        """Create Excel Transformer mode UI"""
        # Template file
        card1 = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        card1.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(card1, text="1. Template Excel", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        self.template_label = tk.Label(card1, text="Keine Datei", bg=ModernStyle.BG_INPUT, fg=ModernStyle.TEXT_SECONDARY, anchor=tk.W)
        self.template_label.pack(fill=tk.X, pady=(0, 10), ipady=5)
        
        ttk.Button(card1, text="📁 Template wählen", style='Primary.TButton', command=self.select_template).pack(fill=tk.X)
        
        # Source file
        card2 = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        card2.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(card2, text="2. Quelldatei Excel", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        self.source_label = tk.Label(card2, text="Keine Datei", bg=ModernStyle.BG_INPUT, fg=ModernStyle.TEXT_SECONDARY, anchor=tk.W)
        self.source_label.pack(fill=tk.X, pady=(0, 10), ipady=5)
        
        ttk.Button(card2, text="📁 Quelldatei wählen", style='Primary.TButton', command=self.select_source).pack(fill=tk.X)
        
        # Defaults - FIXED VERSION
        card3 = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        card3.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(card3, text="3. Standardwerte", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        # Create container with fixed height
        defaults_container = tk.Frame(card3, bg=ModernStyle.BG_CARD, height=200)
        defaults_container.pack(fill=tk.BOTH, expand=False)
        defaults_container.pack_propagate(False)  # Prevent shrinking
        
        canvas = tk.Canvas(defaults_container, bg=ModernStyle.BG_CARD, highlightthickness=0)
        scrollbar = ttk.Scrollbar(defaults_container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=ModernStyle.BG_CARD)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        
        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Update canvas window width
            canvas.itemconfig(canvas_window, width=canvas.winfo_width())
            
        scroll_frame.bind("<Configure>", on_configure)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))
        
        self.excel_config_entries = {}
        fields = [
            ('SATZART', 'D'), ('FIRMA', '9241'), ('SOLL_HABEN', 'S'),
            ('BUCH_KREIS', 'RE'), ('BUCH_JAHR', '2025'), ('BUCH_MONAT', '11'),
            ('Bebuchbar', 'Ja'),
        ]
        
        for field, default in fields:
            f_frame = tk.Frame(scroll_frame, bg=ModernStyle.BG_CARD, pady=3)
            f_frame.pack(fill=tk.X, padx=5)
            tk.Label(f_frame, text=field, bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY, font=('Segoe UI', 9)).pack(anchor=tk.W)
            
            # Use Combobox for FIRMA
            if field == 'FIRMA':
                firma_var = tk.StringVar(value=default)
                try:
                    options = []
                    if (
                        self.mapping_data is not None and
                        'FIRMA' in self.mapping_data.columns
                    ):
                        options = sorted(
                            set(
                                self.mapping_data['FIRMA']
                                .astype(str).str.strip()
                            )
                        )
                    else:
                        try:
                            from pathlib import Path
                            alt_path = Path('kostentreäger_info_3.xlsx')
                            if alt_path.exists():
                                import pandas as pd
                                alt_df = pd.read_excel(alt_path)
                                if 'FIRMA' in alt_df.columns:
                                    options = sorted(
                                        set(
                                            alt_df['FIRMA']
                                            .astype(str).str.strip()
                                        )
                                    )
                        except Exception:
                            pass
                    combo = ttk.Combobox(
                        f_frame, textvariable=firma_var,
                        values=options, state='readonly',
                        font=('Segoe UI', 9)
                    )
                    combo.pack(fill=tk.X, ipady=0)
                    entry_widget = combo
                except Exception:
                    entry = tk.Entry(
                        f_frame, bg=ModernStyle.BG_INPUT,
                        fg=ModernStyle.TEXT_PRIMARY, relief=tk.FLAT,
                        font=('Segoe UI', 9)
                    )
                    entry.insert(0, default)
                    entry.pack(fill=tk.X, ipady=3)
                    entry_widget = entry
            else:
                entry = tk.Entry(
                    f_frame, bg=ModernStyle.BG_INPUT,
                    fg=ModernStyle.TEXT_PRIMARY, relief=tk.FLAT,
                    font=('Segoe UI', 9)
                )
                entry.insert(0, default)
                entry.pack(fill=tk.X, ipady=3)
                entry_widget = entry
            self.excel_config_entries[field] = entry_widget
        
        # Transform button
        card4 = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        card4.pack(fill=tk.X)
        
        ttk.Label(card4, text="4. Transformieren", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        self.excel_transform_btn = ttk.Button(card4, text="🔄 Transformieren", style='Primary.TButton', command=self.transform_excel_data, state=tk.DISABLED)
        self.excel_transform_btn.pack(fill=tk.X, pady=(0, 10))
        
        self.excel_export_btn = ttk.Button(card4, text="💾 Excel speichern", style='Primary.TButton', command=self.export_excel, state=tk.DISABLED)
        self.excel_export_btn.pack(fill=tk.X)
    
    def create_data_grid(self, parent):
        """Create shared data grid"""
        header = tk.Frame(parent, bg=ModernStyle.BG_DARK)
        header.pack(fill=tk.X, pady=(0, 10), padx=10)
        
        ttk.Label(header, text="Datenvorschau & Bearbeitung", style='Header.TLabel').pack(side=tk.LEFT)
        tk.Label(header, text="(Doppelklick zum Bearbeiten)", bg=ModernStyle.BG_DARK, fg=ModernStyle.TEXT_SECONDARY).pack(side=tk.LEFT, padx=10)
        
        tree_frame = tk.Frame(parent, bg=ModernStyle.BG_CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
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
        
        for col in self.columns:
            self.tree.heading(col, text=col)
            width = 100
            if col in ['BUCH_TEXT', 'Kostenträgerbezeichnung']:
                width = 250
            if col in ['SATZART', 'SOLL_HABEN', 'BUCH_MONAT']:
                width = 60
            self.tree.column(col, width=width, minwidth=60, stretch=False)
            
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        self.tree.bind("<<TreeviewEdit>>", self.on_tree_edit)

    def create_ocr_ui(self, parent):
        """Create Scanned PDF (OCR) mode UI"""
        card1 = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        card1.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(card1, text="1. Scanned PDF (OCR)", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))

        self.ocr_file_label = tk.Label(card1, text="Keine Datei", bg=ModernStyle.BG_INPUT, fg=ModernStyle.TEXT_SECONDARY, anchor=tk.W)
        self.ocr_file_label.pack(fill=tk.X, pady=(0, 10), ipady=5)

        btn_frame = tk.Frame(card1, bg=ModernStyle.BG_CARD)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="📁 PDF wählen", style='Primary.TButton', command=self.select_ocr_pdf).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(btn_frame, text="🗂 Vorverarbeitet", style='Primary.TButton', command=self.load_pre_extracted_ocr).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5,5))
        self.ocr_extract_btn = ttk.Button(btn_frame, text="🔎 OCR Extrahieren", style='Primary.TButton', command=self.extract_ocr_data, state=tk.DISABLED)
        self.ocr_extract_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,0))

        self.ocr_status_label = tk.Label(card1, text="", bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY)
        self.ocr_status_label.pack(pady=(10,0))

        # Progress bar for OCR
        self.ocr_progress = ttk.Progressbar(card1, mode='determinate', length=300)
        self.ocr_progress.pack(fill=tk.X, pady=(5,0))
        
        self.ocr_progress_label = tk.Label(card1, text="", bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY, font=('Segoe UI', 8))
        self.ocr_progress_label.pack(pady=(2,0))

        # Help text for OCR setup
        help_text = tk.Label(card1, 
                            text="ℹ️ OCR benötigt Tesseract/Poppler unter InvoiceExtractor/_internal/\nBei Fehlern siehe README → Fehlerbehebung",
                            bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY, 
                            font=('Segoe UI', 8), justify=tk.LEFT)
        help_text.pack(pady=(5, 0), anchor=tk.W)

        # simple defaults for OCR mode
        card2 = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        card2.pack(fill=tk.X, pady=(0, 15))
        ttk.Label(card2, text="2. Einstellungen (Standard)", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 8))

        self.ocr_config_entries = {}
        fields = [ ('SATZART','D'), ('FIRMA','9251'), ('SOLL_HABEN','H'), ('BUCH_KREIS','RA'), ('HABENKONTO','42200') ]
        for field, default in fields:
            f_frame = tk.Frame(card2, bg=ModernStyle.BG_CARD, pady=3)
            f_frame.pack(fill=tk.X, padx=5)
            tk.Label(f_frame, text=field, bg=ModernStyle.BG_CARD, fg=ModernStyle.TEXT_SECONDARY, font=('Segoe UI', 9)).pack(anchor=tk.W)
            
            # Use Combobox for FIRMA
            if field == 'FIRMA':
                firma_var = tk.StringVar(value=default)
                try:
                    options = []
                    if (
                        self.mapping_data is not None and
                        'FIRMA' in self.mapping_data.columns
                    ):
                        options = sorted(
                            set(
                                self.mapping_data['FIRMA']
                                .astype(str).str.strip()
                            )
                        )
                    else:
                        try:
                            from pathlib import Path
                            alt_path = Path('kostentreäger_info_3.xlsx')
                            if alt_path.exists():
                                import pandas as pd
                                alt_df = pd.read_excel(alt_path)
                                if 'FIRMA' in alt_df.columns:
                                    options = sorted(
                                        set(
                                            alt_df['FIRMA']
                                            .astype(str).str.strip()
                                        )
                                    )
                        except Exception:
                            pass
                    combo = ttk.Combobox(
                        f_frame, textvariable=firma_var,
                        values=options, state='readonly',
                        font=('Segoe UI', 9)
                    )
                    combo.pack(fill=tk.X, ipady=0)
                    entry_widget = combo
                except Exception:
                    entry = tk.Entry(
                        f_frame, bg=ModernStyle.BG_INPUT,
                        fg=ModernStyle.TEXT_PRIMARY, relief=tk.FLAT,
                        font=('Segoe UI', 9)
                    )
                    entry.insert(0, default)
                    entry.pack(fill=tk.X, ipady=3)
                    entry_widget = entry
            else:
                entry = tk.Entry(
                    f_frame, bg=ModernStyle.BG_INPUT,
                    fg=ModernStyle.TEXT_PRIMARY, relief=tk.FLAT,
                    font=('Segoe UI', 9)
                )
                entry.insert(0, default)
                entry.pack(fill=tk.X, ipady=3)
                entry_widget = entry
            self.ocr_config_entries[field] = entry_widget

        # Apply settings button
        ttk.Button(card2, text="✓ Einstellungen anwenden", style='Primary.TButton', command=self.apply_ocr_settings).pack(fill=tk.X, pady=(10, 0))

        card3 = tk.Frame(parent, bg=ModernStyle.BG_CARD, padx=15, pady=15)
        card3.pack(fill=tk.X)
        ttk.Label(card3, text="3. Export", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        self.ocr_export_btn = ttk.Button(card3, text="💾 Excel speichern", style='Primary.TButton', command=self.export_excel, state=tk.DISABLED)
        self.ocr_export_btn.pack(fill=tk.X)
    
    # PDF Reader Methods
    def select_pdf(self):
        filename = filedialog.askopenfilename(title="PDF auswählen", filetypes=[("PDF", "*.pdf")])
        if filename:
            self.pdf_path = filename
            self.pdf_file_label.config(text=Path(filename).name, fg=ModernStyle.TEXT_PRIMARY)
            self.pdf_extract_btn.config(state=tk.NORMAL)
            self.pdf_status_label.config(text="Bereit", fg=ModernStyle.TEXT_SECONDARY)

    def extract_pdf_data(self):
        if not self.pdf_path: return
        self.pdf_status_label.config(text="⏳ Extrahiere...", fg=ModernStyle.TEXT_SECONDARY)
        self.pdf_extract_btn.config(state=tk.DISABLED)
        threading.Thread(target=self._extract_pdf_thread).start()

    def _extract_pdf_thread(self):
        try:
            self.raw_pdf_data = extract_invoices(self.pdf_path)
            self.root.after(0, self._pdf_extraction_complete)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Fehler", str(e)))

    def _pdf_extraction_complete(self):
        self.pdf_status_label.config(text=f"✅ {len(self.raw_pdf_data)} Einträge", fg=ModernStyle.SUCCESS)
        self.pdf_extract_btn.config(state=tk.NORMAL)
        self.pdf_export_btn.config(state=tk.NORMAL)
        self.apply_pdf_settings()

    # OCR methods
    def select_ocr_pdf(self):
        """Always let user choose a scanned PDF; separate button loads pre-extracted Excel."""
        filename = filedialog.askopenfilename(title="Scanned PDF auswählen", filetypes=[("PDF", "*.pdf")])
        if filename:
            self.ocr_path = filename
            self.ocr_file_label.config(text=Path(filename).name, fg=ModernStyle.TEXT_PRIMARY)
            self.ocr_extract_btn.config(state=tk.NORMAL)
            self.ocr_status_label.config(text="Bereit (PDF gewählt)", fg=ModernStyle.TEXT_SECONDARY)

    def load_pre_extracted_ocr(self):
        """Load pre-extracted Excel instead of performing OCR."""
        extracted_file = Path("extracted_invoices.xlsx")
        if not extracted_file.exists():
            messagebox.showerror("Fehler", "extracted_invoices.xlsx nicht gefunden.")
            return
        try:
            df = pd.read_excel(extracted_file)
            extraction_log = {
                'filename': extracted_file.name,
                'filepath': str(extracted_file),
                'total_pages': len(df),
                'start_time': pd.Timestamp.now().isoformat(),
                'status': 'completed',
                'pages_completed': len(df),
                'errors': [],
                'end_time': pd.Timestamp.now().isoformat(),
                'output_file': str(extracted_file)
            }
            self.raw_ocr_data = (extraction_log, df)
            self.ocr_file_label.config(text="extracted_invoices.xlsx (geladen)", fg=ModernStyle.TEXT_PRIMARY)
            self.ocr_status_label.config(text="✅ Vorverarbeitete Daten geladen", fg=ModernStyle.SUCCESS)
            self.ocr_extract_btn.config(state=tk.DISABLED)
            self.apply_ocr_settings()
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte Datei nicht laden: {e}")

    def extract_ocr_data(self):
        if not getattr(self, 'ocr_path', None):
            return
        self.ocr_status_label.config(text="⏳ OCR läuft...", fg=ModernStyle.TEXT_SECONDARY)
        self.ocr_progress['value'] = 0
        self.ocr_progress_label.config(text="Starte OCR...")
        self.ocr_extract_btn.config(state=tk.DISABLED)
        threading.Thread(target=self._extract_ocr_thread).start()

    def _extract_ocr_thread(self):
        try:
            # Run OCR extraction on the selected PDF
            try:
                from ocr_analysis.poppler_extractor import PopperExtractor
            except Exception as ie:
                self.root.after(0, lambda: messagebox.showerror("Fehler", f"OCR dependencies missing or import failed: {ie}"))
                self.root.after(0, lambda: self.ocr_status_label.config(text="Fehler beim Starten von OCR", fg='#f15e64'))
                return

            extractor = PopperExtractor()
            extraction_log, df = extractor.extract_pdf(self.ocr_path, output_folder='temp_analysis', 
                                                       progress_callback=self._update_ocr_progress)
            self.raw_ocr_data = (extraction_log, df)
            self.root.after(0, self._ocr_extraction_complete)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Fehler", str(e)))

    def _update_ocr_progress(self, current, total, elapsed_time):
        """Update progress bar during OCR extraction"""
        percentage = int((current / total) * 100) if total > 0 else 0
        self.root.after(0, lambda: self.ocr_progress.config(value=percentage))
        
        # Estimate remaining time
        if current > 0 and elapsed_time > 0:
            avg_time_per_page = elapsed_time / current
            remaining_pages = total - current
            estimated_remaining = avg_time_per_page * remaining_pages
            
            if estimated_remaining < 60:
                time_str = f"{int(estimated_remaining)}s verbleibend"
            else:
                time_str = f"{int(estimated_remaining/60)}m {int(estimated_remaining%60)}s verbleibend"
            
            progress_text = f"{percentage}% ({current}/{total} Seiten) - {time_str}"
        else:
            progress_text = f"{percentage}% ({current}/{total} Seiten)"
        
        self.root.after(0, lambda: self.ocr_progress_label.config(text=progress_text))

    def _ocr_extraction_complete(self):
        if getattr(self, 'raw_ocr_data', None) is None:
            self.ocr_status_label.config(text="Fehler bei OCR", fg='#f15e64')
            self.ocr_extract_btn.config(state=tk.NORMAL)
            return

        log, df = self.raw_ocr_data
        pages = log.get('pages_completed', 0) if isinstance(log, dict) else (len(df) if df is not None else 0)
        self.ocr_progress['value'] = 100
        self.ocr_progress_label.config(text=f"100% ({pages}/{pages} Seiten) - Abgeschlossen")
        self.ocr_status_label.config(text=f"✅ {pages} Seiten extrahiert", fg=ModernStyle.SUCCESS)
        self.ocr_extract_btn.config(state=tk.NORMAL)
        self.ocr_export_btn.config(state=tk.NORMAL)
        self.apply_ocr_settings()

    def apply_ocr_settings(self):
        """Apply OCR settings with field validation"""
        if not getattr(self, 'raw_ocr_data', None):
            return

        log, df = self.raw_ocr_data
        if df is None or df.empty:
            messagebox.showerror(
                "Fehler",
                "Keine Daten aus OCR extrahiert"
            )
            return

        defaults = {
            k: v.get() for k, v in self.ocr_config_entries.items()
        }

        # Import validator
        try:
            from ocr_analysis.field_validators import FieldValidator
        except ImportError:
            FieldValidator = None

        self.processed_data = []
        
        for idx, row_data in df.iterrows():
            # Extract invoice number from OCR data
            inv_num = str(
                row_data.get('Invoice Number', '')
            ).strip()
            
            # Date parsing (DD.MM.YYYY to YYYYMMDD)
            date_str = str(row_data.get('Date', '')).strip()
            beleg_dat = ''
            buch_jahr = ''
            buch_monat = ''
            
            if date_str and date_str != '':
                try:
                    dt = pd.to_datetime(
                        date_str,
                        dayfirst=True,
                        errors='coerce'
                    )
                    if not pd.isna(dt):
                        beleg_dat = dt.strftime('%Y%m%d')
                        buch_jahr = dt.year
                        buch_monat = dt.month
                except Exception:
                    pass

            # Convert amount: "1.234,56" → 123456 cents
            def euros_to_cents(val):
                if val is None or val == '':
                    return None
                try:
                    s = str(val).strip()
                    if s == '':
                        return None
                    
                    # Check if already a large number (likely cents)
                    # Pattern: no comma, large value (>1000)
                    try:
                        num = float(s.replace(',', '.'))
                        # If >1000 and no comma separator, assume cents
                        if num > 1000 and ',' not in s:
                            return int(round(num))
                    except Exception:
                        pass
                    
                    # Convert from euros: 1.234,56 -> 123456 cents
                    parts = s.split(',')
                    if len(parts) == 2:
                        whole, frac = parts
                        whole = whole.replace('.', '')
                        s = f"{whole}.{frac}"
                    else:
                        s = s.replace(',', '.')
                    return int(round(float(s) * 100))
                except Exception:
                    return None

            betrag = euros_to_cents(
                row_data.get('Line Total')
            )
            if betrag is None:
                betrag = euros_to_cents(
                    row_data.get('Total Amount')
                )

            # Extract customer/debitor number
            # Priority: Customer Number field, then invoice suffix
            debi_kredi = str(
                row_data.get('Customer Number', '')
            ).strip()
            
            if not debi_kredi:
                # Fallback to invoice suffix (number after /)
                invoice_suffix = str(
                    row_data.get('Invoice Suffix', '')
                ).strip()
                if invoice_suffix:
                    debi_kredi = invoice_suffix
            
            # Build booking text
            recipient = str(
                row_data.get('Recipient Name', '')
            ).strip()
            description = str(
                row_data.get('Description', '')
            ).strip()[:100]
            
            buch_text = f"{recipient} {description}".strip()

            row = {
                'SATZART': defaults.get('SATZART', 'D'),
                'FIRMA': defaults.get('FIRMA', ''),
                'BELEG_NR': inv_num,
                'BELEG_DAT': beleg_dat,
                'SOLL_HABEN': defaults.get('SOLL_HABEN', ''),
                'BUCH_KREIS': defaults.get('BUCH_KREIS', ''),
                'BUCH_JAHR': buch_jahr,
                'BUCH_MONAT': buch_monat,
                'DEBI_KREDI': debi_kredi,
                'BETRAG': (
                    betrag if betrag is not None else ''
                ),
                'RECHNUNG': inv_num or '',
                'BUCH_TEXT': buch_text,
                'HABENKONTO': defaults.get('HABENKONTO', ''),
                'KOSTSTELLE': '',
                'KOSTTRAGER': '',
                'Kostenträgerbezeichnung': '',
                'Bebuchbar': 'Ja',
                'ocr_confidence': row_data.get(
                    'ocr_confidence', 0.0
                ),
                'validation_required': row_data.get(
                    'validation_required', False
                )
            }

            # Enrich / correct DEBI_KREDI using extended mapping if missing or shortened
            try:
                if self.extended_mapping is not None:
                    debi_val = str(row.get('DEBI_KREDI', '')).strip()
                    firma_val = str(defaults.get('FIRMA', '')).strip()
                    matched_row = None
                    # Criteria: empty debi OR numeric too short (< 7) OR suffix match exists in extended mapping
                    if (not debi_val) or (len(debi_val) < 7):
                        cand = self.extended_mapping
                        if firma_val and 'FIRMA' in cand.columns:
                            cand = cand[cand['FIRMA'] == firma_val]
                        # Try exact Kostenträger match (remove leading zero if added)
                        raw_kost = str(row.get('KOSTTRAGER', '')).lstrip('0')
                        k_match = cand[cand['Kostenträger'].str.lstrip('0') == raw_kost]
                        if not k_match.empty:
                            matched_row = k_match.iloc[0]
                            row['DEBI_KREDI'] = matched_row['DEBI_KREDI']
                        else:
                            # Fallback: any DEBI_KREDI ending with current debi_val
                            if debi_val:
                                suf = cand[cand['DEBI_KREDI'].str.endswith(debi_val)]
                                if not suf.empty:
                                    matched_row = suf.iloc[0]
                                    row['DEBI_KREDI'] = matched_row['DEBI_KREDI']
                    else:
                        # If we have a short value that is suffix of a longer one, replace
                        cand = self.extended_mapping
                        if firma_val and 'FIRMA' in cand.columns:
                            cand = cand[cand['FIRMA'] == firma_val]
                        suf = cand[cand['DEBI_KREDI'].str.endswith(debi_val)]
                        if not suf.empty and suf.iloc[0]['DEBI_KREDI'] != debi_val:
                            matched_row = suf.iloc[0]
                            row['DEBI_KREDI'] = matched_row['DEBI_KREDI']
                    
                    # After enriching DEBI_KREDI, also fill cost center fields from matched row
                    if matched_row is not None:
                        if 'Kostenträger' in matched_row:
                            kosttrager_from_ext = str(matched_row['Kostenträger'])
                            if kosttrager_from_ext and not kosttrager_from_ext.startswith('0'):
                                kosttrager_from_ext = '0' + kosttrager_from_ext
                            row['KOSTTRAGER'] = kosttrager_from_ext
                            
                            if len(kosttrager_from_ext) >= 4:
                                row['KOSTSTELLE'] = kosttrager_from_ext[:4]
                        
                        bez_col = 'Kostenträgerbezeichnung' if 'Kostenträgerbezeichnung' in matched_row else 'Kostenträger Bezeichnung'
                        if bez_col in matched_row:
                            row['Kostenträgerbezeichnung'] = str(matched_row[bez_col])
            except Exception:
                pass

            # Run validation if available
            if FieldValidator:
                validation_results = FieldValidator.validate_all_fields(
                    row
                )
                row['_validation'] = validation_results
                
                # Flag invalid fields
                has_invalid = any(
                    not v['valid']
                    for v in validation_results.values()
                )
                if has_invalid:
                    row['validation_required'] = True

            self.processed_data.append(row)

        # Update grid with color coding
        for item in self.tree.get_children():
            self.tree.delete(item)

        for row in self.processed_data:
            values = [row.get(col, '') for col in self.columns]
            item_id = self.tree.insert('', tk.END, values=values)
            
            # Tag rows needing validation
            if row.get('validation_required', False):
                self.tree.item(
                    item_id,
                    tags=('needs_validation',)
                )
        
        # Configure tags for visual feedback
        self.tree.tag_configure(
            'needs_validation',
            background='#fff3cd',
            foreground='#856404'
        )

    def upload_pdf_mapping(self):
        filename = filedialog.askopenfilename(title="Mapping-Datei", filetypes=[("Excel", "*.xlsx")])
        if not filename: return
        
        try:
            df = pd.read_excel(filename)
            column_map = {
                'Personenkonto': 'Kundennummer',
                'Kostt Hellern 2025': 'Kostenträger',
                'Kostenträger Bezeichnung': 'Kostenträgerbezeichnung'
            }
            df = df.rename(columns=column_map)
            
            required = ['Kundennummer', 'Kostenträger', 'Kostenträgerbezeichnung']
            missing = [col for col in required if col not in df.columns]
            
            if missing:
                messagebox.showerror("Fehler", f"Fehlende Spalten: {', '.join(missing)}")
                return
                
            df = df[required]
            df.to_excel(self.mapping_file_path, index=False)
            self.load_stored_mapping()
            messagebox.showinfo("Erfolg", "Datenbank aktualisiert!")
            
        except Exception as e:
            messagebox.showerror("Fehler", str(e))

    def load_stored_mapping(self):
        if self.mapping_file_path.exists():
            try:
                df = pd.read_excel(self.mapping_file_path)
                # Normalize columns to expected names
                col_map = {
                    'Kundennummer': 'DEBI_KREDI',
                    'Personenkonto': 'DEBI_KREDI',
                    'Kostenträgerbezeichnung': 'Kostenträger Bezeichnung',
                }
                df = df.rename(columns=col_map)
                # Ensure key columns exist
                # Support both old and new schema
                for needed in ['DEBI_KREDI', 'Kostenträger']:
                    if needed not in df.columns:
                        raise ValueError(f"Spalte fehlt: {needed}")
                # Clean data types
                df['DEBI_KREDI'] = (
                    df['DEBI_KREDI'].astype(str).str.replace(' ', '')
                )
                if 'FIRMA' in df.columns:
                    df['FIRMA'] = df['FIRMA'].astype(str).str.strip()
                if 'Kostenträger Bezeichnung' in df.columns:
                    df['Kostenträger Bezeichnung'] = (
                        df['Kostenträger Bezeichnung'].astype(str)
                    )
                self.mapping_data = df
                self.pdf_mapping_status.config(
                    text="✅ Datenbank aktiv",
                    fg=ModernStyle.SUCCESS
                )
            except Exception as e:
                self.pdf_mapping_status.config(
                    text=f"❌ Fehler: {e}",
                    fg='#f15e64'
                )
        else:
            self.pdf_mapping_status.config(
                text="Keine Datenbank",
                fg=ModernStyle.TEXT_SECONDARY
            )

    def load_extended_mapping(self):
        """Load extended mapping file (kostentreäger_info_3.xlsx) if present for DEBI_KREDI enrichment."""
        try:
            ext_path = Path('kostentreäger_info_3.xlsx')
            if not ext_path.exists():
                return
            df = pd.read_excel(ext_path)
            # Basic column normalization
            rename_map = {
                'Kostenträger Bezeichnung': 'Kostenträgerbezeichnung'
            }
            df = df.rename(columns=rename_map)
            needed_cols = ['FIRMA', 'DEBI_KREDI', 'Kostenträger']
            for c in needed_cols:
                if c not in df.columns:
                    return  # silently ignore if structure unexpected
            df['FIRMA'] = df['FIRMA'].astype(str).str.strip()
            df['DEBI_KREDI'] = df['DEBI_KREDI'].astype(str).str.strip()
            df['Kostenträger'] = df['Kostenträger'].astype(str).str.strip()
            if 'Kostenträgerbezeichnung' in df.columns:
                df['Kostenträgerbezeichnung'] = df['Kostenträgerbezeichnung'].astype(str)
            self.extended_mapping = df
        except Exception:
            # Non-fatal; extended mapping optional
            self.extended_mapping = None

    def apply_pdf_settings(self):
        if not self.raw_pdf_data:
            return
        
        settings = {k: v.get() for k, v in self.pdf_config_entries.items()}
        buch_text_template = self.pdf_buch_text_entry.get()
        
        self.processed_data = []
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for item in self.raw_pdf_data:
            beleg_dat = ""
            buch_jahr = ""
            buch_monat = ""
            if item.get('invoice_date'):
                try:
                    date_obj = pd.to_datetime(
                        item['invoice_date'], dayfirst=True
                    )
                    beleg_dat = date_obj.strftime('%Y%m%d')
                    buch_jahr = date_obj.year
                    buch_monat = date_obj.month
                except Exception:
                    pass
                
            betrag = int(item.get('amount', 0) * 100)
            
            buch_text = buch_text_template.format(
                student=item.get('student_name', ''),
                subject=item.get('subject', ''),
                school=item.get('school', ''),
                month=item.get('month_year', '')
            )
            
            kosttrager = settings['KOSTTRAGER']
            kost_bez = settings['Kostenträgerbezeichnung']

            # New mapping: filter by FIRMA, then find closest DEBI_KREDI match
            if self.mapping_data is not None:
                firma_input = str(settings.get('FIRMA', '')).strip()
                debi_kredi_extracted = str(
                    item.get('customer_number', '')
                ).replace(' ', '')
                df = self.mapping_data.copy()
                if firma_input and 'FIRMA' in df.columns:
                    df = df[df['FIRMA'] == firma_input]
                # Try exact match on DEBI_KREDI
                exact = df[df['DEBI_KREDI'] == debi_kredi_extracted]
                chosen = None
                if not exact.empty:
                    chosen = exact.iloc[0]
                else:
                    # If no exact, compute numeric distance for closest match
                    try:
                        target = int(
                            ''.join(
                                filter(
                                    str.isdigit,
                                    debi_kredi_extracted or '0'
                                )
                            )
                        )
                        df_num = df.assign(
                            _num=df['DEBI_KREDI'].apply(
                                lambda s: int(
                                    ''.join(
                                        filter(str.isdigit, str(s))
                                    ) or 0
                                )
                            )
                        )
                        if not df_num.empty:
                            df_num = df_num.assign(
                                _dist=(df_num['_num'] - target).abs()
                            )
                            chosen = df_num.sort_values('_dist').iloc[0]
                    except Exception:
                        pass
                if chosen is not None:
                    # Support both label variants for description
                    kosttrager = str(chosen['Kostenträger'])
                    bez_col = (
                        'Kostenträger Bezeichnung'
                        if 'Kostenträger Bezeichnung' in df.columns
                        else 'Kostenträgerbezeichnung'
                    )
                    kost_bez = str(chosen.get(bez_col, kost_bez))
            
            if kosttrager and not kosttrager.startswith('0'):
                kosttrager = '0' + kosttrager
                
            koststelle = (
                kosttrager[:4]
                if kosttrager and len(kosttrager) >= 4
                else settings['KOSTSTELLE']
            )
            
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

            # Enrich / correct DEBI_KREDI using extended mapping if missing or shortened
            try:
                if self.extended_mapping is not None:
                    debi_val = str(row.get('DEBI_KREDI', '')).strip()
                    firma_val = str(row.get('FIRMA', '')).strip()
                    matched_row = None
                    # Criteria: empty debi OR numeric too short (< 7) OR suffix match exists in extended mapping
                    if (not debi_val) or (len(debi_val) < 7):
                        cand = self.extended_mapping
                        if firma_val and 'FIRMA' in cand.columns:
                            cand = cand[cand['FIRMA'] == firma_val]
                        # Try exact Kostenträger match (remove leading zero if added)
                        raw_kost = str(row.get('KOSTTRAGER', '')).lstrip('0')
                        k_match = cand[cand['Kostenträger'].str.lstrip('0') == raw_kost]
                        if not k_match.empty:
                            matched_row = k_match.iloc[0]
                            row['DEBI_KREDI'] = matched_row['DEBI_KREDI']
                        else:
                            # Fallback: any DEBI_KREDI ending with current debi_val
                            if debi_val:
                                suf = cand[cand['DEBI_KREDI'].str.endswith(debi_val)]
                                if not suf.empty:
                                    matched_row = suf.iloc[0]
                                    row['DEBI_KREDI'] = matched_row['DEBI_KREDI']
                    else:
                        # If we have a short value that is suffix of a longer one, replace
                        cand = self.extended_mapping
                        if firma_val and 'FIRMA' in cand.columns:
                            cand = cand[cand['FIRMA'] == firma_val]
                        suf = cand[cand['DEBI_KREDI'].str.endswith(debi_val)]
                        if not suf.empty and suf.iloc[0]['DEBI_KREDI'] != debi_val:
                            matched_row = suf.iloc[0]
                            row['DEBI_KREDI'] = matched_row['DEBI_KREDI']
                    
                    # After enriching DEBI_KREDI, also fill cost center fields from matched row
                    if matched_row is not None:
                        if 'Kostenträger' in matched_row:
                            kosttrager_from_ext = str(matched_row['Kostenträger'])
                            if kosttrager_from_ext and not kosttrager_from_ext.startswith('0'):
                                kosttrager_from_ext = '0' + kosttrager_from_ext
                            row['KOSTTRAGER'] = kosttrager_from_ext
                            
                            if len(kosttrager_from_ext) >= 4:
                                row['KOSTSTELLE'] = kosttrager_from_ext[:4]
                        
                        bez_col = 'Kostenträgerbezeichnung' if 'Kostenträgerbezeichnung' in matched_row else 'Kostenträger Bezeichnung'
                        if bez_col in matched_row:
                            row['Kostenträgerbezeichnung'] = str(matched_row[bez_col])
            except Exception:
                pass
            
            self.processed_data.append(row)
            values = [row.get(col, '') for col in self.columns]
            self.tree.insert('', tk.END, values=values)

    def preview_pdf_matches(self):
        """Preview how many rows will get mapped with current FIRMA."""
        try:
            if self.mapping_data is None:
                messagebox.showerror("Fehler", "Keine Mapping-Datenbank geladen.")
                return
            firma = ''
            # Get FIRMA from combobox or entry
            firma_widget = self.pdf_config_entries.get('FIRMA')
            if isinstance(firma_widget, ttk.Combobox):
                firma = firma_widget.get().strip()
            else:
                firma = firma_widget.get().strip()
            df = self.mapping_data
            if 'FIRMA' in df.columns:
                df = df[df['FIRMA'].astype(str).str.strip() == str(firma)]
            total = len(self.raw_pdf_data or [])
            mapped = 0
            for item in self.raw_pdf_data or []:
                debi = str(item.get('customer_number', '')).replace(' ', '')
                exact = df[df['DEBI_KREDI'].astype(str) == debi]
                if not exact.empty:
                    mapped += 1
            messagebox.showinfo(
                "Vorschau",
                f"FIRMA {firma}: {mapped}/{total} Einträge mit exaktem Match."
            )
        except Exception as e:
            messagebox.showerror("Fehler", str(e))
    
    # Excel Transformer Methods
    def select_template(self):
        filename = filedialog.askopenfilename(
            title="Template wählen",
            filetypes=[("Excel", "*.xlsx")]
        )
        if filename:
            self.template_path = filename
            self.template_label.config(
                text=Path(filename).name,
                fg=ModernStyle.TEXT_PRIMARY
            )
            self._check_excel_ready()

    def select_source(self):
        filename = filedialog.askopenfilename(
            title="Quelldatei wählen",
            filetypes=[("Excel", "*.xlsx")]
        )
        if filename:
            self.source_path = filename
            self.source_label.config(
                text=Path(filename).name,
                fg=ModernStyle.TEXT_PRIMARY
            )
            self._check_excel_ready()

    def _check_excel_ready(self):
        if self.template_path and self.source_path:
            self.excel_transform_btn.config(state=tk.NORMAL)

    def transform_excel_data(self):
        if not self.template_path or not self.source_path:
            return
        
        try:
            defaults = {
                k: v.get() for k, v in self.excel_config_entries.items()
            }
            
            self.processed_data = bereitspf_transform(
                self.source_path,
                self.template_path,
                defaults=defaults
            )
            
            # Apply mapping database logic (same as PDF Reader)
            for row in self.processed_data:
                # Get customer/debitor number
                customer_num = row.get('DEBI_KREDI', '')
                
                # Default Kostenträger from settings
                kosttrager = row.get(
                    'KOSTTRAGER', defaults.get('KOSTTRAGER', '')
                )
                kost_bez = row.get('Kostenträgerbezeichnung', '')
                
                # New mapping: filter by FIRMA, then closest DEBI_KREDI
                if self.mapping_data is not None:
                    firma_input = str(defaults.get('FIRMA', '')).strip()
                    debi_kredi_extracted = str(customer_num).replace(' ', '')
                    df = self.mapping_data.copy()
                    if firma_input and 'FIRMA' in df.columns:
                        df = df[df['FIRMA'] == firma_input]
                    exact = df[df['DEBI_KREDI'] == debi_kredi_extracted]
                    chosen = None
                    if not exact.empty:
                        chosen = exact.iloc[0]
                    else:
                        try:
                            target = int(
                                ''.join(
                                    filter(
                                        str.isdigit,
                                        debi_kredi_extracted or '0'
                                    )
                                )
                            )
                            df_num = df.assign(
                                _num=df['DEBI_KREDI'].apply(
                                    lambda s: int(
                                        ''.join(
                                            filter(str.isdigit, str(s))
                                        ) or 0
                                    )
                                )
                            )
                            if not df_num.empty:
                                df_num = df_num.assign(
                                    _dist=(df_num['_num'] - target).abs()
                                )
                                chosen = df_num.sort_values('_dist').iloc[0]
                        except Exception:
                            pass
                    if chosen is not None:
                        kosttrager = str(chosen['Kostenträger'])
                        bez_col = (
                            'Kostenträger Bezeichnung'
                            if 'Kostenträger Bezeichnung' in df.columns
                            else 'Kostenträgerbezeichnung'
                        )
                        kost_bez = str(chosen.get(bez_col, kost_bez))
                
                # Logic: Ensure Kostenträger starts with 0
                if kosttrager and not str(kosttrager).startswith('0'):
                    kosttrager = '0' + str(kosttrager)
                    
                # Logic: Koststelle is first 4 digits of Kostenträger
                if kosttrager and len(str(kosttrager)) >= 4:
                    koststelle = str(kosttrager)[:4]
                else:
                    koststelle = row.get(
                        'KOSTSTELLE', defaults.get('KOSTSTELLE', '')
                    )
                    # Also ensure Koststelle starts with 0
                    if koststelle and not str(koststelle).startswith('0'):
                        koststelle = '0' + str(koststelle)
                
                # Update row
                row['KOSTTRAGER'] = kosttrager
                row['KOSTSTELLE'] = koststelle
                row['Kostenträgerbezeichnung'] = kost_bez

                # Enrich / correct DEBI_KREDI in transformer mode
                try:
                    if self.extended_mapping is not None:
                        debi_val = str(row.get('DEBI_KREDI', '')).strip()
                        firma_val = str(defaults.get('FIRMA', '')).strip()
                        matched_row = None
                        if (not debi_val) or (len(debi_val) < 7):
                            cand = self.extended_mapping
                            if firma_val and 'FIRMA' in cand.columns:
                                cand = cand[cand['FIRMA'] == firma_val]
                            raw_kost = str(row.get('KOSTTRAGER', '')).lstrip('0')
                            k_match = cand[cand['Kostenträger'].str.lstrip('0') == raw_kost]
                            if not k_match.empty:
                                matched_row = k_match.iloc[0]
                                row['DEBI_KREDI'] = matched_row['DEBI_KREDI']
                            elif debi_val:
                                suf = cand[cand['DEBI_KREDI'].str.endswith(debi_val)]
                                if not suf.empty:
                                    matched_row = suf.iloc[0]
                                    row['DEBI_KREDI'] = matched_row['DEBI_KREDI']
                        else:
                            cand = self.extended_mapping
                            if firma_val and 'FIRMA' in cand.columns:
                                cand = cand[cand['FIRMA'] == firma_val]
                            suf = cand[cand['DEBI_KREDI'].str.endswith(debi_val)]
                            if not suf.empty and suf.iloc[0]['DEBI_KREDI'] != debi_val:
                                matched_row = suf.iloc[0]
                                row['DEBI_KREDI'] = matched_row['DEBI_KREDI']
                        
                        # After enriching DEBI_KREDI, also fill cost center fields
                        if matched_row is not None:
                            if 'Kostenträger' in matched_row:
                                kosttrager_from_ext = str(matched_row['Kostenträger'])
                                if kosttrager_from_ext and not kosttrager_from_ext.startswith('0'):
                                    kosttrager_from_ext = '0' + kosttrager_from_ext
                                row['KOSTTRAGER'] = kosttrager_from_ext
                                
                                if len(kosttrager_from_ext) >= 4:
                                    row['KOSTSTELLE'] = kosttrager_from_ext[:4]
                            
                            bez_col = 'Kostenträgerbezeichnung' if 'Kostenträgerbezeichnung' in matched_row else 'Kostenträger Bezeichnung'
                            if bez_col in matched_row:
                                row['Kostenträgerbezeichnung'] = str(matched_row[bez_col])
                except Exception:
                    pass
            
            # Clear and repopulate grid
            for item in self.tree.get_children():
                self.tree.delete(item)
                
            for row in self.processed_data:
                values = [row.get(col, '') for col in self.columns]
                self.tree.insert('', tk.END, values=values)
            
            self.excel_export_btn.config(state=tk.NORMAL)
            messagebox.showinfo(
                "Erfolg",
                f"{len(self.processed_data)} Einträge transformiert!"
            )
            
        except Exception as e:
            messagebox.showerror("Fehler", str(e))
    
    # Shared Methods
    def on_tree_edit(self, event):
        self.processed_data = []
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id, "values")
            row = dict(zip(self.columns, values))
            self.processed_data.append(row)

    def export_excel(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not filename:
            return
        
        try:
            df = pd.DataFrame(self.processed_data)
            
            all_columns = [
                'SATZART', 'FIRMA', 'BELEG_NR', 'BELEG_DAT',
                'SOLL_HABEN', 'BUCH_KREIS',
                'BUCH_JAHR', 'BUCH_MONAT', 'DEBI_KREDI', 'BETRAG',
                'RECHNUNG', 'leer',
                'BUCH_TEXT', 'HABENKONTO', 'SOLLKONTO', 'leer_1', 'KOSTSTELLE',
                'KOSTTRAGER', 'Kostenträgerbezeichnung', 'Bebuchbar',
                'Debitoren.Bezeichnung',
                'Debitoren.Aktuelle Anschrift Anschrift-Zusatz',
                'AbgBenutzerdefiniert'
            ]
            
            for col in all_columns:
                if col not in df.columns:
                    df[col] = None
                    
            df = df[all_columns]
            df.to_excel(filename, index=False, engine='openpyxl')
            messagebox.showinfo(
                "Erfolg",
                f"Datei gespeichert:\n{Path(filename).name}"
            )
            
        except Exception as e:
            messagebox.showerror("Fehler", str(e))


def main():
    root = tk.Tk()
    app = MultiProjectApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
