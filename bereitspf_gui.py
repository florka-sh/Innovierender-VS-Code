import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
import threading
from importlib.machinery import SourceFileLoader
import sys


class TransformGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Transform Tool")
        self.root.geometry("700x800")
        self.root.resizable(True, True)
        
        # Load the transform module
        script_path = Path(__file__).parent / "scripts" / "transform_excel.py"
        if not script_path.exists():
            messagebox.showerror("Error", f"transform_excel.py not found at {script_path}")
            sys.exit(1)
        
        loader = SourceFileLoader('transform_excel', str(script_path))
        self.transform_module = loader.load_module()
        
        # Default paths - use current working directory, not script location
        # This allows .exe to work from any directory
        self.cwd = Path.cwd()
        
        # GUI Elements
        self._build_ui()
    
    def _build_ui(self):
        # Title
        title = tk.Label(self.root, text="Excel Transform Tool", font=("Arial", 16, "bold"))
        title.pack(pady=10)
        
        # File Selection Frame
        file_frame = tk.LabelFrame(self.root, text="File Selection", padx=10, pady=10)
        file_frame.pack(fill=tk.BOTH, padx=10, pady=5)
        
        # Template File
        tk.Label(file_frame, text="Template File:").grid(row=0, column=0, sticky=tk.W)
        self.template_var = tk.StringVar(value="9241_1025_Bereitschatspflege_KRED.xlsx")
        tk.Entry(file_frame, textvariable=self.template_var, width=50).grid(row=0, column=1, padx=5)
        tk.Button(file_frame, text="Browse", command=self._browse_template).grid(row=0, column=2, padx=5)
        
        # Source File
        tk.Label(file_frame, text="Source File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.source_var = tk.StringVar(value="Auszahlungsbelege Pflegefamilien 11.2025.xlsx")
        tk.Entry(file_frame, textvariable=self.source_var, width=50).grid(row=1, column=1, padx=5)
        tk.Button(file_frame, text="Browse", command=self._browse_source).grid(row=1, column=2, padx=5)
        
        # Output File
        tk.Label(file_frame, text="Output Filename:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.output_var = tk.StringVar(value="output.xlsx")
        tk.Entry(file_frame, textvariable=self.output_var, width=50).grid(row=2, column=1, padx=5)
        
        # Parameters Frame
        param_frame = tk.LabelFrame(self.root, text="Parameters", padx=10, pady=10)
        param_frame.pack(fill=tk.BOTH, padx=10, pady=5)
        
        # Create input fields
        self.fields = {}
        params = [
            ('SATZART', 'K'),
            ('FIRMA', '9241'),
            ('SOLL_HABEN', 'S'),
            ('BUCH_KREIS', 'RE'),
            ('BUCH_JAHR', '2025'),
            ('BUCH_MONAT', '11'),
        ]
        
        for idx, (label, default) in enumerate(params):
            tk.Label(param_frame, text=f"{label}:").grid(row=idx, column=0, sticky=tk.W, pady=5)
            var = tk.StringVar(value=default)
            self.fields[label] = var
            tk.Entry(param_frame, textvariable=var, width=30).grid(row=idx, column=1, padx=5, sticky=tk.W)
        
        # Output Log Frame
        log_frame = tk.LabelFrame(self.root, text="Log", padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, width=80, state=tk.NORMAL)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Button Frame
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.run_btn = tk.Button(button_frame, text="Run Transform", command=self._run_transform, 
                                  bg="green", fg="white", font=("Arial", 12, "bold"))
        self.run_btn.pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame, text="Clear Log", command=self._clear_log).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Exit", command=self.root.quit).pack(side=tk.RIGHT, padx=5)
    
    def _log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def _clear_log(self):
        self.log_text.delete(1.0, tk.END)
    
    def _browse_template(self):
        fname = filedialog.askopenfilename(
            initialdir=self.cwd,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Select Template File"
        )
        if fname:
            self.template_var.set(fname)
    
    def _browse_source(self):
        fname = filedialog.askopenfilename(
            initialdir=self.cwd,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Select Source File"
        )
        if fname:
            self.source_var.set(fname)
    
    def _run_transform(self):
        self._clear_log()
        self.run_btn.config(state=tk.DISABLED)
        
        # Run in background thread to avoid freezing UI
        thread = threading.Thread(target=self._do_transform)
        thread.daemon = True
        thread.start()
    
    def _do_transform(self):
        try:
            self._log("Starting transformation...")
            
            # Validate paths
            template_path = self.cwd / self.template_var.get()
            source_path = self.cwd / self.source_var.get()
            output_filename = self.output_var.get()
            
            if not output_filename.endswith('.xlsx'):
                output_filename += '.xlsx'
            
            output_path = self.cwd / output_filename
            
            # Validate files exist
            if not template_path.exists():
                self._log(f"❌ Template file not found: {template_path}")
                self.run_btn.config(state=tk.NORMAL)
                return
            
            if not source_path.exists():
                self._log(f"❌ Source file not found: {source_path}")
                self.run_btn.config(state=tk.NORMAL)
                return
            
            self._log(f"✓ Template: {template_path}")
            self._log(f"✓ Source: {source_path}")
            self._log(f"✓ Output: {output_path}")
            
            # Prepare defaults
            defaults = {
                'SATZART': self.fields['SATZART'].get() or None,
                'FIRMA': self.fields['FIRMA'].get() or None,
                'SOLL_HABEN': self.fields['SOLL_HABEN'].get() or None,
                'BUCH_KREIS': self.fields['BUCH_KREIS'].get() or None,
                'BUCH_JAHR': self.fields['BUCH_JAHR'].get() or None,
                'BUCH_MONAT': self.fields['BUCH_MONAT'].get() or None,
                'NO_RENAME': True,  # GUI handles rename, so skip interactive prompt
            }
            
            self._log("\nParameters:")
            for k, v in defaults.items():
                if k != 'NO_RENAME':
                    self._log(f"  {k}: {v}")
            
            self._log("\nRunning transform...")
            
            # Call transform
            self.transform_module.transform(template_path, source_path, output_path, config_path=None, defaults=defaults)
            
            self._log(f"\n✓ Transformation complete!")
            self._log(f"Output saved to: {output_path}")
            
            # Ask to open file
            if messagebox.askyesno("Success", f"Transform complete!\n\nOpen output file?"):
                import subprocess
                subprocess.Popen(['start', str(output_path)], shell=True)
            
        except Exception as e:
            self._log(f"\n❌ Error: {str(e)}")
            import traceback
            self._log(traceback.format_exc())
        
        finally:
            self.run_btn.config(state=tk.NORMAL)


if __name__ == '__main__':
    root = tk.Tk()
    app = TransformGUI(root)
    root.mainloop()
