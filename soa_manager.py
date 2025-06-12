# pip install pandas openpyxl fpdf
# Optionally: pip install ttkthemes for extra themes

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from fpdf import FPDF
import logging

try:
    from ttkthemes import ThemedTk
    THEMED = True
except ImportError:
    THEMED = False

# Setup logging
logging.basicConfig(
    filename='soa_manager.log',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s'
)

CONTROLS_2022 = [
    ("A.5.1", "Policies for information security"),
    ("A.5.2", "Information security roles and responsibilities"),
    ("A.5.3", "Segregation of duties"),
    ("A.5.4", "Management responsibilities"),
    ("A.5.5", "Contact with authorities"),
    ("A.5.6", "Contact with special interest groups"),
    ("A.5.7", "Threat intelligence"),
    ("A.5.8", "Information security in project management"),
    ("A.5.9", "Inventory of information and other associated assets"),
    ("A.5.10", "Acceptable use of information and other associated assets"),
    ("A.5.11", "Return of assets"),
    ("A.5.12", "Classification of information"),
    ("A.5.13", "Labeling of information"),
    ("A.5.14", "Information transfer"),
    ("A.5.15", "Access control"),
    ("A.5.16", "Identity management"),
    ("A.5.17", "Authentication information"),
    ("A.5.18", "Access rights"),
    ("A.5.19", "Information security in supplier relationships"),
    ("A.5.20", "Addressing information security within supplier agreements"),
    ("A.5.21", "Managing information security in the ICT supply chain"),
    ("A.5.22", "Monitoring, review and change management of supplier services"),
    ("A.5.23", "Information security for use of cloud services"),
    ("A.5.24", "Information security incident management planning and preparation"),
    ("A.5.25", "Assessment and decision on information security events"),
    ("A.5.26", "Response to information security incidents"),
    ("A.5.27", "Learning from information security incidents"),
    ("A.5.28", "Collection of evidence"),
    ("A.5.29", "Information security during disruption"),
    ("A.5.30", "ICT readiness for business continuity"),
    ("A.5.31", "Legal, statutory, regulatory and contractual requirements"),
    ("A.5.32", "Intellectual property rights"),
    ("A.5.33", "Protection of records"),
    ("A.5.34", "Privacy and protection of personally identifiable information (PII)"),
    ("A.5.35", "Independent review of information security"),
    ("A.6.1", "Screening"),
    ("A.6.2", "Terms and conditions of employment"),
    ("A.6.3", "Awareness, education and training"),
    ("A.6.4", "Disciplinary process"),
    ("A.6.5", "Responsibilities after termination or change of employment"),
    ("A.7.1", "Physical security perimeter"),
    ("A.7.2", "Physical entry"),
    ("A.7.3", "Securing offices, rooms and facilities"),
    ("A.7.4", "Physical security monitoring"),
    ("A.7.5", "Protection against physical and environmental threats"),
    ("A.7.6", "Working in secure areas"),
    ("A.7.7", "Clear desk and clear screen"),
    ("A.8.1", "User endpoint devices"),
    ("A.8.2", "Privileged access rights"),
    ("A.8.3", "Information access restriction"),
    ("A.8.4", "Access to source code"),
    ("A.8.5", "Secure authentication"),
    ("A.8.6", "Capacity management"),
    ("A.8.7", "Protection against malware"),
    ("A.8.8", "Management of technical vulnerabilities"),
    ("A.8.9", "Configuration management"),
    ("A.8.10", "Deletion of information"),
    ("A.8.11", "Data masking"),
    ("A.8.12", "Data leakage prevention"),
    ("A.8.13", "Information backup"),
    ("A.8.14", "Redundancy of information processing facilities"),
    ("A.8.15", "Logging"),
    ("A.8.16", "Monitoring activities"),
    ("A.8.17", "Clock synchronization"),
    ("A.8.18", "Use of privileged utility programs"),
    ("A.8.19", "Installation of software on operational systems"),
    ("A.8.20", "Networks security"),
    ("A.8.21", "Security of network services"),
    ("A.8.22", "Segregation of networks"),
    ("A.8.23", "Web filtering"),
    ("A.8.24", "Use of cryptography"),
    ("A.8.25", "Secure development lifecycle"),
    ("A.8.26", "Application security requirements"),
    ("A.8.27", "Secure system architecture and engineering principles"),
    ("A.8.28", "Secure coding"),
    ("A.8.29", "Security testing in development and acceptance"),
    ("A.8.30", "Outsourced development"),
    ("A.8.31", "Separation of development, testing and production environments"),
    ("A.8.32", "Change management"),
    ("A.8.33", "Test information"),
    ("A.8.34", "Protection of PII"),
    ("A.8.35", "Audit of information processing facilities"),
]

CONTROL_DICT = dict(CONTROLS_2022)
SOA_COLUMNS = [
    "Control ID", "Control Title", "Applicability", "Justification",
    "Implementation Status", "Responsible Party", "Evidence Location"
]

def create_tooltip(widget, text):
    tipwindow = None
    def enter(event):
        nonlocal tipwindow
        if tipwindow or not text:
            return
        x, y, cx, cy = widget.bbox("insert") if hasattr(widget, 'bbox') else (0, 0, 0, 0)
        x = x + widget.winfo_rootx() + 25
        y = y + widget.winfo_rooty() + 20
        tipwindow = tw = tk.Toplevel(widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw, text=text, justify=tk.LEFT, background="#ffffe0",
            relief=tk.SOLID, borderwidth=1, font=("tahoma", "8", "normal")
        )
        label.pack(ipadx=1)
    def leave(event):
        nonlocal tipwindow
        if tipwindow:
            tipwindow.destroy()
            tipwindow = None
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

class SoAApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ISO/IEC 27001:2022 SoA Manager")
        self.soa_df = pd.DataFrame(columns=SOA_COLUMNS)
        self._init_ui()
        self.refresh_table()

    def _init_ui(self):
        style = ttk.Style()
        if THEMED:
            self.root.set_theme('arc')
        else:
            style.theme_use('clam')
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'))
        style.configure("Treeview", font=('Segoe UI', 9))
        style.map("TButton", foreground=[('active', '#0053ba')])

        # Optional: Theme selector
        if THEMED:
            theme_frame = ttk.Frame(self.root)
            theme_frame.grid(row=0, column=0, sticky="ew")
            ttk.Label(theme_frame, text="Theme:").pack(side="left", padx=3)
            theme_cb = ttk.Combobox(
                theme_frame, values=self.root.get_themes(), width=20, state="readonly"
            )
            theme_cb.pack(side="left", padx=3)
            theme_cb.set(self.root.get_theme())
            def set_theme(event):
                self.root.set_theme(theme_cb.get())
            theme_cb.bind('<<ComboboxSelected>>', set_theme)

        big_frame = ttk.Frame(self.root, padding="10 10 10 10")
        big_frame.grid(row=1, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        big_frame.columnconfigure(1, weight=1)

        entry_frame = ttk.LabelFrame(big_frame, text="Add/Edit Control", padding="10 10 10 10")
        entry_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=10)

        ttk.Label(entry_frame, text="Control ID:").grid(row=0, column=0, sticky="e")
        self.control_id = ttk.Combobox(entry_frame, values=list(CONTROL_DICT.keys()), state="readonly", width=20)
        self.control_id.grid(row=0, column=1, sticky="ew", padx=5)
        self.control_id.bind("<<ComboboxSelected>>", self.autofill_title)
        create_tooltip(self.control_id, "Select a control from the list. Required.")

        ttk.Label(entry_frame, text="Control Title:").grid(row=0, column=2, sticky="e")
        self.control_title = ttk.Entry(entry_frame, width=50)
        self.control_title.grid(row=0, column=3, sticky="ew", padx=5)
        create_tooltip(self.control_title, "The title will auto-fill when you select a Control ID. Required.")

        ttk.Label(entry_frame, text="Applicable?").grid(row=1, column=0, sticky="e")
        self.applicability = ttk.Combobox(entry_frame, values=["Yes", "No"], state="readonly", width=20)
        self.applicability.grid(row=1, column=1, sticky="ew", padx=5)
        create_tooltip(self.applicability, "Is this control applicable? Choose Yes or No.")

        ttk.Label(entry_frame, text="Justification:").grid(row=1, column=2, sticky="e")
        self.justification = ttk.Entry(entry_frame, width=50)
        self.justification.grid(row=1, column=3, sticky="ew", padx=5)
        create_tooltip(self.justification, "Provide justification for applicability (required if 'No').")

        ttk.Label(entry_frame, text="Implementation Status:").grid(row=2, column=0, sticky="e")
        self.status = ttk.Combobox(entry_frame, values=["Implemented", "Planned", "Not Implemented"], state="readonly", width=20)
        self.status.grid(row=2, column=1, sticky="ew", padx=5)
        create_tooltip(self.status, "Implementation status of the control.")

        ttk.Label(entry_frame, text="Responsible Party:").grid(row=2, column=2, sticky="e")
        self.owner = ttk.Entry(entry_frame, width=20)
        self.owner.grid(row=2, column=3, sticky="w", padx=5)
        create_tooltip(self.owner, "Person or team responsible.")

        ttk.Label(entry_frame, text="Evidence Location:").grid(row=3, column=0, sticky="e")
        self.evidence = ttk.Entry(entry_frame, width=70)
        self.evidence.grid(row=3, column=1, columnspan=3, sticky="ew", padx=5)
        create_tooltip(self.evidence, "Where to find supporting evidence for this control.")

        btn_frame = ttk.Frame(big_frame)
        btn_frame.grid(row=1, column=0, sticky='w', padx=5, pady=5)
        ttk.Button(btn_frame, text="Add Control", command=self.add_control).grid(row=0, column=0, padx=2)
        ttk.Button(btn_frame, text="Export CSV", command=lambda: self.export_file('csv')).grid(row=0, column=1, padx=2)
        ttk.Button(btn_frame, text="Import CSV", command=lambda: self.import_file('csv')).grid(row=0, column=2, padx=2)
        ttk.Button(btn_frame, text="Export Excel", command=lambda: self.export_file('xlsx')).grid(row=0, column=3, padx=2)
        ttk.Button(btn_frame, text="Import Excel", command=lambda: self.import_file('xlsx')).grid(row=0, column=4, padx=2)
        ttk.Button(btn_frame, text="Export PDF", command=self.export_pdf).grid(row=0, column=5, padx=2)

        tree_frame = ttk.Frame(big_frame)
        tree_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=8)
        big_frame.rowconfigure(2, weight=1)
        big_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_frame, columns=SOA_COLUMNS, show="headings", height=14)
        for col in SOA_COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor="center")

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

    def autofill_title(self, event=None):
        selected = self.control_id.get()
        self.control_title.delete(0, tk.END)
        self.control_title.insert(0, CONTROL_DICT.get(selected, ""))

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        for idx, (_, row) in enumerate(self.soa_df.iterrows()):
            tags = ('oddrow',) if idx % 2 else ('evenrow',)
            self.tree.insert("", "end", values=list(row), tags=tags)
        self.tree.tag_configure('oddrow', background='#f6f6fc')
        self.tree.tag_configure('evenrow', background='#e9f5fd')

    def validate_entry(self, entry):
        if not entry["Control ID"]:
            return False, "Control ID is required."
        if not entry["Control Title"]:
            return False, "Control Title is required."
        if entry["Applicability"] == "No" and not entry["Justification"]:
            return False, "Justification is required if Applicability is 'No'."
        return True, ""

    def add_control(self):
        entry = {
            "Control ID": self.control_id.get(),
            "Control Title": self.control_title.get(),
            "Applicability": self.applicability.get(),
            "Justification": self.justification.get(),
            "Implementation Status": self.status.get(),
            "Responsible Party": self.owner.get(),
            "Evidence Location": self.evidence.get()
        }
        valid, msg = self.validate_entry(entry)
        if not valid:
            messagebox.showwarning("Missing Data", msg)
            logging.warning(f"Add Control failed: {msg}")
            return
        self.soa_df.loc[len(self.soa_df)] = entry
        self.refresh_table()
        messagebox.showinfo("Success", "Control added.")
        logging.info(f"Control added: {entry['Control ID']}")

    def export_file(self, filetype):
        filetypes = {
            'csv': (("CSV files", "*.csv"), ".csv"),
            'xlsx': (("Excel files", "*.xlsx"), ".xlsx")
        }
        label, ext = filetypes[filetype]
        path = filedialog.asksaveasfilename(defaultextension=ext, filetypes=[label])
        if path:
            try:
                if filetype == 'csv':
                    self.soa_df.to_csv(path, index=False)
                elif filetype == 'xlsx':
                    self.soa_df.to_excel(path, index=False, engine='openpyxl')
                messagebox.showinfo("Success", f"{filetype.upper()} saved to {path}")
                logging.info(f"Exported {filetype.upper()} to {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export {filetype.upper()}: {e}")
                logging.error(f"Export {filetype.upper()} failed: {e}")

    def import_file(self, filetype):
        filetypes = {
            'csv': (("CSV files", "*.csv"), pd.read_csv),
            'xlsx': (("Excel files", "*.xlsx"), lambda p: pd.read_excel(p, engine='openpyxl'))
        }
        label, loader = filetypes[filetype]
        path = filedialog.askopenfilename(filetypes=[label])
        if path:
            try:
                df = loader(path)
                if set(SOA_COLUMNS).issubset(df.columns):
                    self.soa_df = df[SOA_COLUMNS]
                    self.refresh_table()
                    messagebox.showinfo("Success", f"{filetype.upper()} imported.")
                    logging.info(f"Imported {filetype.upper()} from {path}")
                else:
                    messagebox.showerror("Error", f"{filetype.upper()} columns do not match expected format.")
                    logging.warning(f"Import {filetype.upper()} failed: columns mismatch")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import {filetype.upper()}: {e}")
                logging.error(f"Import {filetype.upper()} failed: {e}")

    def export_pdf(self):
        class SoAPDF(FPDF):
            def header(self):
                self.set_font("Arial", "B", 12)
                self.cell(0, 10, "ISO 27001:2022 Statement of Applicability", ln=1, align="C")
                self.ln(2)

            def soa_table(self, df):
                self.set_font("Arial", "B", 8)
                headers = df.columns.tolist()
                col_widths = []
                for i, header in enumerate(headers):
                    max_len = max([len(str(header))] + [len(str(row[header])) for _, row in df.iterrows()])
                    col_widths.append(min(max(20, max_len * 2.5), 45))
                for i, header in enumerate(headers):
                    self.cell(col_widths[i], 6, header[:20], border=1, align="C")
                self.ln()
                self.set_font("Arial", "", 8)
                for _, row in df.iterrows():
                    for i, header in enumerate(headers):
                        text = str(row[header])[:40]
                        self.cell(col_widths[i], 6, text, border=1)
                    self.ln()
                    if self.get_y() > 260:
                        self.add_page()
                        self.soa_table_headers(headers, col_widths)

            def soa_table_headers(self, headers, col_widths):
                self.set_font("Arial", "B", 8)
                for i, header in enumerate(headers):
                    self.cell(col_widths[i], 6, header[:20], border=1, align="C")
                self.ln()
                self.set_font("Arial", "", 8)

        pdf = SoAPDF()
        pdf.add_page()
        pdf.soa_table(self.soa_df)
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path:
            try:
                pdf.output(path)
                messagebox.showinfo("Success", f"PDF saved to {path}")
                logging.info(f"Exported PDF to {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export PDF: {e}")
                logging.error(f"Export PDF failed: {e}")

if __name__ == "__main__":
    if THEMED:
        root = ThemedTk(theme="arc")
    else:
        root = tk.Tk()
    app = SoAApp(root)
    root.mainloop()
