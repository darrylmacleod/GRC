# pip install pandas openpyxl fpdf

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from fpdf import FPDF

# Full control mapping
controls_2022 = [
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
    ("A.8.35", "Audit of information processing facilities")
]

# Dictionary for auto-fill
control_dict = dict(controls_2022)
soa_df = pd.DataFrame(columns=[
    "Control ID", "Control Title", "Applicability", "Justification",
    "Implementation Status", "Responsible Party", "Evidence Location"
])

class SoAApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ISO/IEC 27001:2022 SoA Manager")

        self.control_id = ttk.Combobox(root, values=list(control_dict.keys()))
        self.control_id.grid(row=0, column=1, padx=5, pady=2)
        tk.Label(root, text="Control ID:").grid(row=0, column=0)
        self.control_id.bind("<<ComboboxSelected>>", self.autofill_title)

        self.control_title = tk.Entry(root, width=60)
        self.control_title.grid(row=1, column=1, padx=5, pady=2)
        tk.Label(root, text="Control Title:").grid(row=1, column=0)

        self.applicability = ttk.Combobox(root, values=["Yes", "No"])
        self.applicability.grid(row=2, column=1, padx=5, pady=2)
        tk.Label(root, text="Applicable?").grid(row=2, column=0)

        self.justification = tk.Entry(root, width=60)
        self.justification.grid(row=3, column=1, padx=5, pady=2)
        tk.Label(root, text="Justification:").grid(row=3, column=0)

        self.status = ttk.Combobox(root, values=["Implemented", "Planned", "Not Implemented"])
        self.status.grid(row=4, column=1, padx=5, pady=2)
        tk.Label(root, text="Implementation Status:").grid(row=4, column=0)

        self.owner = tk.Entry(root, width=30)
        self.owner.grid(row=5, column=1, padx=5, pady=2)
        tk.Label(root, text="Responsible Party:").grid(row=5, column=0)

        self.evidence = tk.Entry(root, width=60)
        self.evidence.grid(row=6, column=1, padx=5, pady=2)
        tk.Label(root, text="Evidence Location:").grid(row=6, column=0)

        tk.Button(root, text="Add Control", command=self.add_control).grid(row=7, column=0, pady=6)
        tk.Button(root, text="Export to CSV", command=self.export_csv).grid(row=7, column=1, sticky="w")
        tk.Button(root, text="Import from CSV", command=self.import_csv).grid(row=7, column=1, sticky="e")
        tk.Button(root, text="Export to Excel", command=self.export_excel).grid(row=8, column=0)
        tk.Button(root, text="Import from Excel", command=self.import_excel).grid(row=8, column=1, sticky="w")
        tk.Button(root, text="Export to PDF", command=self.export_pdf).grid(row=8, column=1, sticky="e")

        self.tree = ttk.Treeview(root, columns=list(soa_df.columns), show="headings", height=12)
        for col in soa_df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130)
        self.tree.grid(row=9, column=0, columnspan=2, padx=10, pady=10)
        scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=9, column=2, sticky="ns")

    def autofill_title(self, event):
        selected = self.control_id.get()
        self.control_title.delete(0, tk.END)
        self.control_title.insert(0, control_dict.get(selected, ""))

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        for _, row in soa_df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def add_control(self):
        global soa_df
        entry = {
            "Control ID": self.control_id.get(),
            "Control Title": self.control_title.get(),
            "Applicability": self.applicability.get(),
            "Justification": self.justification.get(),
            "Implementation Status": self.status.get(),
            "Responsible Party": self.owner.get(),
            "Evidence Location": self.evidence.get()
        }
        soa_df.loc[len(soa_df)] = entry
        self.refresh_table()
        messagebox.showinfo("Success", "Control added.")

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv")
        if path:
            soa_df.to_csv(path, index=False)

    def import_csv(self):
        global soa_df
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if path:
            soa_df = pd.read_csv(path)
            self.refresh_table()

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if path:
            soa_df.to_excel(path, index=False, engine='openpyxl')

    def import_excel(self):
        global soa_df
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            soa_df = pd.read_excel(path, engine='openpyxl')
            self.refresh_table()

    def export_pdf(self):
        class SoAPDF(FPDF):
            def header(self):
                self.set_font("Arial", "B", 12)
                self.cell(0, 10, "ISO 27001:2022 Statement of Applicability", ln=1, align="C")
                self.ln(2)
            def soa_table(self, df):
                self.set_font("Arial", "", 8)
                headers = df.columns.tolist()
                col_widths = [20, 40, 15, 30, 25, 25, 40]
                for i, header in enumerate(headers):
                    self.cell(col_widths[i], 6, header[:15], border=1)
                self.ln()
                for _, row in df.iterrows():
                    for i, header in enumerate(headers):
                        self.cell(col_widths[i], 6, str(row[header])[:30], border=1)
                    self.ln()
        pdf = SoAPDF()
        pdf.add_page()
        pdf.soa_table(soa_df)
        path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if path:
            pdf.output(path)

if __name__ == "__main__":
    root = tk.Tk()
    app = SoAApp(root)
    root.mainloop()
