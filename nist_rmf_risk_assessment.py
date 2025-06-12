import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import pandas as pd
import re

def calculate_risk(impact, likelihood):
    score_map = {"Low": 1, "Moderate": 2, "High": 3}
    try:
        return score_map[impact] * score_map[likelihood]
    except KeyError:
        return 0

class RiskAssessmentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NIST RMF Risk Assessment Tool")

        self.system_info = {}
        self.threats = []

        self.create_system_info_frame()

    def create_system_info_frame(self):
        self.clear_frame()
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="x")

        self.system_name_var = tk.StringVar()
        self.system_desc_var = tk.StringVar()
        self.sensitivity_var = tk.StringVar()
        self.criticality_var = tk.StringVar()

        ttk.Label(frame, text="System Name:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.system_name_var).grid(row=0, column=1, sticky="ew")
        ttk.Label(frame, text="Description:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.system_desc_var).grid(row=1, column=1, sticky="ew")
        ttk.Label(frame, text="Sensitivity (Low/Moderate/High):").grid(row=2, column=0, sticky="w")
        ttk.Combobox(frame, textvariable=self.sensitivity_var, values=["Low","Moderate","High"]).grid(row=2, column=1, sticky="ew")
        ttk.Label(frame, text="Criticality (Low/Moderate/High):").grid(row=3, column=0, sticky="w")
        ttk.Combobox(frame, textvariable=self.criticality_var, values=["Low","Moderate","High"]).grid(row=3, column=1, sticky="ew")
        ttk.Button(frame, text="Next", command=self.save_system_info).grid(row=4, column=0, columnspan=2, pady=10)

    def save_system_info(self):
        if not all([self.system_name_var.get(), self.system_desc_var.get(),
                    self.sensitivity_var.get(), self.criticality_var.get()]):
            messagebox.showerror("Error", "All fields are required.")
            return
        self.system_info = {
            "System Name": self.system_name_var.get(),
            "Description": self.system_desc_var.get(),
            "Sensitivity": self.sensitivity_var.get(),
            "Criticality": self.criticality_var.get()
        }
        self.create_threats_frame()

    def create_threats_frame(self):
        self.clear_frame()
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="x")
        ttk.Label(frame, text="Add Threats and Vulnerabilities", font=("Arial", 12, "bold")).pack(pady=5)

        add_btn = ttk.Button(frame, text="Add Threat", command=self.add_threat_dialog)
        add_btn.pack()
        self.threats_tree = ttk.Treeview(frame, columns=("threat", "vuln", "impact", "likelihood", "mitigation", "score"), show="headings")
        for col in self.threats_tree["columns"]:
            self.threats_tree.heading(col, text=col.capitalize())
        self.threats_tree.pack(pady=10, fill="x")

        next_btn = ttk.Button(frame, text="Next", command=self.create_register_frame)
        next_btn.pack()

    def add_threat_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Threat")
        dialog.grab_set()
        threat_var = tk.StringVar()
        vuln_var = tk.StringVar()
        impact_var = tk.StringVar()
        likelihood_var = tk.StringVar()
        mitigation_var = tk.StringVar()
        frame = ttk.Frame(dialog, padding=10)
        frame.pack()

        ttk.Label(frame, text="Threat:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=threat_var).grid(row=0, column=1)
        ttk.Label(frame, text="Associated Vulnerability:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=vuln_var).grid(row=1, column=1)
        ttk.Label(frame, text="Impact (Low/Moderate/High):").grid(row=2, column=0, sticky="w")
        ttk.Combobox(frame, textvariable=impact_var, values=["Low","Moderate","High"]).grid(row=2, column=1)
        ttk.Label(frame, text="Likelihood (Low/Moderate/High):").grid(row=3, column=0, sticky="w")
        ttk.Combobox(frame, textvariable=likelihood_var, values=["Low","Moderate","High"]).grid(row=3, column=1)
        ttk.Label(frame, text="Mitigation:").grid(row=4, column=0, sticky="w")
        ttk.Entry(frame, textvariable=mitigation_var).grid(row=4, column=1)

        def save_and_close():
            if not all([threat_var.get(), vuln_var.get(), impact_var.get(), likelihood_var.get(), mitigation_var.get()]):
                messagebox.showerror("Error", "All fields are required.", parent=dialog)
                return
            risk_score = calculate_risk(impact_var.get(), likelihood_var.get())
            entry = {
                "Threat": threat_var.get(),
                "Vulnerability": vuln_var.get(),
                "Impact": impact_var.get(),
                "Likelihood": likelihood_var.get(),
                "Mitigation": mitigation_var.get(),
                "Risk Score": risk_score
            }
            self.threats.append(entry)
            self.threats_tree.insert("", "end", values=(entry["Threat"], entry["Vulnerability"], entry["Impact"], entry["Likelihood"], entry["Mitigation"], entry["Risk Score"]))
            dialog.destroy()

        ttk.Button(frame, text="Save", command=save_and_close).grid(row=5, column=0, columnspan=2, pady=5)

    def create_register_frame(self):
        if not self.threats:
            messagebox.showerror("Error", "Add at least one threat.")
            return
        self.clear_frame()
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="x")
        ttk.Label(frame, text="Risk Register Summary", font=("Arial", 12, "bold")).pack(pady=5)

        df = pd.DataFrame(self.threats)
        df["Risk Level"] = df["Risk Score"].apply(lambda x: "Low" if x <= 2 else "Moderate" if x <= 4 else "High")
        self.risk_df = df

        self.register_tree = ttk.Treeview(frame, columns=("Threat", "Risk Level", "Mitigation"), show="headings")
        for col in ("Threat", "Risk Level", "Mitigation"):
            self.register_tree.heading(col, text=col)
        for _, row in df.iterrows():
            self.register_tree.insert("", "end", values=(row["Threat"], row["Risk Level"], row["Mitigation"]))
        self.register_tree.pack(pady=10, fill="x")

        ttk.Button(frame, text="Export to Excel", command=self.export_results).pack(pady=5)

    def export_results(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"risk_assessment_{re.sub(r'[^\w\s]', '', self.system_info['System Name']).replace(' ', '_')}.xlsx"
        )
        if filename:
            with pd.ExcelWriter(filename) as writer:
                pd.DataFrame([self.system_info]).to_excel(writer, sheet_name="System Info", index=False)
                self.risk_df.to_excel(writer, sheet_name="Risk Register", index=False)
            messagebox.showinfo("Exported", f"Risk assessment exported to: {filename}")

    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = RiskAssessmentApp(root)
    root.mainloop()
