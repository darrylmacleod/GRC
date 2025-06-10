# Install Required Libraries
# pip install pandas openpyxl tkcalendar matplotlib

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from datetime import datetime, date

# === Constants ===
RISK_LEVEL_THRESHOLDS = {
    'Critical': 15,
    'High': 10,
    'Medium': 5
}

RISK_LEVEL_ORDER = ['Low', 'Medium', 'High', 'Critical']

EXCEL_COLUMNS = [
    "Risk ID", "Risk Description", "Impact", "Likelihood", "Risk Score", "Risk Level",
    "Risk Owner", "Due Date", "Notes"
]

# === Risk Data Model ===
class RiskRegisterModel:
    def __init__(self):
        self.risks = []
        self.next_id = 1

    def add_risk(self, risk):
        risk["Risk ID"] = self.next_id
        self.risks.append(risk)
        self.next_id += 1

    def remove_risk(self, risk_id):
        self.risks = [r for r in self.risks if r["Risk ID"] != risk_id]

    def update_risk(self, risk_id, new_data):
        for idx, r in enumerate(self.risks):
            if r["Risk ID"] == risk_id:
                self.risks[idx].update(new_data)
                break

    def to_dataframe(self):
        return pd.DataFrame(self.risks, columns=EXCEL_COLUMNS)

    def clear(self):
        self.risks.clear()
        self.next_id = 1

    def load_from_excel(self, filename):
        df = pd.read_excel(filename)
        self.risks = df.to_dict(orient='records')
        if self.risks:
            self.next_id = max(r["Risk ID"] for r in self.risks) + 1
        else:
            self.next_id = 1

    def save_to_excel(self, filename):
        df = self.to_dataframe()
        df.to_excel(filename, index=False)

# === Risk Calculation Functions ===
def calculate_risk_score(impact, likelihood):
    return impact * likelihood

def risk_level(score):
    if score >= RISK_LEVEL_THRESHOLDS['Critical']:
        return 'Critical'
    elif score >= RISK_LEVEL_THRESHOLDS['High']:
        return 'High'
    elif score >= RISK_LEVEL_THRESHOLDS['Medium']:
        return 'Medium'
    else:
        return 'Low'

# === Main Application Class ===
class RiskRegisterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automated Risk Register")

        self.model = RiskRegisterModel()
        self.selected_risk_id = None

        # === Input Fields ===
        tk.Label(root, text="Risk Description").grid(row=0, column=0)
        self.desc_entry = tk.Entry(root, width=50)
        self.desc_entry.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(root, text="Impact (1–5)").grid(row=1, column=0)
        self.impact_var = tk.StringVar()
        self.impact_menu = ttk.Combobox(root, textvariable=self.impact_var, values=[1, 2, 3, 4, 5], state="readonly", width=10)
        self.impact_menu.grid(row=1, column=1, sticky='w')

        tk.Label(root, text="Likelihood (1–5)").grid(row=2, column=0)
        self.likelihood_var = tk.StringVar()
        self.likelihood_menu = ttk.Combobox(root, textvariable=self.likelihood_var, values=[1, 2, 3, 4, 5], state="readonly", width=10)
        self.likelihood_menu.grid(row=2, column=1, sticky='w')

        tk.Label(root, text="Risk Owner").grid(row=3, column=0)
        self.owner_entry = tk.Entry(root, width=30)
        self.owner_entry.grid(row=3, column=1, sticky='w')

        tk.Label(root, text="Due Date").grid(row=4, column=0)
        self.due_date = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.due_date.grid(row=4, column=1, sticky='w')

        tk.Label(root, text="Notes").grid(row=5, column=0)
        self.notes_entry = tk.Entry(root, width=50)
        self.notes_entry.grid(row=5, column=1, padx=10, pady=5)

        # === Buttons ===
        tk.Button(root, text="Add/Update Risk", command=self.add_or_update_risk).grid(row=6, column=0, pady=10)
        tk.Button(root, text="Export to Excel", command=self.export_to_excel).grid(row=6, column=1, sticky='w')
        tk.Button(root, text="Show Risk Chart", command=self.show_risk_chart).grid(row=6, column=1, sticky='e')
        tk.Button(root, text="Clear Form", command=self.clear_inputs).grid(row=6, column=1)
        tk.Button(root, text="Delete Selected Risk", command=self.delete_selected_risk).grid(row=6, column=0, sticky='e')
        tk.Button(root, text="Load Risks from Excel", command=self.load_from_excel).grid(row=6, column=1, sticky='s')

        # === Treeview Table with Scrollbar ===
        columns = EXCEL_COLUMNS
        self.tree = ttk.Treeview(root, columns=columns, show='headings', height=10)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110)
        self.tree.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

        scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=7, column=2, sticky='ns')

        # Enable row selection for edit/delete
        self.tree.bind('<ButtonRelease-1>', self.on_tree_select)

        # === Chart Placeholder ===
        self.chart_canvas = None

    def validate_inputs(self):
        desc = self.desc_entry.get().strip()
        owner = self.owner_entry.get().strip()
        notes = self.notes_entry.get().strip()
        try:
            impact = int(self.impact_var.get())
            likelihood = int(self.likelihood_var.get())
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Please select Impact and Likelihood (1–5).")
            return None
        if not desc or not owner:
            messagebox.showerror("Error", "Please fill in all required fields (Description, Owner).")
            return None
        due = self.due_date.get_date()
        return desc, impact, likelihood, owner, due, notes

    def add_or_update_risk(self):
        validated = self.validate_inputs()
        if not validated:
            return
        desc, impact, likelihood, owner, due, notes = validated
        score = calculate_risk_score(impact, likelihood)
        level = risk_level(score)
        risk = {
            "Risk Description": desc,
            "Impact": impact,
            "Likelihood": likelihood,
            "Risk Score": score,
            "Risk Level": level,
            "Risk Owner": owner,
            "Due Date": due,
            "Notes": notes
        }
        if self.selected_risk_id:  # Update
            risk["Risk ID"] = self.selected_risk_id
            self.model.update_risk(self.selected_risk_id, risk)
            self.refresh_treeview()
            self.selected_risk_id = None
        else:  # Add new
            self.model.add_risk(risk)
            self.tree.insert("", "end", values=(
                risk["Risk ID"], desc, impact, likelihood, score, level, owner, due.strftime('%Y-%m-%d'), notes
            ))
        self.clear_inputs()

    def refresh_treeview(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for risk in self.model.risks:
            self.tree.insert("", "end", values=(
                risk["Risk ID"], risk["Risk Description"], risk["Impact"], risk["Likelihood"],
                risk["Risk Score"], risk["Risk Level"], risk["Risk Owner"],
                risk["Due Date"].strftime('%Y-%m-%d') if isinstance(risk["Due Date"], (datetime, date)) else str(risk["Due Date"]),
                risk["Notes"]
            ))

    def clear_inputs(self):
        self.desc_entry.delete(0, tk.END)
        self.owner_entry.delete(0, tk.END)
        self.notes_entry.delete(0, tk.END)
        self.impact_var.set('')
        self.likelihood_var.set('')
        self.selected_risk_id = None
        self.tree.selection_remove(self.tree.selection())

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if selected:
            values = self.tree.item(selected[0], 'values')
            self.selected_risk_id = int(values[0])
            # Fill entries for editing
            self.desc_entry.delete(0, tk.END)
            self.desc_entry.insert(0, values[1])
            self.impact_var.set(values[2])
            self.likelihood_var.set(values[3])
            self.owner_entry.delete(0, tk.END)
            self.owner_entry.insert(0, values[6])
            # Convert string to date for DateEntry, handle errors
            try:
                date_obj = datetime.strptime(values[7], "%Y-%m-%d").date()
                self.due_date.set_date(date_obj)
            except Exception as e:
                messagebox.showerror("Date Error", f"Could not parse date: {values[7]}\n{e}")
            self.notes_entry.delete(0, tk.END)
            self.notes_entry.insert(0, values[8])

    def delete_selected_risk(self):
        if not self.selected_risk_id:
            messagebox.showwarning("Delete Risk", "No risk selected.")
            return
        self.model.remove_risk(self.selected_risk_id)
        self.refresh_treeview()
        self.clear_inputs()

    def export_to_excel(self):
        if not self.model.risks:
            messagebox.showwarning("No Data", "No risks to export.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        self.model.save_to_excel(file_path)
        messagebox.showinfo("Exported", f"Risk Register exported to '{file_path}'")

    def load_from_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        self.model.load_from_excel(file_path)
        self.refresh_treeview()
        messagebox.showinfo("Loaded", f"Risks loaded from '{file_path}'")

    def show_risk_chart(self):
        if not self.model.risks:
            messagebox.showwarning("No Data", "No risks to visualize.")
            return

        df = self.model.to_dataframe()
        counts = df['Risk Level'].value_counts().reindex(RISK_LEVEL_ORDER, fill_value=0)

        if self.chart_canvas:
            self.chart_canvas.get_tk_widget().destroy()
            plt.close('all')  # Free matplotlib resources

        fig = plt.Figure(figsize=(6, 3), dpi=100)
        ax = fig.add_subplot(111)
        counts.plot(kind='bar', ax=ax, color='skyblue')
        ax.set_title('Risk Level Distribution')
        ax.set_ylabel('Number of Risks')
        ax.set_xlabel('Risk Level')

        self.chart_canvas = FigureCanvasTkAgg(fig, master=self.root)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().grid(row=8, column=0, columnspan=2, pady=10)

# === Run Application ===
if __name__ == "__main__":
    root = tk.Tk()
    # Optionally: root.geometry("1100x600")
    app = RiskRegisterApp(root)
    root.mainloop()
