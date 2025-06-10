# Install Required Libraries
# pip install pandas openpyxl tkcalendar matplotlib

import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

# === Risk Calculation Functions ===
def calculate_risk_score(impact, likelihood):
    return impact * likelihood

def risk_level(score):
    if score >= 15:
        return 'Critical'
    elif score >= 10:
        return 'High'
    elif score >= 5:
        return 'Medium'
    else:
        return 'Low'

# === Main Application Class ===
class RiskRegisterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automated Risk Register")

        self.risks = []
        self.risk_id = 1

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
        self.due_date = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.due_date.grid(row=4, column=1, sticky='w')

        tk.Label(root, text="Notes").grid(row=5, column=0)
        self.notes_entry = tk.Entry(root, width=50)
        self.notes_entry.grid(row=5, column=1, padx=10, pady=5)

        # === Buttons ===
        tk.Button(root, text="Add Risk", command=self.add_risk).grid(row=6, column=0, pady=10)
        tk.Button(root, text="Export to Excel", command=self.export_to_excel).grid(row=6, column=1, sticky='w')
        tk.Button(root, text="Show Risk Chart", command=self.show_risk_chart).grid(row=6, column=1, sticky='e')

        # === Treeview Table ===
        columns = ('ID', 'Description', 'Impact', 'Likelihood', 'Score', 'Level', 'Owner', 'Due Date', 'Notes')
        self.tree = ttk.Treeview(root, columns=columns, show='headings', height=10)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110)
        self.tree.grid(row=7, column=0, columnspan=2, padx=10, pady=10)

        # === Chart Placeholder ===
        self.chart_canvas = None

    def add_risk(self):
        desc = self.desc_entry.get()
        try:
            impact = int(self.impact_var.get())
            likelihood = int(self.likelihood_var.get())
        except ValueError:
            messagebox.showerror("Error", "Please select Impact and Likelihood (1–5).")
            return

        owner = self.owner_entry.get()
        due = self.due_date.get_date()
        notes = self.notes_entry.get()

        if not desc or not owner:
            messagebox.showerror("Error", "Please fill in all required fields.")
            return

        score = calculate_risk_score(impact, likelihood)
        level = risk_level(score)

        risk = {
            "Risk ID": self.risk_id,
            "Risk Description": desc,
            "Impact": impact,
            "Likelihood": likelihood,
            "Risk Score": score,
            "Risk Level": level,
            "Risk Owner": owner,
            "Due Date": due,
            "Notes": notes
        }

        self.risks.append(risk)
        self.tree.insert("", "end", values=(
            self.risk_id, desc, impact, likelihood, score, level, owner, due.strftime('%Y-%m-%d'), notes
        ))
        self.risk_id += 1

        # Clear inputs
        self.desc_entry.delete(0, tk.END)
        self.owner_entry.delete(0, tk.END)
        self.notes_entry.delete(0, tk.END)
        self.impact_var.set('')
        self.likelihood_var.set('')

    def export_to_excel(self):
        if not self.risks:
            messagebox.showwarning("No Data", "No risks to export.")
            return
        df = pd.DataFrame(self.risks)
        df.to_excel("risk_register_gui.xlsx", index=False)
        messagebox.showinfo("Exported", "Risk Register exported to 'risk_register_gui.xlsx'")

    def show_risk_chart(self):
        if not self.risks:
            messagebox.showwarning("No Data", "No risks to visualize.")
            return

        df = pd.DataFrame(self.risks)
        counts = df['Risk Level'].value_counts().reindex(['Low', 'Medium', 'High', 'Critical'], fill_value=0)

        if self.chart_canvas:
            self.chart_canvas.get_tk_widget().destroy()

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
    app = RiskRegisterApp(root)
    root.mainloop()
