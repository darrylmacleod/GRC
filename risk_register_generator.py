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
RISK_COLORS = {
    'Low': '#d4f4dd',
    'Medium': '#fff6b3',
    'High': '#ffd6a0',
    'Critical': '#ffb3b3'
}

# === Tooltip Helper ===
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tipwindow or not self.text:
            return
        x, y, _, cy = self.widget.bbox("insert") if hasattr(self.widget, "bbox") else (0,0,0,0)
        x = x + self.widget.winfo_rootx() + 25
        y = y + cy + self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

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

    def load_from_csv(self, filename):
        df = pd.read_csv(filename)
        self.risks = df.to_dict(orient='records')
        if self.risks:
            self.next_id = max(r["Risk ID"] for r in self.risks) + 1
        else:
            self.next_id = 1

    def save_to_csv(self, filename):
        df = self.to_dataframe()
        df.to_csv(filename, index=False)

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
        self.root.title("Darryl's Risk Register Generator")
        self.root.geometry("1200x750")

        # Make the layout responsive
        self.root.grid_rowconfigure(4, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        self.model = RiskRegisterModel()
        self.selected_risk_id = None

        # --- Input Frames ---
        details_frame = ttk.LabelFrame(root, text="Risk Details", padding=(10,10))
        details_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew", columnspan=2)
        owner_frame = ttk.LabelFrame(root, text="Owner & Notes", padding=(10,10))
        owner_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew", columnspan=2)

        # --- Inputs: Risk Details ---
        tk.Label(details_frame, text="Risk Description").grid(row=0, column=0, sticky="w")
        self.desc_entry = tk.Entry(details_frame, width=50)
        self.desc_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ToolTip(self.desc_entry, "Describe the risk in detail.")

        tk.Label(details_frame, text="Impact (1–5)").grid(row=1, column=0, sticky="w")
        self.impact_var = tk.StringVar(value="1")
        self.impact_menu = ttk.Combobox(details_frame, textvariable=self.impact_var, values=[1, 2, 3, 4, 5], 
                                        state="readonly", width=10)
        self.impact_menu.grid(row=1, column=1, sticky='w')
        ToolTip(self.impact_menu, "1=Lowest impact, 5=Highest impact.")

        tk.Label(details_frame, text="Likelihood (1–5)").grid(row=2, column=0, sticky="w")
        self.likelihood_var = tk.StringVar(value="1")
        self.likelihood_menu = ttk.Combobox(details_frame, textvariable=self.likelihood_var, values=[1, 2, 3, 4, 5],
                                            state="readonly", width=10)
        self.likelihood_menu.grid(row=2, column=1, sticky='w')
        ToolTip(self.likelihood_menu, "1=Lowest likelihood, 5=Highest likelihood.")

        # --- Inputs: Owner & Notes ---
        tk.Label(owner_frame, text="Risk Owner").grid(row=0, column=0, sticky="w")
        self.owner_entry = tk.Entry(owner_frame, width=30)
        self.owner_entry.grid(row=0, column=1, sticky='w')
        ToolTip(self.owner_entry, "Person responsible for managing this risk.")

        tk.Label(owner_frame, text="Due Date").grid(row=1, column=0, sticky="w")
        self.due_date = DateEntry(owner_frame, width=12, background='darkblue', foreground='white',
                                  borderwidth=2, date_pattern='yyyy-mm-dd')
        self.due_date.grid(row=1, column=1, sticky='w')
        ToolTip(self.due_date, "Date by which the risk should be addressed.")

        tk.Label(owner_frame, text="Notes").grid(row=2, column=0, sticky="w")
        self.notes_entry = tk.Entry(owner_frame, width=50)
        self.notes_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        ToolTip(self.notes_entry, "Additional information or mitigation plan.")

        details_frame.grid_columnconfigure(1, weight=1)
        owner_frame.grid_columnconfigure(1, weight=1)

        # --- Action Buttons Frame ---
        actions_frame = ttk.Frame(root)
        actions_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew", columnspan=2)

        tk.Button(actions_frame, text="Add/Update Risk", command=self.add_or_update_risk).grid(row=0, column=0, padx=5, pady=5)
        tk.Button(actions_frame, text="Delete Selected Risk", command=self.delete_selected_risk).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(actions_frame, text="Clear Form", command=self.clear_inputs).grid(row=0, column=2, padx=5, pady=5)
        tk.Button(actions_frame, text="Export to Excel", command=self.export_to_excel).grid(row=0, column=3, padx=5, pady=5)
        tk.Button(actions_frame, text="Export to CSV", command=self.export_to_csv).grid(row=0, column=4, padx=5, pady=5)
        tk.Button(actions_frame, text="Load Risks from Excel", command=self.load_from_excel).grid(row=0, column=5, padx=5, pady=5)
        tk.Button(actions_frame, text="Load Risks from CSV", command=self.load_from_csv).grid(row=0, column=6, padx=5, pady=5)
        tk.Button(actions_frame, text="Show Risk Chart", command=self.show_risk_chart).grid(row=0, column=7, padx=5, pady=5)
        tk.Button(actions_frame, text="Export Chart as PNG", command=self.export_chart_png).grid(row=0, column=8, padx=5, pady=5)

        # --- Search Bar ---
        search_frame = ttk.Frame(root)
        search_frame.grid(row=3, column=0, columnspan=2, pady=5, sticky='ew')
        tk.Label(search_frame, text="Search:").pack(side='left')
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(side='left')
        search_entry.bind('<KeyRelease>', self.perform_search)

        # --- Treeview Table with Scrollbars ---
        table_frame = ttk.Frame(root)
        table_frame.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        columns = EXCEL_COLUMNS
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor='center')
        self.tree.grid(row=0, column=0, sticky='nsew')

        v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # Enable row selection for edit/delete
        self.tree.bind('<ButtonRelease-1>', self.on_tree_select)

        # Chart Placeholder and Data
        self.chart_canvas = None
        self.chart_figure = None

    def validate_inputs(self):
        desc = self.desc_entry.get().strip()
        owner = self.owner_entry.get().strip()
        notes = self.notes_entry.get().strip()
        try:
            impact = int(self.impact_var.get())
            likelihood = int(self.likelihood_var.get())
            if not (1 <= impact <= 5) or not (1 <= likelihood <= 5):
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Impact and Likelihood must be integers from 1 to 5.")
            return None
        if not desc or not owner:
            messagebox.showerror("Error", "Please fill in all required fields (Description, Owner).")
            return None
        due = self.due_date.get_date()
        if due < date.today():
            messagebox.showerror("Error", "Due Date cannot be in the past.")
            return None
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
            messagebox.showinfo("Updated", "Risk updated successfully.")
        else:  # Add new
            self.model.add_risk(risk)
            self.insert_treeview_row(risk)
            messagebox.showinfo("Added", "Risk added successfully.")
        self.clear_inputs()

    def refresh_treeview(self, filtered=None):
        for row in self.tree.get_children():
            self.tree.delete(row)
        risks = filtered if filtered is not None else self.model.risks
        for risk in risks:
            self.insert_treeview_row(risk)

    def insert_treeview_row(self, risk):
        color = RISK_COLORS.get(risk["Risk Level"], "#fff")
        values = (
            risk["Risk ID"], risk["Risk Description"], risk["Impact"], risk["Likelihood"],
            risk["Risk Score"], risk["Risk Level"], risk["Risk Owner"],
            risk["Due Date"].strftime('%Y-%m-%d') if isinstance(risk["Due Date"], (datetime, date)) else str(risk["Due Date"]),
            risk["Notes"]
        )
        self.tree.insert("", "end", values=values, tags=(risk["Risk Level"],))
        self.tree.tag_configure(risk["Risk Level"], background=color)

    def clear_inputs(self):
        self.desc_entry.delete(0, tk.END)
        self.owner_entry.delete(0, tk.END)
        self.notes_entry.delete(0, tk.END)
        self.impact_var.set('1')
        self.likelihood_var.set('1')
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
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected risk?"):
            return
        self.model.remove_risk(self.selected_risk_id)
        self.refresh_treeview()
        self.clear_inputs()
        messagebox.showinfo("Deleted", "Risk deleted successfully.")

    def export_to_excel(self):
        if not self.model.risks:
            messagebox.showwarning("No Data", "No risks to export.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        try:
            self.model.save_to_excel(file_path)
            messagebox.showinfo("Exported", f"Risk Register exported to '{file_path}'")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export: {e}")

    def export_to_csv(self):
        if not self.model.risks:
            messagebox.showwarning("No Data", "No risks to export.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return
        try:
            self.model.save_to_csv(file_path)
            messagebox.showinfo("Exported", f"Risk Register exported to '{file_path}'")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export: {e}")

    def load_from_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        if self.model.risks and not messagebox.askyesno("Overwrite Register?", "Loading will overwrite the current register. Continue?"):
            return
        try:
            self.model.load_from_excel(file_path)
            self.refresh_treeview()
            messagebox.showinfo("Loaded", f"Risks loaded from '{file_path}'")
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load: {e}")

    def load_from_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return
        if self.model.risks and not messagebox.askyesno("Overwrite Register?", "Loading will overwrite the current register. Continue?"):
            return
        try:
            self.model.load_from_csv(file_path)
            self.refresh_treeview()
            messagebox.showinfo("Loaded", f"Risks loaded from '{file_path}'")
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load: {e}")

    def show_risk_chart(self):
        if not self.model.risks:
            messagebox.showwarning("No Data", "No risks to visualize.")
            return

        df = self.model.to_dataframe()
        counts = df['Risk Level'].value_counts().reindex(RISK_LEVEL_ORDER, fill_value=0)

        if self.chart_canvas:
            self.chart_canvas.get_tk_widget().destroy()
            plt.close('all')
        self.chart_figure = fig = plt.Figure(figsize=(6, 3), dpi=100)
        ax = fig.add_subplot(111)
        bars = counts.plot(kind='bar', ax=ax, color=[RISK_COLORS[rl] for rl in RISK_LEVEL_ORDER])
        ax.set_title('Risk Level Distribution')
        ax.set_ylabel('Number of Risks')
        ax.set_xlabel('Risk Level')
        for i, v in enumerate(counts):
            ax.text(i, v + 0.1, str(v), ha='center', va='bottom', fontsize=10)
        self.chart_canvas = FigureCanvasTkAgg(fig, master=self.root)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().grid(row=5, column=0, columnspan=2, pady=10)

    def export_chart_png(self):
        if not self.chart_figure:
            messagebox.showwarning("No Chart", "Please generate the chart first by clicking 'Show Risk Chart'.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
        if not file_path:
            return
        try:
            self.chart_figure.savefig(file_path)
            messagebox.showinfo("Exported", f"Chart exported as PNG to '{file_path}'")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export chart: {e}")

    def perform_search(self, event=None):
        search_term = self.search_var.get().lower()
        if not search_term:
            self.refresh_treeview()
            return
        filtered = [r for r in self.model.risks if search_term in str(r).lower()]
        self.refresh_treeview(filtered)

# === Run Application ===
if __name__ == "__main__":
    root = tk.Tk()
    app = RiskRegisterApp(root)
    root.mainloop()
