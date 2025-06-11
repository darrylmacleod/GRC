# riskreggen.py
# pip install pandas openpyxl tkcalendar matplotlib reportlab

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog, colorchooser
from tkcalendar import DateEntry
import pandas as pd
import json
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from datetime import datetime, date
import threading
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

CONFIG_FILE = "riskreggen_config.json"
AUTOSAVE_FILE = "riskreggen_autosave.csv"
AUTOSAVE_INTERVAL = 120  # seconds

# ==== Config Management ====

DEFAULT_CONFIG = {
    "RISK_LEVEL_THRESHOLDS": {
        "Critical": 15,
        "High": 10,
        "Medium": 5
    },
    "RISK_LEVEL_ORDER": ["Low", "Medium", "High", "Critical"],
    "RISK_COLORS": {
        "Low": "#d4f4dd",
        "Medium": "#fff6b3",
        "High": "#ffd6a0",
        "Critical": "#ffb3b3"
    },
    "DEFAULT_THEME": "light"
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass  # fallback to default
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)

# ==== Tooltip Helper ====
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
                         font=("Segoe UI", "9", "normal"))
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

# ==== Data Model ====
EXCEL_COLUMNS = [
    "Risk ID", "Risk Description", "Impact", "Likelihood", "Risk Score", "Risk Level",
    "Risk Owner", "Due Date", "Notes", "Priority", "History"
]

class RiskRegisterModel:
    def __init__(self):
        self.risks = []
        self.next_id = 1
        self.undo_stack = []
        self.redo_stack = []
        self.history_map = {}  # risk_id -> list of changes

    def add_risk(self, risk):
        risk = risk.copy()
        risk["Risk ID"] = self.next_id
        risk.setdefault("History", "")
        self.risks.append(risk)
        self.next_id += 1
        self._log_history(risk["Risk ID"], "Created")
        self._save_state()

    def remove_risk(self, risk_id):
        self.risks = [r for r in self.risks if r["Risk ID"] != risk_id]
        self._save_state()

    def update_risk(self, risk_id, new_data):
        for idx, r in enumerate(self.risks):
            if r["Risk ID"] == risk_id:
                desc = f"Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                self._log_history(risk_id, desc)
                self.risks[idx].update(new_data)
                break
        self._save_state()

    def duplicate_risk(self, risk_id):
        for risk in self.risks:
            if risk["Risk ID"] == risk_id:
                new_risk = risk.copy()
                new_risk["Risk ID"] = self.next_id
                new_risk["History"] = f"Duplicated from {risk_id} on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                self.add_risk(new_risk)
                break

    def _log_history(self, risk_id, desc):
        for risk in self.risks:
            if risk["Risk ID"] == risk_id:
                if "History" not in risk or not risk["History"]:
                    risk["History"] = desc
                else:
                    risk["History"] += f"\n{desc}"
                break

    def _save_state(self):
        # Keep a snapshot for undo/redo (up to 20)
        state = [r.copy() for r in self.risks]
        self.undo_stack.append(state)
        if len(self.undo_stack) > 20:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def undo(self):
        if self.undo_stack:
            self.redo_stack.append([r.copy() for r in self.risks])
            self.risks = self.undo_stack.pop()

    def redo(self):
        if self.redo_stack:
            self.undo_stack.append([r.copy() for r in self.risks])
            self.risks = self.redo_stack.pop()

    def to_dataframe(self):
        return pd.DataFrame(self.risks, columns=EXCEL_COLUMNS)

    def clear(self):
        self.risks.clear()
        self.next_id = 1
        self.undo_stack.clear()
        self.redo_stack.clear()

    def load_from_excel(self, filename):
        df = pd.read_excel(filename)
        self.risks = df.to_dict(orient='records')
        if self.risks:
            self.next_id = max(r["Risk ID"] for r in self.risks) + 1
        else:
            self.next_id = 1
        self._save_state()

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
        self._save_state()

    def save_to_csv(self, filename):
        df = self.to_dataframe()
        df.to_csv(filename, index=False)

    def load_from_json(self, filename):
        with open(filename, "r") as f:
            self.risks = json.load(f)
        if self.risks:
            self.next_id = max(r["Risk ID"] for r in self.risks) + 1
        else:
            self.next_id = 1
        self._save_state()

    def save_to_json(self, filename):
        with open(filename, "w") as f:
            json.dump(self.risks, f, indent=2)

# ==== Risk Calculation Functions ====
def calculate_risk_score(impact, likelihood):
    return impact * likelihood

def risk_level(score, thresholds):
    if score >= thresholds['Critical']:
        return 'Critical'
    elif score >= thresholds['High']:
        return 'High'
    elif score >= thresholds['Medium']:
        return 'Medium'
    else:
        return 'Low'

# ==== PDF Export ====
def export_to_pdf(df, filename):
    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter
    c.setFont("Helvetica", 10)
    x, y = 40, height - 40
    for col in df.columns:
        c.drawString(x, y, col)
        x += 100
    y -= 20
    x = 40
    for _, row in df.iterrows():
        for col in df.columns:
            text = str(row[col])[:20]
            c.drawString(x, y, text)
            x += 100
        y -= 20
        x = 40
        if y < 50:
            c.showPage()
            y = height - 40
    c.save()

# ==== Main Application ====
class RiskRegisterApp:
    def __init__(self, root):
        self.root = root
        self.config = load_config()
        self.theme = self.config.get("DEFAULT_THEME", "light")
        self._set_theme(self.theme)
        self.root.title("RiskRegGen")
        self.root.geometry("1280x850")
        self.root.minsize(1000, 700)
        self.model = RiskRegisterModel()
        self.selected_risk_id = None
        self.chart_canvas = None
        self.chart_figure = None
        self.last_searched = ""
        self._setup_styles()
        self._setup_menu()
        self._setup_frames()
        self._autosave_timer()
        self._setup_shortcuts()
        self.refresh_treeview()

    # ===== Theme =====
    def _set_theme(self, theme):
        try:
            self.root.tk.call("source", "sun-valley.tcl")
            if theme == "dark":
                ttk.Style().theme_use("sun-valley-dark")
            else:
                ttk.Style().theme_use("sun-valley-light")
        except Exception:
            ttk.Style().theme_use("clam")

    def toggle_theme(self):
        self.theme = "dark" if self.theme == "light" else "light"
        self._set_theme(self.theme)
        self.config["DEFAULT_THEME"] = self.theme
        save_config(self.config)

    # ===== Styles =====
    def _setup_styles(self):
        style = ttk.Style()
        style.configure("TLabel", font=("Segoe UI", 11))
        style.configure("TButton", font=("Segoe UI Semibold", 11), padding=8)
        style.configure("TLabelframe", font=("Segoe UI Semibold", 12, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=28)
        style.configure("TEntry", font=("Segoe UI", 11))
        style.configure("TCombobox", font=("Segoe UI", 11))

    # ===== Menu =====
    def _setup_menu(self):
        menubar = tk.Menu(self.root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Export to Excel", command=self.export_to_excel, accelerator="Ctrl+S")
        filemenu.add_command(label="Export to CSV", command=self.export_to_csv)
        filemenu.add_command(label="Export to PDF", command=self.export_to_pdf)
        filemenu.add_command(label="Export to JSON", command=self.export_to_json)
        filemenu.add_command(label="Export Chart as PNG", command=self.export_chart_png)
        filemenu.add_separator()
        filemenu.add_command(label="Load Risks from Excel", command=self.load_from_excel, accelerator="Ctrl+O")
        filemenu.add_command(label="Load Risks from CSV", command=self.load_from_csv)
        filemenu.add_command(label="Load Risks from JSON", command=self.load_from_json)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.root.destroy)
        menubar.add_cascade(label="File", menu=filemenu)

        editmenu = tk.Menu(menubar, tearoff=0)
        editmenu.add_command(label="Undo", command=self.undo, accelerator="Ctrl+Z")
        editmenu.add_command(label="Redo", command=self.redo, accelerator="Ctrl+Y")
        editmenu.add_separator()
        editmenu.add_command(label="Settings", command=self.open_settings)
        menubar.add_cascade(label="Edit", menu=editmenu)

        thememenu = tk.Menu(menubar, tearoff=0)
        thememenu.add_command(label="Toggle Dark/Light", command=self.toggle_theme)
        menubar.add_cascade(label="Theme", menu=thememenu)

        searchmenu = tk.Menu(menubar, tearoff=0)
        searchmenu.add_command(label="Advanced Search", command=self.advanced_search, accelerator="Ctrl+F")
        menubar.add_cascade(label="Search", menu=searchmenu)

        self.root.config(menu=menubar)

    # ===== Frames and Widgets =====
    def _setup_frames(self):
        # Input frames
        details_frame = ttk.LabelFrame(self.root, text="Risk Details", padding=(16,12))
        details_frame.grid(row=0, column=0, padx=16, pady=(14,7), sticky="ew", columnspan=3)
        owner_frame = ttk.LabelFrame(self.root, text="Owner & Notes", padding=(16,12))
        owner_frame.grid(row=1, column=0, padx=16, pady=7, sticky="ew", columnspan=3)
        # Details
        ttk.Label(details_frame, text="Risk Description").grid(row=0, column=0, sticky="w", padx=6, pady=3)
        self.desc_entry = ttk.Entry(details_frame, width=60)
        self.desc_entry.grid(row=0, column=1, padx=8, pady=4, sticky="ew")
        ToolTip(self.desc_entry, "Describe the risk in detail. Max 200 chars.")
        ttk.Label(details_frame, text="Impact (1–5)").grid(row=1, column=0, sticky="w", padx=6, pady=3)
        self.impact_var = tk.StringVar(value="1")
        self.impact_menu = ttk.Combobox(details_frame, textvariable=self.impact_var, values=[1,2,3,4,5],
                                        state="readonly", width=8)
        self.impact_menu.grid(row=1, column=1, padx=8, sticky='w')
        ToolTip(self.impact_menu, "1=Lowest impact, 5=Highest impact.")
        ttk.Label(details_frame, text="Likelihood (1–5)").grid(row=2, column=0, sticky="w", padx=6, pady=3)
        self.likelihood_var = tk.StringVar(value="1")
        self.likelihood_menu = ttk.Combobox(details_frame, textvariable=self.likelihood_var, values=[1,2,3,4,5],
                                            state="readonly", width=8)
        self.likelihood_menu.grid(row=2, column=1, padx=8, sticky='w')
        ToolTip(self.likelihood_menu, "1=Lowest likelihood, 5=Highest likelihood.")
        ttk.Label(details_frame, text="Priority").grid(row=3, column=0, sticky="w", padx=6, pady=3)
        self.priority_var = tk.StringVar(value="Medium")
        self.priority_menu = ttk.Combobox(details_frame, textvariable=self.priority_var, values=["Low","Medium","High"],
                                          state="readonly", width=8)
        self.priority_menu.grid(row=3, column=1, padx=8, sticky='w')
        ToolTip(self.priority_menu, "Set the urgency/priority of this risk.")
        details_frame.grid_columnconfigure(1, weight=1)
        # Owner & Notes
        ttk.Label(owner_frame, text="Risk Owner").grid(row=0, column=0, sticky="w", padx=6, pady=3)
        self.owner_entry = ttk.Entry(owner_frame, width=30)
        self.owner_entry.grid(row=0, column=1, padx=8, sticky='w')
        ToolTip(self.owner_entry, "Person responsible for managing this risk. Max 50 chars.")
        ttk.Label(owner_frame, text="Due Date").grid(row=1, column=0, sticky="w", padx=6, pady=3)
        self.due_date = DateEntry(owner_frame, width=14, background='darkblue', foreground='white',
                                  borderwidth=2, date_pattern='yyyy-mm-dd', font=("Segoe UI", 11))
        self.due_date.grid(row=1, column=1, padx=8, sticky='w')
        ToolTip(self.due_date, "Date by which the risk should be addressed.")
        ttk.Label(owner_frame, text="Notes").grid(row=2, column=0, sticky="w", padx=6, pady=3)
        self.notes_entry = ttk.Entry(owner_frame, width=60)
        self.notes_entry.grid(row=2, column=1, padx=8, pady=4, sticky="ew")
        ToolTip(self.notes_entry, "Additional information or mitigation plan. Max 200 chars.")
        owner_frame.grid_columnconfigure(1, weight=1)
        # Buttons
        actions_frame = ttk.Frame(self.root)
        actions_frame.grid(row=2, column=0, padx=16, pady=12, sticky="ew", columnspan=3)
        actions_frame.grid_columnconfigure(8, weight=1)
        ttk.Button(actions_frame, text="Add/Update Risk", command=self.add_or_update_risk).grid(row=0, column=0, padx=6, pady=4)
        ttk.Button(actions_frame, text="Delete Selected Risk", command=self.delete_selected_risk).grid(row=0, column=1, padx=6, pady=4)
        ttk.Button(actions_frame, text="Clear Form", command=self.clear_inputs).grid(row=0, column=2, padx=6, pady=4)
        ttk.Button(actions_frame, text="Duplicate Risk", command=self.duplicate_risk).grid(row=0, column=3, padx=6, pady=4)
        ttk.Button(actions_frame, text="Show Risk Chart", command=self.show_risk_chart).grid(row=0, column=4, padx=6, pady=4)
        ttk.Button(actions_frame, text="View Change History", command=self.view_history).grid(row=0, column=5, padx=6, pady=4)
        # Search bar
        search_frame = ttk.Frame(self.root)
        search_frame.grid(row=3, column=0, columnspan=3, pady=5, sticky='ew')
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=6)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side='left', padx=2)
        search_entry.bind('<KeyRelease>', self.perform_search)
        # Table
        table_frame = ttk.Frame(self.root)
        table_frame.grid(row=4, column=0, columnspan=3, padx=16, pady=12, sticky="nsew")
        columns = EXCEL_COLUMNS
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=18)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=115, anchor='center')
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        self.tree.bind('<ButtonRelease-1>', self.on_tree_select)
        # Chart Placeholder
        self.chart_canvas = None
        self.chart_figure = None

    # ===== Keyboard Shortcuts =====
    def _setup_shortcuts(self):
        self.root.bind("<Control-s>", lambda e: self.export_to_excel())
        self.root.bind("<Control-o>", lambda e: self.load_from_excel())
        self.root.bind("<Control-z>", lambda e: self.undo())
        self.root.bind("<Control-y>", lambda e: self.redo())
        self.root.bind("<Control-f>", lambda e: self.advanced_search())

    # ===== Validation =====
    def validate_inputs(self):
        desc = self.desc_entry.get().strip()
        owner = self.owner_entry.get().strip()
        notes = self.notes_entry.get().strip()
        priority = self.priority_var.get()
        if len(desc) == 0 or len(owner) == 0:
            messagebox.showerror("Error", "Please fill in all required fields (Description, Owner).")
            return None
        if len(desc) > 200 or len(owner) > 50 or len(notes) > 200:
            messagebox.showerror("Error", "Text fields exceed maximum allowed length.")
            return None
        try:
            impact = int(self.impact_var.get())
            likelihood = int(self.likelihood_var.get())
            if not (1 <= impact <= 5) or not (1 <= likelihood <= 5):
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Impact and Likelihood must be integers from 1 to 5.")
            return None
        due = self.due_date.get_date()
        if due < date.today():
            messagebox.showerror("Error", "Due Date cannot be in the past.")
            return None
        if (due - date.today()).days > 365*5:
            messagebox.showerror("Error", "Due Date cannot be more than 5 years in the future.")
            return None
        return desc, impact, likelihood, owner, due, notes, priority

    # ===== Add or Update Risk =====
    def add_or_update_risk(self):
        validated = self.validate_inputs()
        if not validated:
            return
        desc, impact, likelihood, owner, due, notes, priority = validated
        score = calculate_risk_score(impact, likelihood)
        level = risk_level(score, self.config["RISK_LEVEL_THRESHOLDS"])
        risk = {
            "Risk Description": desc,
            "Impact": impact,
            "Likelihood": likelihood,
            "Risk Score": score,
            "Risk Level": level,
            "Risk Owner": owner,
            "Due Date": due,
            "Notes": notes,
            "Priority": priority
        }
        if self.selected_risk_id:  # Update
            risk["Risk ID"] = self.selected_risk_id
            self.model.update_risk(self.selected_risk_id, risk)
            self.refresh_treeview()
            self.selected_risk_id = None
            messagebox.showinfo("Updated", "Risk updated successfully.")
        else:  # Add new
            self.model.add_risk(risk)
            self.refresh_treeview()
            messagebox.showinfo("Added", "Risk added successfully.")
        self.clear_inputs()

    # ===== Treeview =====
    def refresh_treeview(self, filtered=None):
        for row in self.tree.get_children():
            self.tree.delete(row)
        risks = filtered if filtered is not None else self.model.risks
        for risk in risks:
            self.insert_treeview_row(risk)

    def insert_treeview_row(self, risk):
        color = self.config["RISK_COLORS"].get(risk["Risk Level"], "#fff")
        values = (
            risk.get("Risk ID", ""),
            risk.get("Risk Description", ""),
            risk.get("Impact", ""),
            risk.get("Likelihood", ""),
            risk.get("Risk Score", ""),
            risk.get("Risk Level", ""),
            risk.get("Risk Owner", ""),
            risk.get("Due Date", "").strftime('%Y-%m-%d') if isinstance(risk.get("Due Date", ""), (datetime, date)) else str(risk.get("Due Date", "")),
            risk.get("Notes", ""),
            risk.get("Priority", ""),
            risk.get("History", "")[:30]  # show a snippet only
        )
        self.tree.insert("", "end", values=values, tags=(risk["Risk Level"],))
        self.tree.tag_configure(risk["Risk Level"], background=color)

    def clear_inputs(self):
        self.desc_entry.delete(0, tk.END)
        self.owner_entry.delete(0, tk.END)
        self.notes_entry.delete(0, tk.END)
        self.impact_var.set('1')
        self.likelihood_var.set('1')
        self.priority_var.set('Medium')
        self.selected_risk_id = None
        self.tree.selection_remove(self.tree.selection())

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if selected:
            values = self.tree.item(selected[0], 'values')
            self.selected_risk_id = int(values[0])
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
            self.priority_var.set(values[9])

    # ===== Risk Functions =====
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

    def duplicate_risk(self):
        if not self.selected_risk_id:
            messagebox.showwarning("Duplicate Risk", "No risk selected.")
            return
        self.model.duplicate_risk(self.selected_risk_id)
        self.refresh_treeview()
        messagebox.showinfo("Duplicated", "Risk duplicated successfully.")

    def view_history(self):
        if not self.selected_risk_id:
            messagebox.showwarning("View History", "No risk selected.")
            return
        for risk in self.model.risks:
            if risk["Risk ID"] == self.selected_risk_id:
                history = risk.get("History", "")
                messagebox.showinfo(f"History for Risk {self.selected_risk_id}", history or "No history available.")
                return

    # ===== Export/Import =====
    def export_to_excel(self, *_):
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

    def export_to_pdf(self):
        if not self.model.risks:
            messagebox.showwarning("No Data", "No risks to export.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not file_path:
            return
        try:
            df = self.model.to_dataframe()
            export_to_pdf(df, file_path)
            messagebox.showinfo("Exported", f"Risk Register exported to '{file_path}'")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export: {e}")

    def export_to_json(self):
        if not self.model.risks:
            messagebox.showwarning("No Data", "No risks to export.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if not file_path:
            return
        try:
            self.model.save_to_json(file_path)
            messagebox.showinfo("Exported", f"Risk Register exported to '{file_path}'")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export: {e}")

    def load_from_excel(self, *_):
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

    def load_from_json(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not file_path:
            return
        if self.model.risks and not messagebox.askyesno("Overwrite Register?", "Loading will overwrite the current register. Continue?"):
            return
        try:
            self.model.load_from_json(file_path)
            self.refresh_treeview()
            messagebox.showinfo("Loaded", f"Risks loaded from '{file_path}'")
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load: {e}")

    # ===== Chart =====
    def show_risk_chart(self):
        if not self.model.risks:
            messagebox.showwarning("No Data", "No risks to visualize.")
            return
        df = self.model.to_dataframe()
        counts = df['Risk Level'].value_counts().reindex(self.config["RISK_LEVEL_ORDER"], fill_value=0)
        if self.chart_canvas:
            self.chart_canvas.get_tk_widget().destroy()
            plt.close('all')
        self.chart_figure = fig = plt.Figure(figsize=(6, 3.5), dpi=100)
        ax = fig.add_subplot(111)
        bars = counts.plot(kind='bar', ax=ax, color=[self.config["RISK_COLORS"][rl] for rl in self.config["RISK_LEVEL_ORDER"]])
        ax.set_title('Risk Level Distribution')
        ax.set_ylabel('Number of Risks')
        ax.set_xlabel('Risk Level')
        for i, v in enumerate(counts):
            ax.text(i, v + 0.1, str(v), ha='center', va='bottom', fontsize=10)
        ax.legend(["# of Risks"])
        self.chart_canvas = FigureCanvasTkAgg(fig, master=self.root)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().grid(row=5, column=0, columnspan=3, pady=10)

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

    # ===== Search =====
    def perform_search(self, event=None):
        search_term = self.search_var.get().lower()
        if not search_term:
            self.refresh_treeview()
            return
        filtered = [r for r in self.model.risks if search_term in str(r).lower()]
        self.refresh_treeview(filtered)

    def advanced_search(self, *_):
        dialog = tk.Toplevel(self.root)
        dialog.title("Advanced Search")
        tk.Label(dialog, text="Risk Level:").grid(row=0, column=0, sticky="w", padx=6, pady=3)
        level_var = tk.StringVar(value="")
        ttk.Combobox(dialog, textvariable=level_var, values=[""]+self.config["RISK_LEVEL_ORDER"], state="readonly").grid(row=0, column=1, padx=8, pady=3)
        tk.Label(dialog, text="Owner:").grid(row=1, column=0, sticky="w", padx=6, pady=3)
        owner_var = tk.StringVar()
        tk.Entry(dialog, textvariable=owner_var).grid(row=1, column=1, padx=8, pady=3)
        tk.Label(dialog, text="Priority:").grid(row=2, column=0, sticky="w", padx=6, pady=3)
        priority_var = tk.StringVar(value="")
        ttk.Combobox(dialog, textvariable=priority_var, values=["","Low","Medium","High"], state="readonly").grid(row=2, column=1, padx=8, pady=3)
        tk.Label(dialog, text="Due Date (Before):").grid(row=3, column=0, sticky="w", padx=6, pady=3)
        due_var = tk.StringVar()
        DateEntry(dialog, textvariable=due_var, width=14, background='darkblue', foreground='white',
                  borderwidth=2, date_pattern='yyyy-mm-dd', font=("Segoe UI", 11)).grid(row=3, column=1, padx=8, pady=3)
        def do_search():
            results = self.model.risks
            if level_var.get():
                results = [r for r in results if r["Risk Level"] == level_var.get()]
            if owner_var.get():
                results = [r for r in results if owner_var.get().lower() in str(r["Risk Owner"]).lower()]
            if priority_var.get():
                results = [r for r in results if r.get("Priority","") == priority_var.get()]
            if due_var.get():
                try:
                    search_due = datetime.strptime(due_var.get(), "%Y-%m-%d").date()
                    results = [r for r in results if isinstance(r["Due Date"], (datetime, date)) and r["Due Date"] <= search_due]
                except Exception:
                    pass
            self.refresh_treeview(results)
            dialog.destroy()
        ttk.Button(dialog, text="Search", command=do_search).grid(row=4, column=0, columnspan=2, pady=10)
        dialog.grab_set()

    # ===== Undo/Redo =====
    def undo(self, *_):
        self.model.undo()
        self.refresh_treeview()

    def redo(self, *_):
        self.model.redo()
        self.refresh_treeview()

    # ===== Settings =====
    def open_settings(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Settings")
        tk.Label(dialog, text="Risk Level Thresholds:").grid(row=0, column=0, sticky="w", padx=6, pady=3)
        entries = {}
        for i, key in enumerate(self.config["RISK_LEVEL_THRESHOLDS"]):
            tk.Label(dialog, text=f"{key}:").grid(row=i+1, column=0, sticky="w", padx=6, pady=3)
            var = tk.IntVar(value=self.config["RISK_LEVEL_THRESHOLDS"][key])
            entries[key] = var
            tk.Entry(dialog, textvariable=var).grid(row=i+1, column=1, padx=8, pady=3)
        tk.Label(dialog, text="Risk Level Colors:").grid(row=0, column=2, sticky="w", padx=6, pady=3)
        color_vars = {}
        for i, key in enumerate(self.config["RISK_COLORS"]):
            tk.Label(dialog, text=f"{key}:").grid(row=i+1, column=2, sticky="w", padx=6, pady=3)
            var = tk.StringVar(value=self.config["RISK_COLORS"][key])
            color_vars[key] = var
            def choose_color(k=key, v=var):
                color = colorchooser.askcolor(title=f"Choose color for {k}")
                if color and color[1]:
                    v.set(color[1])
            btn = ttk.Button(dialog, text="Pick", command=choose_color)
            btn.grid(row=i+1, column=4, padx=2, pady=3)
            tk.Entry(dialog, textvariable=var, width=10).grid(row=i+1, column=3, padx=2, pady=3)
        def save_settings():
            for k in entries:
                self.config["RISK_LEVEL_THRESHOLDS"][k] = entries[k].get()
            for k in color_vars:
                self.config["RISK_COLORS"][k] = color_vars[k].get()
            save_config(self.config)
            dialog.destroy()
            messagebox.showinfo("Saved", "Settings saved. Changes will apply to new/updated risks and next chart.")
        ttk.Button(dialog, text="Save", command=save_settings).grid(row=6, column=2, columnspan=2, pady=10)
        dialog.grab_set()

    # ===== Autosave =====
    def _autosave_timer(self):
        def do_autosave():
            if self.model.risks:
                try:
                    self.model.save_to_csv(AUTOSAVE_FILE)
                except Exception:
                    pass
            self.root.after(AUTOSAVE_INTERVAL*1000, self._autosave_timer)
        self.root.after(AUTOSAVE_INTERVAL*1000, do_autosave)

# ===== Run Application =====
if __name__ == "__main__":
    root = tk.Tk()
    app = RiskRegisterApp(root)
    root.grid_rowconfigure(4, weight=1)
    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)
    root.mainloop()
