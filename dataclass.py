import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import json
import logging
from typing import Dict, List, Optional, Any
from difflib import get_close_matches
from enum import Enum, auto
import openpyxl

try:
    from rapidfuzz import process as rapidfuzz_process
    FUZZY_LIB = 'rapidfuzz'
except ImportError:
    FUZZY_LIB = 'difflib'

# --- Logging setup ---
logging.basicConfig(filename='dataclassifier.log', level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')


class ClassificationLevel(Enum):
    PUBLIC = "Public"
    INTERNAL = "Internal"
    CONFIDENTIAL = "Confidential"
    RESTRICTED = "Restricted"

    @classmethod
    def list(cls):
        return [e.value for e in cls]


# --- Classification Data Model ---
class ClassificationMap:
    """
    Handles classification data and logic.

    Attributes:
        classification_map (Dict[str, str]): Maps data types to classification levels.
        classification_levels (List[str]): List of known classification levels.
    """
    def __init__(self):
        self.classification_map: Dict[str, str] = {
            # Public
            "Company website content": ClassificationLevel.PUBLIC.value,
            "Marketing materials": ClassificationLevel.PUBLIC.value,
            "Press releases": ClassificationLevel.PUBLIC.value,
            "Published research papers": ClassificationLevel.PUBLIC.value,
            "Public regulatory filings": ClassificationLevel.PUBLIC.value,
            # Internal
            "Internal memos and communications": ClassificationLevel.INTERNAL.value,
            "Organization charts": ClassificationLevel.INTERNAL.value,
            "Training materials": ClassificationLevel.INTERNAL.value,
            "Project plans and status reports": ClassificationLevel.INTERNAL.value,
            "Meeting agendas and minutes (non-sensitive)": ClassificationLevel.INTERNAL.value,
            # Confidential
            "Employee ID numbers": ClassificationLevel.CONFIDENTIAL.value,
            "Business strategies": ClassificationLevel.CONFIDENTIAL.value,
            "Contract details": ClassificationLevel.CONFIDENTIAL.value,
            "Internal financial statements": ClassificationLevel.CONFIDENTIAL.value,
            "Customer contact information": ClassificationLevel.CONFIDENTIAL.value,
            "Source code": ClassificationLevel.CONFIDENTIAL.value,
            "Non-public pricing or product roadmaps": ClassificationLevel.CONFIDENTIAL.value,
            "Vendor agreements": ClassificationLevel.CONFIDENTIAL.value,
            "Intellectual property documentation": ClassificationLevel.CONFIDENTIAL.value,
            # Restricted
            "Full names, addresses, phone numbers": ClassificationLevel.RESTRICTED.value,
            "Social Insurance Numbers (SIN)/Social Security Numbers (SSN)": ClassificationLevel.RESTRICTED.value,
            "Driver’s license numbers": ClassificationLevel.RESTRICTED.value,
            "Dates of birth": ClassificationLevel.RESTRICTED.value,
            "Passport numbers": ClassificationLevel.RESTRICTED.value,
            "Medical records": ClassificationLevel.RESTRICTED.value,
            "Health insurance data": ClassificationLevel.RESTRICTED.value,
            "Lab test results": ClassificationLevel.RESTRICTED.value,
            "Appointment histories": ClassificationLevel.RESTRICTED.value,
            "Credit card numbers": ClassificationLevel.RESTRICTED.value,
            "CVV codes": ClassificationLevel.RESTRICTED.value,
            "Cardholder names and billing addresses": ClassificationLevel.RESTRICTED.value,
            "Bank account numbers": ClassificationLevel.RESTRICTED.value,
            "Routing numbers": ClassificationLevel.RESTRICTED.value,
            "Tax returns": ClassificationLevel.RESTRICTED.value,
            "Passwords": ClassificationLevel.RESTRICTED.value,
            "API keys": ClassificationLevel.RESTRICTED.value,
            "Encryption keys": ClassificationLevel.RESTRICTED.value,
            "Biometric data": ClassificationLevel.RESTRICTED.value,
            "Litigation documents": ClassificationLevel.RESTRICTED.value,
            "Legal holds": ClassificationLevel.RESTRICTED.value,
            "Regulatory investigation materials": ClassificationLevel.RESTRICTED.value,
            "Board meeting minutes": ClassificationLevel.RESTRICTED.value,
            "Merger/acquisition plans": ClassificationLevel.RESTRICTED.value,
            "Due diligence documents": ClassificationLevel.RESTRICTED.value,
        }
        self.classification_levels: List[str] = ClassificationLevel.list()

    def add_classification(self, data_type: str, level: str) -> bool:
        """
        Add a new data type and classification level.

        Returns True if added, False if invalid input.
        """
        if data_type and level:
            self.classification_map[data_type] = level
            if level not in self.classification_levels:
                self.classification_levels.append(level)
            logging.info(f"Added data type: {data_type} with level: {level}")
            return True
        logging.warning(f"Attempted to add invalid classification: {data_type}, {level}")
        return False

    def import_csv(self, file_path: str) -> None:
        """Import classification data from a CSV file. Skips invalid rows and logs warnings."""
        try:
            with open(file_path, mode='r', encoding='utf-8') as file:
                reader = csv.reader(file)
                next(reader, None)  # Skip header
                for row in reader:
                    if len(row) == 2:
                        self.add_classification(row[0], row[1])
                    else:
                        logging.warning(f"Skipped invalid CSV row: {row}")
        except Exception as e:
            logging.exception(f"Error importing CSV: {e}")
            raise

    def export_csv(self, file_path: str) -> None:
        """Export classification data to a CSV file."""
        try:
            with open(file_path, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(["Data Type", "Classification"])
                for data_type, level in self.classification_map.items():
                    writer.writerow([data_type, level])
        except Exception as e:
            logging.exception(f"Error exporting CSV: {e}")
            raise

    def import_excel(self, file_path: str) -> None:
        """Import classification data from an Excel (.xlsx) file."""
        try:
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                if row and len(row) >= 2:
                    self.add_classification(str(row[0]), str(row[1]))
                else:
                    logging.warning(f"Skipped invalid Excel row at {i}: {row}")
        except Exception as e:
            logging.exception(f"Error importing Excel: {e}")
            raise

    def export_excel(self, file_path: str) -> None:
        """Export classification data to an Excel (.xlsx) file."""
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Data Type", "Classification"])
            for data_type, level in self.classification_map.items():
                ws.append([data_type, level])
            wb.save(file_path)
        except Exception as e:
            logging.exception(f"Error exporting Excel: {e}")
            raise

    def import_json(self, file_path: str) -> None:
        """Import classification data from a JSON file."""
        try:
            with open(file_path, mode='r', encoding='utf-8') as file:
                data = json.load(file)
                for obj in data:
                    if isinstance(obj, dict) and "Data Type" in obj and "Classification" in obj:
                        self.add_classification(obj["Data Type"], obj["Classification"])
                    else:
                        logging.warning(f"Skipped invalid JSON entry: {obj}")
        except Exception as e:
            logging.exception(f"Error importing JSON: {e}")
            raise

    def export_json(self, file_path: str) -> None:
        """Export classification data to a JSON file."""
        try:
            data = [{"Data Type": dt, "Classification": cl} for dt, cl in self.classification_map.items()]
            with open(file_path, mode='w', encoding='utf-8') as file:
                json.dump(data, file, indent=2)
        except Exception as e:
            logging.exception(f"Error exporting JSON: {e}")
            raise

    def get_classification(self, data_type: str) -> str:
        """
        Get classification for a data type, or fuzzy match if not found.

        Args:
            data_type (str): The data type to classify.

        Returns:
            str: The classification level, or "Unknown" if not found.
        """
        result = self.classification_map.get(data_type)
        if result:
            logging.info(f"Exact match for '{data_type}' as '{result}'")
            return result
        # Fuzzy matching
        keys = list(self.classification_map.keys())
        if FUZZY_LIB == 'rapidfuzz':
            matches = rapidfuzz_process.extract(data_type, keys, limit=1, score_cutoff=60)
            if matches:
                logging.info(f"Fuzzy match for '{data_type}' as '{matches[0][0]}'")
                return self.classification_map[matches[0][0]]
        else:
            matches = get_close_matches(data_type, keys, n=1, cutoff=0.6)
            if matches:
                logging.info(f"Fuzzy match for '{data_type}' as '{matches[0]}'")
                return self.classification_map[matches[0]]
        logging.warning(f"No match for '{data_type}'")
        return "Unknown"

    def get_all_types(self) -> List[str]:
        """Return all data types."""
        return list(self.classification_map.keys())

    def get_levels(self) -> List[str]:
        """Return all classification levels."""
        return list(self.classification_levels)


# --- Theme Management ---
class ThemeManager:
    """Handles theme application for the GUI."""
    LIGHT = {
        "bg": "#FFFFFF",
        "fg": "#333",
        "Public": "green",
        "Internal": "#2057B3",
        "Confidential": "#D37D00",
        "Restricted": "#C00"
    }
    DARK = {
        "bg": "#232323",
        "fg": "#F0F0F0",
        "Public": "#55D46A",
        "Internal": "#6AC6FF",
        "Confidential": "#FFD16A",
        "Restricted": "#FF7575"
    }

    def __init__(self):
        self.theme = "light"

    def toggle(self):
        self.theme = "dark" if self.theme == "light" else "light"

    def get_color(self, level: str) -> str:
        palette = self.DARK if self.theme == "dark" else self.LIGHT
        return palette.get(level, palette["fg"])

    def get_bg(self) -> str:
        return self.DARK["bg"] if self.theme == "dark" else self.LIGHT["bg"]

    def is_dark(self) -> bool:
        return self.theme == "dark"


# --- Main App ---
class DataClassifierApp:
    """
    Tkinter GUI for classification tool.

    Args:
        root (tk.Tk): The root application window.
        cmap (ClassificationMap): The data classification logic/model.
    """
    def __init__(self, root: tk.Tk, cmap: ClassificationMap):
        self.root = root
        self.cmap = cmap
        self.theme_mgr = ThemeManager()
        self.root.title("Data Classification Tool")

        # Modern theme
        style = ttk.Style()
        style.theme_use('clam')
        self.setup_style(style)

        # Main layout
        self.main = ttk.Frame(root, padding="18 16 18 16")
        self.main.grid(row=0, column=0, sticky="NSEW")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        self.main.columnconfigure(0, weight=1)

        # Menu bar
        self.init_menu()

        # Search section
        row = 0
        ttk.Label(self.main, text="Search Data Type:").grid(row=row, column=0, sticky="W")
        row += 1

        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.update_dropdown)
        self.search_entry = ttk.Entry(self.main, textvariable=self.search_var, width=50)
        self.search_entry.grid(row=row, column=0, sticky="EW", pady=(0, 8))
        self.search_entry.bind("<Control-e>", lambda e: self.export_csv())  # Example shortcut
        self.add_tooltip(self.search_entry, "Type to search for a data type.")

        row += 1
        self.data_type_var = tk.StringVar()
        self.dropdown = ttk.Combobox(self.main, textvariable=self.data_type_var, width=60)
        self.dropdown['values'] = self.cmap.get_all_types()
        self.dropdown.grid(row=row, column=0, sticky="EW", pady=(0, 8))
        self.dropdown.bind("<<ComboboxSelected>>", self.classify_data)
        self.add_tooltip(self.dropdown, "Select a data type to see its classification.")

        row += 1
        self.result_label = ttk.Label(self.main, text="", font=("Arial", 12, "bold"), anchor="center")
        self.result_label.grid(row=row, column=0, pady=(0, 12), sticky="EW")

        row += 1
        ttk.Separator(self.main, orient="horizontal").grid(row=row, column=0, sticky="EW", pady=(0,12))

        row += 1

        # Add custom section
        add_frame = ttk.LabelFrame(self.main, text="Add Custom Data Type", padding="10 8 10 8")
        add_frame.grid(row=row, column=0, sticky="EW", padx=0, pady=(0,12))
        add_frame.columnconfigure(0, weight=1)
        self.new_data_var = tk.StringVar()
        self.new_classification_var = tk.StringVar()
        entry = ttk.Entry(add_frame, textvariable=self.new_data_var, width=40)
        entry.grid(row=0, column=0, sticky="EW", pady=(0, 8))
        self.add_tooltip(entry, "Enter a new data type.")

        self.new_class_dropdown = ttk.Combobox(
            add_frame, textvariable=self.new_classification_var,
            values=self.cmap.get_levels(), width=37
        )
        self.new_class_dropdown.grid(row=1, column=0, sticky="EW", pady=(0, 8))
        self.add_tooltip(self.new_class_dropdown, "Select a classification level.")

        btn = ttk.Button(add_frame, text="Add Data Type", command=self.add_custom_type)
        btn.grid(row=2, column=0, sticky="EW")
        self.add_tooltip(btn, "Add the custom data type to the list.")

        # Export/import section
        row += 1
        export_frame = ttk.Frame(self.main)
        export_frame.grid(row=row, column=0, sticky="EW")

        ttk.Button(export_frame, text="Export Classifications to CSV", command=self.export_csv).grid(row=0, column=0, sticky="EW", padx=(0,8))
        ttk.Button(export_frame, text="Import Classifications from CSV", command=self.import_csv).grid(row=0, column=1, sticky="EW")
        ttk.Button(export_frame, text="Export to Excel", command=self.export_excel).grid(row=1, column=0, sticky="EW", padx=(0,8), pady=(5,0))
        ttk.Button(export_frame, text="Import from Excel", command=self.import_excel).grid(row=1, column=1, sticky="EW", pady=(5,0))
        ttk.Button(export_frame, text="Export as JSON", command=self.export_json).grid(row=2, column=0, sticky="EW", padx=(0,8), pady=(5,0))
        ttk.Button(export_frame, text="Import from JSON", command=self.import_json).grid(row=2, column=1, sticky="EW", pady=(5,0))

        export_frame.columnconfigure(0, weight=1)
        export_frame.columnconfigure(1, weight=1)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.main, textvariable=self.status_var, anchor="w", relief="sunken")
        row += 1
        self.status_bar.grid(row=row, column=0, sticky="EW")
        self.set_status("Ready.")

        self.apply_theme()

    def setup_style(self, style: ttk.Style):
        """Setup ttk style for modern look/feel."""
        style.configure("TLabel", font=("Segoe UI", 11))
        style.configure("TButton", font=("Segoe UI", 10), padding=6)
        style.configure("TEntry", font=("Segoe UI", 10))
        style.configure("TCombobox", font=("Segoe UI", 10))
        style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"))

    def init_menu(self):
        """Initialize the menu bar."""
        menubar = tk.Menu(self.root)
        # Help
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="How to Use", command=self.show_help)
        menubar.add_cascade(label="Help", menu=helpmenu)
        # Theme
        thememenu = tk.Menu(menubar, tearoff=0)
        thememenu.add_command(label="Toggle Light/Dark", command=self.toggle_theme)
        menubar.add_cascade(label="Theme", menu=thememenu)
        self.root.config(menu=menubar)

    def update_dropdown(self, *args) -> None:
        """Update dropdown list based on search query (with fuzzy search)."""
        query = self.search_var.get().lower()
        all_types = self.cmap.get_all_types()
        if not query:
            filtered = all_types
        else:
            matches = [k for k in all_types if query in k.lower()]
            # Performance: use fuzzy match for large sets
            if FUZZY_LIB == 'rapidfuzz':
                fuzzy_matches = [m[0] for m in rapidfuzz_process.extract(query, all_types, limit=10, score_cutoff=50)]
            else:
                fuzzy_matches = get_close_matches(query, all_types, n=10, cutoff=0.5)
            filtered = list(dict.fromkeys(matches + fuzzy_matches))
        self.dropdown['values'] = filtered

    def classify_data(self, event=None) -> None:
        """Display the classification for the selected data type."""
        selected_data = self.data_type_var.get()
        classification = self.cmap.get_classification(selected_data)
        if classification == "Unknown":
            self.result_label.configure(text="Classification: Unknown", foreground="#555")
            self.set_status(f"Unknown classification for '{selected_data}'")
        else:
            color = self.theme_mgr.get_color(classification)
            self.result_label.configure(
                text=f"Classification: {classification}",
                foreground=color
            )
            self.set_status(f"Classified '{selected_data}' as '{classification}'.")

    def add_custom_type(self) -> None:
        """Add a custom data type and classification level."""
        new_data = self.new_data_var.get().strip()
        new_class = self.new_classification_var.get().strip()
        if not new_data or not new_class:
            self.set_status("Please enter a valid data type and classification level.", error=True)
            return
        if self.cmap.add_classification(new_data, new_class):
            self.dropdown['values'] = self.cmap.get_all_types()
            self.new_class_dropdown['values'] = self.cmap.get_levels()
            self.result_label.configure(
                text=f"Added: {new_data} as {new_class}",
                foreground=self.theme_mgr.get_color(new_class)
            )
            self.set_status(f"Added: '{new_data}' as '{new_class}'.")
            self.new_data_var.set("")
            self.new_classification_var.set("")
        else:
            self.set_status("Failed to add the data type.", error=True)

    def export_csv(self) -> None:
        """Export classification data to a CSV file."""
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                self.cmap.export_csv(file_path)
                self.set_status(f"Classification data saved to {file_path}")
            except IOError as e:
                self.set_status(f"Failed to save file: {e}", error=True)

    def import_csv(self) -> None:
        """Import classification data from a CSV file."""
        file_path = filedialog.askopenfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                self.cmap.import_csv(file_path)
                self.dropdown['values'] = self.cmap.get_all_types()
                self.new_class_dropdown['values'] = self.cmap.get_levels()
                self.set_status(f"Classification data loaded from {file_path}")
            except Exception as e:
                self.set_status(f"Failed to import file: {e}", error=True)

    def export_excel(self) -> None:
        """Export classification data to an Excel file."""
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.cmap.export_excel(file_path)
                self.set_status(f"Classification data saved to {file_path}")
            except Exception as e:
                self.set_status(f"Failed to save file: {e}", error=True)

    def import_excel(self) -> None:
        """Import classification data from an Excel file."""
        file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.cmap.import_excel(file_path)
                self.dropdown['values'] = self.cmap.get_all_types()
                self.new_class_dropdown['values'] = self.cmap.get_levels()
                self.set_status(f"Classification data loaded from {file_path}")
            except Exception as e:
                self.set_status(f"Failed to import file: {e}", error=True)

    def export_json(self) -> None:
        """Export classification data to a JSON file."""
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                self.cmap.export_json(file_path)
                self.set_status(f"Classification data saved to {file_path}")
            except Exception as e:
                self.set_status(f"Failed to save file: {e}", error=True)

    def import_json(self) -> None:
        """Import classification data from a JSON file."""
        file_path = filedialog.askopenfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                self.cmap.import_json(file_path)
                self.dropdown['values'] = self.cmap.get_all_types()
                self.new_class_dropdown['values'] = self.cmap.get_levels()
                self.set_status(f"Classification data loaded from {file_path}")
            except Exception as e:
                self.set_status(f"Failed to import file: {e}", error=True)

    def show_help(self) -> None:
        """Display help message."""
        help_text = (
            "How to Use the Data Classification Tool:\n"
            "- Search for a data type using the search box.\n"
            "- Select a data type to see its classification.\n"
            "- Add a custom data type and classification below.\n"
            "- Use the buttons at the bottom to import/export data (CSV, Excel, or JSON).\n"
            "- Use the Theme menu to toggle light/dark mode.\n"
            "- All actions are logged to 'dataclassifier.log'.\n"
            "- Keyboard shortcut: Ctrl+E to export as CSV."
        )
        messagebox.showinfo("Help", help_text)

    def toggle_theme(self) -> None:
        """Toggle between light and dark themes."""
        self.theme_mgr.toggle()
        self.apply_theme()

    def apply_theme(self) -> None:
        """Apply the current theme to the GUI."""
        bg = self.theme_mgr.get_bg()
        self.root.configure(bg=bg)
        self.main.configure(style="Dark.TFrame" if self.theme_mgr.is_dark() else "TFrame")
        for widget in self.root.winfo_children():
            try:
                widget.configure(bg=bg)
            except Exception:
                pass

    def set_status(self, message: str, error: bool = False) -> None:
        """Set the status bar message."""
        self.status_var.set(message)
        if error:
            self.status_bar.configure(foreground="red")
            logging.error(message)
        else:
            self.status_bar.configure(foreground="green")
            logging.info(message)

    def add_tooltip(self, widget, text: str) -> None:
        """Add a tooltip to a widget."""
        def on_enter(event):
            x, y, cx, cy = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 20
            self.tipwindow = tw = tk.Toplevel(widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            label = tk.Label(tw, text=text, background="#ffffe0", relief='solid', borderwidth=1, font=("tahoma", "8", "normal"))
            label.pack(ipadx=1)
        def on_leave(event):
            if hasattr(self, 'tipwindow') and self.tipwindow:
                self.tipwindow.destroy()
                self.tipwindow = None
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

if __name__ == "__main__":
    root = tk.Tk()
    cmap = ClassificationMap()
    app = DataClassifierApp(root, cmap)
    root.mainloop()
