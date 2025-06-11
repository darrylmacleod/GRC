import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv

# Initial classification map
classification_map = {
    # Public
    "Company website content": "Public",
    "Marketing materials": "Public",
    "Press releases": "Public",
    "Published research papers": "Public",
    "Public regulatory filings": "Public",

    # Internal
    "Internal memos and communications": "Internal",
    "Organization charts": "Internal",
    "Training materials": "Internal",
    "Project plans and status reports": "Internal",
    "Meeting agendas and minutes (non-sensitive)": "Internal",

    # Confidential
    "Employee ID numbers": "Confidential",
    "Business strategies": "Confidential",
    "Contract details": "Confidential",
    "Internal financial statements": "Confidential",
    "Customer contact information": "Confidential",
    "Source code": "Confidential",
    "Non-public pricing or product roadmaps": "Confidential",
    "Vendor agreements": "Confidential",
    "Intellectual property documentation": "Confidential",

    # Restricted
    "Full names, addresses, phone numbers": "Restricted",
    "Social Insurance Numbers (SIN)/Social Security Numbers (SSN)": "Restricted",
    "Driver’s license numbers": "Restricted",
    "Dates of birth": "Restricted",
    "Passport numbers": "Restricted",
    "Medical records": "Restricted",
    "Health insurance data": "Restricted",
    "Lab test results": "Restricted",
    "Appointment histories": "Restricted",
    "Credit card numbers": "Restricted",
    "CVV codes": "Restricted",
    "Cardholder names and billing addresses": "Restricted",
    "Bank account numbers": "Restricted",
    "Routing numbers": "Restricted",
    "Tax returns": "Restricted",
    "Passwords": "Restricted",
    "API keys": "Restricted",
    "Encryption keys": "Restricted",
    "Biometric data": "Restricted",
    "Litigation documents": "Restricted",
    "Legal holds": "Restricted",
    "Regulatory investigation materials": "Restricted",
    "Board meeting minutes": "Restricted",
    "Merger/acquisition plans": "Restricted",
    "Due diligence documents": "Restricted",
}

# GUI App
class DataClassifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Classification Tool")
        self.filtered_values = list(classification_map.keys())

        # Search box
        tk.Label(root, text="Search Data Type:", font=("Arial", 10)).pack(pady=(10, 0))
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.update_dropdown)
        self.search_entry = tk.Entry(root, textvariable=self.search_var, width=50)
        self.search_entry.pack(pady=5)

        # Dropdown
        self.data_type_var = tk.StringVar()
        self.dropdown = ttk.Combobox(root, textvariable=self.data_type_var, width=60)
        self.dropdown['values'] = self.filtered_values
        self.dropdown.pack(pady=5)
        self.dropdown.bind("<<ComboboxSelected>>", self.classify_data)

        # Result
        self.result_label = tk.Label(root, text="", font=("Arial", 14, "bold"))
        self.result_label.pack(pady=10)

        # Add new entry section
        tk.Label(root, text="Add Custom Data Type", font=("Arial", 10)).pack(pady=(10, 0))
        self.new_data_var = tk.StringVar()
        self.new_classification_var = tk.StringVar()
        tk.Entry(root, textvariable=self.new_data_var, width=50).pack(pady=2)
        ttk.Combobox(root, textvariable=self.new_classification_var, values=["Public", "Internal", "Confidential", "Restricted"], width=47).pack(pady=2)
        tk.Button(root, text="Add Data Type", command=self.add_custom_type).pack(pady=5)

        # Export button
        tk.Button(root, text="Export Classifications to CSV", command=self.export_csv).pack(pady=(10, 5))

    def update_dropdown(self, *args):
        query = self.search_var.get().lower()
        self.filtered_values = [key for key in classification_map.keys() if query in key.lower()]
        self.dropdown['values'] = self.filtered_values

    def classify_data(self, event=None):
        selected_data = self.data_type_var.get()
        classification = classification_map.get(selected_data, "Unknown")
        self.result_label.config(
            text=f"Classification: {classification}",
            fg=self.get_color(classification)
        )

    def add_custom_type(self):
        new_data = self.new_data_var.get().strip()
        new_class = self.new_classification_var.get().strip()

        if not new_data or new_class not in ["Public", "Internal", "Confidential", "Restricted"]:
            messagebox.showerror("Invalid Input", "Please enter a valid data type and classification level.")
            return

        classification_map[new_data] = new_class
        self.update_dropdown()
        self.result_label.config(text=f"Added: {new_data} as {new_class}", fg=self.get_color(new_class))
        self.new_data_var.set("")
        self.new_classification_var.set("")

    def export_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            with open(file_path, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(["Data Type", "Classification"])
                for data_type, level in classification_map.items():
                    writer.writerow([data_type, level])
            messagebox.showinfo("Exported", f"Classification data saved to {file_path}")

    def get_color(self, level):
        return {
            "Public": "green",
            "Internal": "blue",
            "Confidential": "orange",
            "Restricted": "red"
        }.get(level, "black")

# Run app
if __name__ == "__main__":
    root = tk.Tk()
    app = DataClassifierApp(root)
    root.mainloop()
