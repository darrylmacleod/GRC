import logging
from typing import List, Dict, Any, Optional
from functools import partial
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.simpledialog import askstring
from dataclasses import dataclass, asdict, field
import re
import json

# Configure logging with timestamps and log levels
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Default risk levels (modifiable at runtime)
DEFAULT_RISK_LEVELS = ['Low', 'Medium', 'High']

# Default risk matrix and priority thresholds
DEFAULT_RISK_MATRIX = {
    'Low': {'Low': 1, 'Medium': 2, 'High': 3},
    'Medium': {'Low': 2, 'Medium': 4, 'High': 6},
    'High': {'Low': 3, 'Medium': 6, 'High': 9}
}
PRIORITY_THRESHOLDS = {'High': 5, 'Medium': 2}

@dataclass
class Risk:
    """
    Represents a risk item.

    Args:
        name (str): Name of the risk
        likelihood (str): Likelihood value
        impact (str): Impact value
        score (int): Calculated risk score
        priority (str): Calculated priority
    """
    name: str
    likelihood: str
    impact: str
    score: int = 0
    priority: str = "Low"

def validate_input(value: str, valid_options: List[str], field_name: str = "Value") -> None:
    """Validate that a value is within the valid options."""
    if value not in valid_options:
        raise ValueError(f"{field_name} '{value}' is invalid. Must be one of {valid_options}.")

def validate_risk_name(name: str, existing_names: List[str]) -> None:
    """Ensure risk name is non-empty, unique, and contains only alphanumeric and spaces."""
    if not name or not name.strip():
        raise ValueError("Risk name cannot be empty.")
    if not re.match(r'^[\w\s\-]+$', name):
        raise ValueError("Risk name must contain only letters, numbers, spaces, or hyphens.")
    if name in existing_names:
        raise ValueError(f"Risk name '{name}' already exists.")

def calculate_risk(
    likelihood: str,
    impact: str,
    risk_matrix: Dict[str, Dict[str, int]] = DEFAULT_RISK_MATRIX,
    risk_levels: List[str] = DEFAULT_RISK_LEVELS
) -> int:
    """
    Calculate risk score based on likelihood and impact.

    Example:
        calculate_risk('High', 'Medium')
    """
    validate_input(likelihood, risk_levels, "Likelihood")
    validate_input(impact, risk_levels, "Impact")
    try:
        return risk_matrix[likelihood][impact]
    except KeyError:
        raise ValueError(f"Invalid likelihood '{likelihood}' or impact '{impact}' for selected matrix.")

def calculate_priority(
    score: int,
    thresholds: Dict[str, int] = PRIORITY_THRESHOLDS
) -> str:
    """
    Assign a priority based on the risk score.

    Example:
        calculate_priority(7)
    """
    if score > thresholds['High']:
        return 'High'
    elif score > thresholds['Medium']:
        return 'Medium'
    else:
        return 'Low'

def assess_risks(
    risks: List[Risk],
    risk_matrix: Dict[str, Dict[str, int]] = DEFAULT_RISK_MATRIX,
    thresholds: Dict[str, int] = PRIORITY_THRESHOLDS,
    risk_levels: List[str] = DEFAULT_RISK_LEVELS
) -> List[Risk]:
    """
    Assess a list of risks and prioritize them.

    Example:
        assess_risks([Risk("Test", "Low", "High")])
    """
    assessed_risks = []
    for risk in risks:
        try:
            score = calculate_risk(risk.likelihood, risk.impact, risk_matrix, risk_levels)
            priority = calculate_priority(score, thresholds)
            assessed = Risk(risk.name, risk.likelihood, risk.impact, score, priority)
            assessed_risks.append(assessed)
        except (KeyError, ValueError) as e:
            logging.error(f"Error assessing risk '{risk.name}': {e}")
    result = sorted(assessed_risks, key=lambda x: x.score, reverse=True)
    logging.info("Risks assessed: %s", [asdict(r) for r in result])
    return result

# ---- GUI Section ----

class RiskAssessmentGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Risk Assessment Calculator")
        self.risk_levels = list(DEFAULT_RISK_LEVELS)
        self.risk_matrix = dict(DEFAULT_RISK_MATRIX)
        self.thresholds = dict(PRIORITY_THRESHOLDS)
        self.risks: List[Risk] = []

        self.name_var = tk.StringVar()
        self.likelihood_var = tk.StringVar(value=self.risk_levels[0])
        self.impact_var = tk.StringVar(value=self.risk_levels[0])

        frame = tk.Frame(root)
        frame.pack(padx=10, pady=10)

        tk.Label(frame, text="Risk Name:").grid(row=0, column=0, sticky="e")
        self.name_entry = tk.Entry(frame, textvariable=self.name_var, width=25)
        self.name_entry.grid(row=0, column=1, padx=5)
        self.name_entry.bind("<FocusIn>", lambda e: self.show_tooltip("Enter a unique, descriptive risk name."))

        tk.Label(frame, text="Likelihood:").grid(row=1, column=0, sticky="e")
        self.likelihood_combo = ttk.Combobox(frame, textvariable=self.likelihood_var, values=self.risk_levels, state="readonly", width=15)
        self.likelihood_combo.grid(row=1, column=1)
        self.likelihood_combo.bind("<FocusIn>", lambda e: self.show_tooltip("Likelihood: How probable is the risk?"))

        tk.Label(frame, text="Impact:").grid(row=2, column=0, sticky="e")
        self.impact_combo = ttk.Combobox(frame, textvariable=self.impact_var, values=self.risk_levels, state="readonly", width=15)
        self.impact_combo.grid(row=2, column=1)
        self.impact_combo.bind("<FocusIn>", lambda e: self.show_tooltip("Impact: How severe would the effect be?"))

        tk.Button(frame, text="Add Risk", command=self.add_risk).grid(row=3, column=0, columnspan=2, pady=5)
        tk.Button(frame, text="Edit Risk Levels", command=self.edit_risk_levels).grid(row=4, column=0, columnspan=2)
        tk.Button(frame, text="Save Risks", command=self.save_risks).grid(row=5, column=0)
        tk.Button(frame, text="Load Risks", command=self.load_risks).grid(row=5, column=1)

        # Risk list
        self.tree = ttk.Treeview(root, columns=("Name", "Likelihood", "Impact"), show="headings", height=6)
        self.tree.heading("Name", text="Risk Name")
        self.tree.heading("Likelihood", text="Likelihood")
        self.tree.heading("Impact", text="Impact")
        self.tree.pack(padx=10, pady=5, fill="x")
        self.tree.bind("<Delete>", self.delete_selected_risk)
        self.tree.bind("<Return>", self.edit_selected_risk)

        tk.Button(root, text="Assess Risks", command=self.display_results).pack(pady=10)
        
        self.results_text = tk.Text(root, height=8, width=60, state="disabled")
        self.results_text.pack(padx=10, pady=5)

        self.tooltip = tk.Label(root, text="", background="yellow", relief="solid", borderwidth=1)
        self.tooltip.pack_forget()

        self.root.bind("<Tab>", self.focus_next_widget)
        self.root.bind("<Shift-Tab>", self.focus_prev_widget)

    def show_tooltip(self, text: str):
        self.tooltip.config(text=text)
        self.tooltip.place(x=10, y=5)
        self.tooltip.lift()
        self.root.after(2000, self.tooltip.pack_forget)

    def focus_next_widget(self, event):
        event.widget.tk_focusNext().focus()
        return "break"

    def focus_prev_widget(self, event):
        event.widget.tk_focusPrev().focus()
        return "break"

    def get_existing_names(self) -> List[str]:
        return [risk.name for risk in self.risks]

    def add_risk(self):
        name = self.name_var.get().strip()
        likelihood = self.likelihood_var.get()
        impact = self.impact_var.get()
        try:
            validate_risk_name(name, self.get_existing_names())
            validate_input(likelihood, self.risk_levels, "Likelihood")
            validate_input(impact, self.risk_levels, "Impact")
        except ValueError as e:
            logging.warning(f"Add risk validation failed: {e}")
            messagebox.showerror("Input Error", str(e))
            return
        new_risk = Risk(name, likelihood, impact)
        self.risks.append(new_risk)
        self.tree.insert("", "end", values=(name, likelihood, impact))
        self.name_var.set("")

    def delete_selected_risk(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
        for item in selected:
            name = self.tree.item(item, "values")[0]
            self.risks = [risk for risk in self.risks if risk.name != name]
            self.tree.delete(item)

    def edit_selected_risk(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
        item = selected[0]
        values = self.tree.item(item, "values")
        name = values[0]
        risk = next((r for r in self.risks if r.name == name), None)
        if not risk:
            return
        # Let user edit likelihood and impact
        new_likelihood = askstring("Edit Likelihood", f"Current: {risk.likelihood}\nEnter new Likelihood ({', '.join(self.risk_levels)}):")
        new_impact = askstring("Edit Impact", f"Current: {risk.impact}\nEnter new Impact ({', '.join(self.risk_levels)}):")
        try:
            if new_likelihood:
                validate_input(new_likelihood, self.risk_levels, "Likelihood")
                risk.likelihood = new_likelihood
            if new_impact:
                validate_input(new_impact, self.risk_levels, "Impact")
                risk.impact = new_impact
            self.tree.item(item, values=(risk.name, risk.likelihood, risk.impact))
        except ValueError as e:
            messagebox.showerror("Edit Error", str(e))

    def edit_risk_levels(self):
        new_levels = askstring("Edit Risk Levels", f"Current: {', '.join(self.risk_levels)}\nEnter new levels, comma-separated:")
        if new_levels:
            levels = [lvl.strip() for lvl in new_levels.split(",") if lvl.strip()]
            if len(levels) < 2:
                messagebox.showerror("Edit Error", "At least two risk levels are required.")
                return
            self.risk_levels = levels
            # Update combo boxes
            self.likelihood_combo['values'] = self.risk_levels
            self.impact_combo['values'] = self.risk_levels
            # NOTE: For full customization, user should also update risk matrix via future dialog

    def display_results(self):
        if not self.risks:
            messagebox.showinfo("No Risks", "Please add at least one risk.")
            return
        results = assess_risks(self.risks, self.risk_matrix, self.thresholds, self.risk_levels)
        self.results_text.configure(state="normal")
        self.results_text.delete(1.0, tk.END)
        for risk in results:
            self.results_text.insert(
                tk.END, f"Name: {risk.name}, Score: {risk.score}, Priority: {risk.priority}\n"
            )
        self.results_text.configure(state="disabled")

    def save_risks(self):
        try:
            with open("risks.json", "w") as f:
                json.dump([asdict(risk) for risk in self.risks], f, indent=2)
            messagebox.showinfo("Save Risks", "Risks saved successfully.")
        except Exception as e:
            logging.error(f"Error saving risks: {e}")
            messagebox.showerror("Save Error", str(e))

    def load_risks(self):
        try:
            with open("risks.json", "r") as f:
                risk_dicts = json.load(f)
            self.risks = [Risk(**rd) for rd in risk_dicts]
            self.tree.delete(*self.tree.get_children())
            for risk in self.risks:
                self.tree.insert("", "end", values=(risk.name, risk.likelihood, risk.impact))
            messagebox.showinfo("Load Risks", "Risks loaded successfully.")
        except Exception as e:
            logging.error(f"Error loading risks: {e}")
            messagebox.showerror("Load Error", str(e))

if __name__ == '__main__':
    root = tk.Tk()
    app = RiskAssessmentGUI(root)
    root.mainloop()

"""
# Unit Test Scaffold (to be placed in a separate test file)

import unittest

class TestRiskAssessment(unittest.TestCase):
    def test_calculate_risk(self):
        self.assertEqual(calculate_risk('Low', 'Low'), 1)
        self.assertRaises(ValueError, calculate_risk, 'X', 'Low')
    def test_priority(self):
        self.assertEqual(calculate_priority(7), 'High')
        self.assertEqual(calculate_priority(4), 'Medium')
        self.assertEqual(calculate_priority(1), 'Low')
    def test_validation(self):
        self.assertRaises(ValueError, validate_risk_name, '', [])
        self.assertRaises(ValueError, validate_risk_name, 'Invalid!', [])
        self.assertRaises(ValueError, validate_risk_name, 'Dup', ['Dup'])

if __name__ == "__main__":
    unittest.main()
"""
