import logging
from typing import List, Dict, Any
from functools import partial
import tkinter as tk
from tkinter import ttk, messagebox

# Configure logging
logging.basicConfig(level=logging.INFO)

# Constants for risk levels
RISK_LEVELS = ['Low', 'Medium', 'High']

# Default risk matrix and priority thresholds
DEFAULT_RISK_MATRIX = {
    'Low': {'Low': 1, 'Medium': 2, 'High': 3},
    'Medium': {'Low': 2, 'Medium': 4, 'High': 6},
    'High': {'Low': 3, 'Medium': 6, 'High': 9}
}
PRIORITY_THRESHOLDS = {'High': 5, 'Medium': 2}


def validate_input(value: str, valid_options: List[str]) -> None:
    """
    Validate that a value is within the valid options.

    Args:
        value (str): The value to validate.
        valid_options (List[str]): Allowed options.

    Raises:
        ValueError: If value is not in valid_options.
    """
    if value not in valid_options:
        raise ValueError(f"Invalid value '{value}'. Must be one of {valid_options}.")


def calculate_risk(
    likelihood: str,
    impact: str,
    risk_matrix: Dict[str, Dict[str, int]] = DEFAULT_RISK_MATRIX
) -> int:
    """
    Calculate risk score based on likelihood and impact.

    Args:
        likelihood (str): Risk likelihood ('Low', 'Medium', 'High').
        impact (str): Risk impact ('Low', 'Medium', 'High').
        risk_matrix (dict): Optional custom risk matrix.

    Returns:
        int: Risk score.

    Raises:
        ValueError: If likelihood or impact are invalid.
    """
    validate_input(likelihood, RISK_LEVELS)
    validate_input(impact, RISK_LEVELS)
    try:
        return risk_matrix[likelihood][impact]
    except KeyError:
        raise ValueError(f"Invalid likelihood '{likelihood}' or impact '{impact}'.")


def calculate_priority(
    score: int,
    thresholds: Dict[str, int] = PRIORITY_THRESHOLDS
) -> str:
    """
    Assign a priority based on the risk score.

    Args:
        score (int): The risk score.
        thresholds (dict): Priority thresholds.

    Returns:
        str: Priority level ('Low', 'Medium', 'High').
    """
    if score > thresholds['High']:
        return 'High'
    elif score > thresholds['Medium']:
        return 'Medium'
    else:
        return 'Low'


def assess_risks(
    risks: List[Dict[str, Any]],
    risk_matrix: Dict[str, Dict[str, int]] = DEFAULT_RISK_MATRIX,
    thresholds: Dict[str, int] = PRIORITY_THRESHOLDS
) -> List[Dict[str, Any]]:
    """
    Assess a list of risks and prioritize them.

    Args:
        risks (List[Dict[str, Any]]): List of risks, each a dict with 'name', 'likelihood', 'impact'.
        risk_matrix (dict): Optional custom risk matrix.
        thresholds (dict): Optional custom priority thresholds.

    Returns:
        List[Dict[str, Any]]: Risks with calculated score and priority, sorted by score descending.
    """
    assessed_risks = []
    for risk in risks:
        try:
            score = calculate_risk(risk['likelihood'], risk['impact'], risk_matrix)
            priority = calculate_priority(score, thresholds)
            assessed_risks.append({
                'name': risk['name'],
                'score': score,
                'priority': priority
            })
        except (KeyError, ValueError) as e:
            logging.error(f"Error assessing risk '{risk.get('name', 'Unknown')}': {e}")

    sort_key = partial(lambda x: x['score'])
    result = sorted(assessed_risks, key=sort_key, reverse=True)
    logging.info("Risks assessed: %s", result)
    return result


# ---- GUI Section ----

class RiskAssessmentGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Risk Assessment Calculator")

        self.risks: List[Dict[str, Any]] = []

        # Input fields
        self.name_var = tk.StringVar()
        self.likelihood_var = tk.StringVar(value=RISK_LEVELS[0])
        self.impact_var = tk.StringVar(value=RISK_LEVELS[0])

        frame = tk.Frame(root)
        frame.pack(padx=10, pady=10)

        tk.Label(frame, text="Risk Name:").grid(row=0, column=0, sticky="e")
        tk.Entry(frame, textvariable=self.name_var, width=25).grid(row=0, column=1, padx=5)

        tk.Label(frame, text="Likelihood:").grid(row=1, column=0, sticky="e")
        ttk.Combobox(frame, textvariable=self.likelihood_var, values=RISK_LEVELS, state="readonly", width=15).grid(row=1, column=1)

        tk.Label(frame, text="Impact:").grid(row=2, column=0, sticky="e")
        ttk.Combobox(frame, textvariable=self.impact_var, values=RISK_LEVELS, state="readonly", width=15).grid(row=2, column=1)

        tk.Button(frame, text="Add Risk", command=self.add_risk).grid(row=3, column=0, columnspan=2, pady=5)

        # Risk list
        self.tree = ttk.Treeview(root, columns=("Name", "Likelihood", "Impact"), show="headings", height=6)
        self.tree.heading("Name", text="Risk Name")
        self.tree.heading("Likelihood", text="Likelihood")
        self.tree.heading("Impact", text="Impact")
        self.tree.pack(padx=10, pady=5, fill="x")

        # Assess button
        tk.Button(root, text="Assess Risks", command=self.display_results).pack(pady=10)

        # Results
        self.results_text = tk.Text(root, height=8, width=60, state="disabled")
        self.results_text.pack(padx=10, pady=5)

    def add_risk(self):
        name = self.name_var.get().strip()
        likelihood = self.likelihood_var.get()
        impact = self.impact_var.get()
        if not name:
            messagebox.showerror("Input Error", "Risk name cannot be empty.")
            return
        self.risks.append({'name': name, 'likelihood': likelihood, 'impact': impact})
        self.tree.insert("", "end", values=(name, likelihood, impact))
        self.name_var.set("")

    def display_results(self):
        if not self.risks:
            messagebox.showinfo("No Risks", "Please add at least one risk.")
            return
        results = assess_risks(self.risks)
        self.results_text.configure(state="normal")
        self.results_text.delete(1.0, tk.END)
        for risk in results:
            self.results_text.insert(tk.END, f"Name: {risk['name']}, Score: {risk['score']}, Priority: {risk['priority']}\n")
        self.results_text.configure(state="disabled")


if __name__ == '__main__':
    root = tk.Tk()
    app = RiskAssessmentGUI(root)
    root.mainloop()
