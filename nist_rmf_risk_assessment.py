import pandas as pd
from datetime import datetime

# Define risk calculation
def calculate_risk(impact, likelihood):
    score_map = {"Low": 1, "Moderate": 2, "High": 3}
    return score_map[impact] * score_map[likelihood]

# Step 1: Categorize System
def categorize_system():
    print("=== STEP 1: Categorize the System ===")
    system_name = input("System Name: ")
    system_description = input("System Description: ")
    sensitivity = input("Data Sensitivity (Low/Moderate/High): ")
    criticality = input("System Criticality (Low/Moderate/High): ")
    return {
        "System Name": system_name,
        "Description": system_description,
        "Sensitivity": sensitivity,
        "Criticality": criticality,
    }

# Step 2: Identify Threats and Vulnerabilities
def identify_threats():
    print("\n=== STEP 2: Identify Threats and Vulnerabilities ===")
    threats = []
    while True:
        threat = input("Enter a threat (or type 'done' to finish): ")
        if threat.lower() == 'done':
            break
        vuln = input(f"→ Associated vulnerability for '{threat}': ")
        impact = input("→ Impact if exploited (Low/Moderate/High): ")
        likelihood = input("→ Likelihood of occurrence (Low/Moderate/High): ")
        mitigation = input("→ Existing mitigation in place: ")
        risk_score = calculate_risk(impact, likelihood)
        threats.append({
            "Threat": threat,
            "Vulnerability": vuln,
            "Impact": impact,
            "Likelihood": likelihood,
            "Mitigation": mitigation,
            "Risk Score": risk_score
        })
    return threats

# Step 3: Report Risk Register
def create_risk_register(threats):
    print("\n=== STEP 3: Risk Register Summary ===")
    df = pd.DataFrame(threats)
    df["Risk Level"] = df["Risk Score"].apply(
        lambda x: "Low" if x <= 2 else "Moderate" if x <= 4 else "High"
    )
    print(df[["Threat", "Risk Level", "Mitigation"]])
    return df

# Step 4: Save Results
def export_results(system_info, df):
    print("\n=== STEP 4: Exporting Risk Report ===")
    filename = f"risk_assessment_{system_info['System Name'].replace(' ', '_')}.xlsx"
    with pd.ExcelWriter(filename) as writer:
        pd.DataFrame([system_info]).to_excel(writer, sheet_name="System Info", index=False)
        df.to_excel(writer, sheet_name="Risk Register", index=False)
    print(f"✅ Risk assessment exported to: {filename}")

# === MAIN FLOW ===
if __name__ == "__main__":
    print("NIST RMF Risk Assessment Tool\n")
    system_info = categorize_system()
    threats = identify_threats()
    risk_df = create_risk_register(threats)
    export_results(system_info, risk_df)
