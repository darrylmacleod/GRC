def calculate_risk(likelihood, impact):
    """Calculate risk score based on likelihood and impact"""
    risk_matrix = {
        'Low': {'Low': 1, 'Medium': 2, 'High': 3},
        'Medium': {'Low': 2, 'Medium': 4, 'High': 6},
        'High': {'Low': 3, 'Medium': 6, 'High': 9}
    }
    return risk_matrix.get(likelihood, {}).get(impact, 0)

def assess_risks(risks):
    """Assess a list of risks and prioritize them"""
    assessed_risks = []
    for risk in risks:
        score = calculate_risk(risk['likelihood'], risk['impact'])
        assessed_risks.append({
            'name': risk['name'],
            'score': score,
            'priority': 'High' if score > 5 else ('Medium' if score > 2 else 'Low')
        })
    return sorted(assessed_risks, key=lambda x: x['score'], reverse=True)

# Example usage
risks = [
    {'name': 'Data breach', 'likelihood': 'Medium', 'impact': 'High'},
    {'name': 'System downtime', 'likelihood': 'Low', 'impact': 'High'}
]

prioritized_risks = assess_risks(risks)
print("Prioritized Risks:", prioritized_risks)
