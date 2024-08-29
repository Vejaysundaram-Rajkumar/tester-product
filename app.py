# app.py
from flask import Flask, render_template, request

app = Flask(__name__)

# Example predefined goals for each role
roles_and_goals = {
    'CEO': ['Revenue Growth', 'Market Expansion', 'Customer Satisfaction', 'Employee Retention', 'Innovation'],
    'CFO': ['Cost Reduction', 'Profit Margin', 'Cash Flow', 'Debt Management', 'Financial Reporting'],
    'COO': ['Operational Efficiency', 'Supply Chain Management', 'Quality Control', 'Production Volume', 'Safety Standards']
}

@app.route('/', methods=['GET', 'POST'])
def company_form():
    if request.method == 'POST':
        # Process basic company details
        company_name = request.form.get('company_name')
        market_value = request.form.get('market_value')
        location = request.form.get('location')
        founder_name = request.form.get('founder_name')
        business_type = request.form.get('business_type')

        # Check if we are processing goals (second step)
        if 'submit_goals' in request.form:
            selected_goals = {}
            for role, goals in roles_and_goals.items():
                for goal in goals:
                    if request.form.get(f'{role}_{goal}'):
                        target_value = request.form.get(f'{role}_{goal}_target')
                        actual_value = request.form.get(f'{role}_{goal}_actual')
                        try:
                            gap = abs(float(target_value) - float(actual_value))
                        except ValueError:
                            gap = 'N/A'  # In case of non-numeric input
                        selected_goals[goal] = {'target': target_value, 'actual': actual_value, 'gap': gap}
            
            # Render the results page with company details and gap analysis
            return render_template('results.html', company_name=company_name, market_value=market_value,
                                   location=location, founder_name=founder_name, business_type=business_type,
                                   selected_goals=selected_goals)

        # Proceed to display goals section
        return render_template('index.html', step='goals', company_name=company_name, market_value=market_value,
                               location=location, founder_name=founder_name, business_type=business_type,
                               roles_and_goals=roles_and_goals)

    return render_template('index.html', step='details', roles_and_goals=roles_and_goals)

if __name__ == '__main__':
    app.run(host='0.0.0.0')
