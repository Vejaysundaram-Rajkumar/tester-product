from flask import Flask, render_template, request

app = Flask(__name__)

# Example predefined metrics for each category
categories_and_metrics = {
    'Category 1': ['Metric 1-1', 'Metric 1-2', 'Metric 1-3', 'Metric 1-4', 'Metric 1-5', 'Metric 1-6', 'Metric 1-7', 'Metric 1-8', 'Metric 1-9', 'Metric 1-10'],
    'Category 2': ['Metric 2-1', 'Metric 2-2', 'Metric 2-3', 'Metric 2-4', 'Metric 2-5', 'Metric 2-6', 'Metric 2-7', 'Metric 2-8', 'Metric 2-9', 'Metric 2-10'],
    'Category 3': ['Metric 3-1', 'Metric 3-2', 'Metric 3-3', 'Metric 3-4', 'Metric 3-5', 'Metric 3-6', 'Metric 3-7', 'Metric 3-8', 'Metric 3-9', 'Metric 3-10'],
    'Category 4': ['Metric 4-1', 'Metric 4-2', 'Metric 4-3', 'Metric 4-4', 'Metric 4-5', 'Metric 4-6', 'Metric 4-7', 'Metric 4-8', 'Metric 4-9', 'Metric 4-10'],
    'Category 5': ['Metric 5-1', 'Metric 5-2', 'Metric 5-3', 'Metric 5-4', 'Metric 5-5', 'Metric 5-6', 'Metric 5-7', 'Metric 5-8', 'Metric 5-9', 'Metric 5-10'],
    'Category 6': ['Metric 6-1', 'Metric 6-2', 'Metric 6-3', 'Metric 6-4', 'Metric 6-5', 'Metric 6-6', 'Metric 6-7', 'Metric 6-8', 'Metric 6-9', 'Metric 6-10'],
    'Category 7': ['Metric 7-1', 'Metric 7-2', 'Metric 7-3', 'Metric 7-4', 'Metric 7-5', 'Metric 7-6', 'Metric 7-7', 'Metric 7-8', 'Metric 7-9', 'Metric 7-10'],
    'Category 8': ['Metric 8-1', 'Metric 8-2', 'Metric 8-3', 'Metric 8-4', 'Metric 8-5', 'Metric 8-6', 'Metric 8-7', 'Metric 8-8', 'Metric 8-9', 'Metric 8-10'],
    'Category 9': ['Metric 9-1', 'Metric 9-2', 'Metric 9-3', 'Metric 9-4', 'Metric 9-5', 'Metric 9-6', 'Metric 9-7', 'Metric 9-8', 'Metric 9-9', 'Metric 9-10']
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
            selected_metrics = {}
            selected_metrics_count = 0

            # Process selected metrics and gather their target/actual values
            for category, metrics in categories_and_metrics.items():
                for metric in metrics:
                    if request.form.get(f'{category}_{metric}'):
                        target_value = request.form.get(f'{category}_{metric}_target')
                        actual_value = request.form.get(f'{category}_{metric}_actual')
                        try:
                            gap = abs(float(target_value) - float(actual_value))
                        except ValueError:
                            gap = 'N/A'  # In case of non-numeric input
                        selected_metrics[metric] = {'target': target_value, 'actual': actual_value, 'gap': gap}
                        selected_metrics_count += 1
                        if selected_metrics_count >= 10:
                            break
                if selected_metrics_count >= 10:
                    break

            # Render the results page with company details and gap analysis
            return render_template('results.html', company_name=company_name, market_value=market_value,
                                   location=location, founder_name=founder_name, business_type=business_type,
                                   selected_metrics=selected_metrics)

        # Proceed to display goals section
        return render_template('index.html', step='goals', company_name=company_name, market_value=market_value,
                               location=location, founder_name=founder_name, business_type=business_type,
                               categories_and_metrics=categories_and_metrics)

    return render_template('index.html', step='details', categories_and_metrics=categories_and_metrics)

if __name__ == '__main__':
    app.run(host='0.0.0.0')
