from flask import Flask, render_template, request
import pandas as pd
import json

app = Flask(__name__)

# Load the list of metrics from the master sheet and the details from metrics sheet.
def load_metrics_from_excel(file_path):
    # Load categories and metrics from the first sheet
    df = pd.read_excel(file_path)
    metrics_list = {}
    
    # Populate categories and metrics
    for index, row in df.iterrows():
        category = row[0]
        metrics = row[1:].dropna().tolist()  # Extract all non-null values as metrics
        metrics_list[category] = metrics
    
    # Load the metric details (description, criteria) from the "metrics" sheet
    df = pd.read_excel(file_path, sheet_name="metrics")
    metrics_data = {}
    
    # Populate the metric data with description and criteria
    for index, row in df.iterrows():
        metric_name = row['Metric']
        description = row['Definition']  
        criteria = []

        for col in df.columns[3:]:  # Criteria from column starting from D onward
            value = row[col]
            if pd.isna(value):
                break
            criteria.append(value)
        
        # Store both description and criteria
        metrics_data[metric_name] = {
            'description': description,
            'criteria': criteria
        }

    return metrics_list, metrics_data

# Load the metrics and KPI data
metrics_list, metrics_data = load_metrics_from_excel('Master health check icsa Feb 24.xlsx')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        selected_metrics = request.form.getlist('selected_metrics')  # Retrieve selected metrics
        
        # Handle the case where the list is a string inside another list
        if selected_metrics and isinstance(selected_metrics[0], str):
            try:
                selected_metrics = json.loads(selected_metrics[0])
            except json.JSONDecodeError:
                print("Error parsing the metrics list string.")
        
        return render_template('results.html', selected_metrics=selected_metrics, metrics_data=metrics_data)

    return render_template('index.html', metrics_list=metrics_list, metrics_data=metrics_data)

@app.route('/save_results', methods=['POST'])
def save_results():
    criteria_values = {}
    for key, value in request.form.items():
        if key.startswith('criteria_'):
            criteria_values[key] = float(value)  # Convert criteria values to float
    
    # Perform calculations based on the criteria values
    outcomes = {}
    for metric, values in criteria_values.items():
        outcomes[metric] = sum(values)
    
    return render_template('outcome.html', outcomes=outcomes)

if __name__ == '__main__':
    app.run(host='0.0.0.0')
