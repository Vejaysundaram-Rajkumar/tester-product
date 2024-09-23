from flask import Flask, render_template, request
import pandas as pd
import json
app = Flask(__name__)

# Load the list of metrics from the master sheet and the details from individual KPI sheets
def load_metrics_from_excel(file_path):
    
    df = pd.read_excel(file_path)
    metrics_list = {}

    
    for index, row in df.iterrows():
        category = row[0]
        metrics = row[1:].dropna().tolist()  
        metrics_list[category] = metrics

    
    df = pd.read_excel(file_path, sheet_name="metrics")
    metrics_data = {}
        
    for index, row in df.iterrows():
        metric_name = row['Metric']
        criteria = []
        for col in df.columns[3:]:  # Collect criteria from columns starting from D onward
            value = row[col]
            if pd.isna(value):
                break
            criteria.append(value)
        
        metrics_data[metric_name] = criteria

    return metrics_list, metrics_data

# Load the metrics and KPI data
metrics_list, metrics_data = load_metrics_from_excel('Master health check icsa Feb 24.xlsx')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        selected_metrics = request.form.getlist('selected_metrics')  # Get the form input
        
        # Handle the case where the list is a string inside another list
        if selected_metrics and isinstance(selected_metrics[0], str):
            try:
                # Parse the string into a proper list
                selected_metrics = json.loads(selected_metrics[0])
            except json.JSONDecodeError:
                print("Error parsing the metrics list string.")
        
        print("Parsed Selected Metrics:", selected_metrics)  # Debugging line
        
        return render_template('results.html', selected_metrics=selected_metrics, metrics_data=metrics_data)

    return render_template('index.html', metrics_list=metrics_list)


@app.route('/save_results', methods=['POST'])
def save_results():
    criteria_values = {}
    for key, value in request.form.items():
        if key.startswith('criteria_'):
            criteria_values[key] = float(value)  # Convert criteria values to float
    
    # Here you would apply the formulas for each metric using the criteria_values
    outcomes = {}
    for metric, values in criteria_values.items():
        # Dummy logic: Sum criteria values (you can replace this with actual formula logic)
        outcomes[metric] = sum(values)
    
    return render_template('outcome.html', outcomes=outcomes)

if __name__ == '__main__':
    app.run(host='0.0.0.0')
