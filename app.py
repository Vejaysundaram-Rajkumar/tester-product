from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

# Load metrics from Excel
def load_metrics_from_excel(file_path):
    df = pd.read_excel(file_path)
    categories_and_metrics = {}


    for index, row in df.iterrows():
        category = row[0]
        metrics = row[1:].dropna().tolist() 
        categories_and_metrics[category] = metrics

    return categories_and_metrics

# Load the metrics from the uploaded Excel file
categories_and_metrics = load_metrics_from_excel('Master health check icsa Feb 24.xlsx')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        selected_metrics = request.form.get('selected_metrics')
        selected_metrics = eval(selected_metrics)  
        return render_template('results.html', selected_metrics=selected_metrics)

    return render_template('index.html', categories_and_metrics=categories_and_metrics)

if __name__ == '__main__':
    app.run(host='0.0.0.0')
