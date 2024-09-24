from flask import Flask, render_template, request ,flash ,redirect, url_for
import pandas as pd
import json
import openpyxl
import re
import xlwings as xw
import os

import psutil



app = Flask(__name__)


def force_close_excel():
    # Iterate over all running processes and kill any Excel instances
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == 'EXCEL.EXE':  # Target Excel processes
            try:
                proc.kill()  # Kill the process
                print(f"Closed Excel process with PID: {proc.info['pid']}")
            except Exception as e:
                print(f"Could not close Excel process: {e}")



# Calculation with the user input and excel updatation
def calculate_outcome(user_input):
    # Step 1: Parse the user input to extract metrics, criteria, and targets
    metrics_data = {}
    targets_data = {}

    for key, value in user_input.items():
        if key.startswith('criteria_'):
            # Extract the metric and criteria from the key
            pattern = r'criteria_(.*?)_(.*)'  # Regex to match the metric and criteria
            match = re.match(pattern, key)
            
            if match:
                metric_name = match.group(1)  # Extracted metric name
                criteria_name = match.group(2)  # Extracted criteria name

                # If the metric is not yet in the dictionary, initialize it with empty lists
                if metric_name not in metrics_data:
                    metrics_data[metric_name] = [[], []]  # First list for criteria names, second for values

                # Add the criteria and its value to the lists
                metrics_data[metric_name][0].append(criteria_name)
                metrics_data[metric_name][1].append(value)
        
        elif key.startswith('target_'):
            # Extract the metric name for the target
            metric_name = key.replace('target_', '')
            targets_data[metric_name] = value

    # Step 2: Load the Excel workbook and update the dataset with the extracted criteria values
    file_path = 'Master health check icsa Feb 24.xlsx'
    workbook = openpyxl.load_workbook(file_path, data_only=False)  # Load to keep formulas intact
    sheet = workbook['metrics']  # Adjust to the correct sheet name if necessary

    # Dictionary to store outcome by metric row for fetching later
    metric_row_mapping = {}

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        metric_name = row[0].value  # Assuming metric name is in column A
        if not metric_name:
            continue  # Skip if there's no metric name

        # Check if the metric exists in the parsed user input
        if metric_name in metrics_data:
            criteria_list, value_list = metrics_data[metric_name]

            # Store the row number for outcome reference later
            metric_row_mapping[metric_name] = row[0].row  # Store the row number of this metric

            # Update the Excel sheet with user values corresponding to the criteria
            for idx, criteria in enumerate(criteria_list):
                for col_idx in range(3, 8):  # Columns D to H in the Excel sheet contain the criteria names
                    if row[col_idx].value == criteria:
                        sheet.cell(row=row[0].row, column=col_idx + 1).value = value_list[idx]
                        break

    # Save the updated workbook temporarily to recalculate outcomes
    updated_file_path = 'Updated_Metrics_Dataset.xlsx'
    workbook.save(updated_file_path)

    # Step 3: Use xlwings to reopen the updated Excel file and fetch the computed outcomes
    app = xw.App(visible=False)
    workbook_xlwings = xw.Book(updated_file_path)
    sheet_xlwings = workbook_xlwings.sheets['metrics']  # Adjust to the correct sheet name if necessary

    # Step 4: Create a new workbook for the structured output
    new_workbook = openpyxl.Workbook()
    new_sheet = new_workbook.active
    new_sheet.title = "Selected Metrics"

    # Step 5: Add headers for the structured layout
    headers = ["Metric", "Criteria", "Value", "Outcome", "Target"]
    new_sheet.append(headers)

    # Step 6: Loop over the metrics and write them into the new sheet in the desired structure
    for metric_name, (criteria_list, value_list) in metrics_data.items():
        # Extract the outcome using xlwings
        metric_row_in_sheet = metric_row_mapping[metric_name]
        outcome_cell = sheet_xlwings.range(f'I{metric_row_in_sheet}').value  # Assuming outcome is in column I
        # Determine the target for this metric (if provided, else 'NA')
        target_value = targets_data.get(metric_name, 'NA')

        # Write the metric name, criteria, and values row by row
        for idx, criteria_name in enumerate(criteria_list):
            # If this is the first row for this metric, we write the outcome and target; otherwise, leave them blank
            if idx == 0:
                new_sheet.append([metric_name,  # Write the metric name once
                                criteria_name,  # Criteria name
                                value_list[idx],  # User input value
                                outcome_cell,  # Outcome
                                target_value])  # Target value
            else:
                new_sheet.append(["",  # Blank for metric name after first row
                                criteria_name,
                                value_list[idx],
                                "",  # Outcome blank for subsequent criteria rows
                                ""])  # Leave target blank for subsequent criteria rows

    # Step 7: Close the xlwings workbook and save the new workbook with the structured output
    workbook_xlwings.close()
    app.quit()

    # Step 8: Save the new workbook
    new_file_path = 'Structured_Metrics_Output.xlsx'
    new_workbook.save(new_file_path)

    # Step 9: Delete the updated temporary Excel file
    if os.path.exists(updated_file_path):
        os.remove(updated_file_path)

    print(f"New structured Excel file saved as {new_file_path}  and temporary file deleted.")


# Load the list of metrics from the master sheet and the details from metrics sheet.
def load_metrics_from_excel(file_path):
    # Load categories and metrics from the first sheet
    df = pd.read_excel(file_path)
    metrics_list = {}
    
    # Populate categories and metrics
    for index, row in df.iterrows():
        category = row[0]
        metrics = row[1:].dropna().tolist()  # Extract all  metrics
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
    converted_input = {}
    user_input = request.form.to_dict()
    for key, value in user_input.items():
        try:
            if '.' in value:
                converted_input[key] = float(value)
            else:
                converted_input[key] = int(value)
        except (ValueError, TypeError):
            converted_input[key] = 'NA'
    
    # Generate the structured metrics output
    calculate_outcome(converted_input)

    # Read the generated Excel file
    output_file = 'Structured_Metrics_Output.xlsx'
    df = pd.read_excel(output_file)

    # Replace NaN values with an empty string, but keep 'NA' values intact
    df = df.where(pd.notnull(df), '')

    # Convert DataFrame to a list of dictionaries for rendering
    results_data = df.to_dict(orient='records')

    return render_template('analysis.html', results_data=results_data)

@app.route('/close_excel', methods=['POST'])
def close_excel():
    try:
        force_close_excel()
        flash('Excel process has been successfully closed.')
    except Exception as e:
        flash(f"Error occurred while closing Excel: {e}")
    
    # Redirect back to the results or home page after execution
    return redirect(url_for('index'))




if __name__ == '__main__':
    app.run(host='0.0.0.0')
