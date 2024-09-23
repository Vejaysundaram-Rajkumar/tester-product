import pandas as pd

def load_metrics_with_criteria(file_path, sheet_name="Cash KPIs"):
    # Read the specific sheet from the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Dictionary to store metrics and their criteria
    metrics_data = {}
    
    # Loop through each row
    for index, row in df.iterrows():
        metric_name = row['Metric']  # Assuming 'Metric' is the column name in column 'A'
        
        # Collect criteria from columns starting from 'D' (assumed criteria start at column D)
        criteria = []
        for col in df.columns[3:]:  # Columns D onward for Criteria 1, Criteria 2, etc.
            value = row[col]
            if pd.isna(value):  # Stop if 'NA' or empty value is found
                break
            criteria.append(value)
        
        # Add the metric and its criteria to the dictionary
        metrics_data[metric_name] = criteria
    
    return metrics_data

# Example usage
file_path = 'Master health check icsa Feb 24.xlsx'
metrics_data = load_metrics_with_criteria(file_path, sheet_name="Cash KPIs")
print(metrics_data)
