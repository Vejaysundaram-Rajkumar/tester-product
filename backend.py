import pandas as pd

# Load the uploaded Excel file
file_path = 'test_data.xlsx'  # Replace with your file path

# Read the Excel file
excel_data = pd.read_excel(file_path)

# Function to compute the outcome based on the formula and user input
def compute_outcome_with_user_input(row):
    # Get the row index (Excel formulas are 1-based, so add 2 to match the correct row in this case)
    row_num = row.name + 2
    
    # Ask the user for Criteria 1 and Criteria 2 inputs
    criteria_1 = float(input(f"Enter the value for Criteria 1 (Metric {row['Metric']}): "))
    criteria_2 = float(input(f"Enter the value for Criteria 2 (Metric {row['Metric']}): "))
    
    # If Criteria 3 exists, ask the user for input; else, set it to None
    if not pd.isna(row['Criteria 3']):
        criteria_3 = float(input(f"Enter the value for Criteria 3 (Metric {row['Metric']}): "))
    else:
        criteria_3 = None
    
    # Update the row in the DataFrame with the user-provided values
    row['Criteria 1'] = criteria_1
    row['Criteria 2'] = criteria_2
    row['Criteria 3'] = criteria_3
    
    # Replace Excel-like references (D2, E2, etc.) with the user-provided values
    formula = row['Formula']
    formula = formula.replace(f'D{row_num}', str(criteria_1))
    formula = formula.replace(f'E{row_num}', str(criteria_2))
    if f'F{row_num}' in formula and criteria_3 is not None:
        formula = formula.replace(f'F{row_num}', str(criteria_3))
    
    try:
        # Evaluate the formula and return the result
        outcome = eval(formula)
    except Exception as e:
        outcome = None  # If there's any issue, set outcome to None
    
    return outcome

# Iterate over each row and compute the outcome based on user input
for idx, row in excel_data.iterrows():
    print(f"\nProcessing Metric {row['Metric']}:")
    outcome = compute_outcome_with_user_input(row)
    
    # Update the Outcome column for this row
    excel_data.at[idx, 'Outcome'] = outcome
    # Update Criteria columns with user input
    excel_data.at[idx, 'Criteria 1'] = row['Criteria 1']
    excel_data.at[idx, 'Criteria 2'] = row['Criteria 2']
    excel_data.at[idx, 'Criteria 3'] = row['Criteria 3']

# Save the updated DataFrame back to an Excel file
updated_file_path = 'updated_with_user_input_and_criteria.xlsx'  # Replace with your desired file path
excel_data.to_excel(updated_file_path, index=False)

print(f"\nExcel file updated and saved as {updated_file_path}")
