import pandas as pd

def sco():
    def fetch_calculated_values(file_paths, sheet_name='Additional Calculations'):
        combined_values = {}
        try:
            for file_path, keys in file_paths:
                data = pd.read_excel(file_path, sheet_name=sheet_name)
                for key in keys:
                    combined_values[key] = data[key].values[0] if key in data.columns else None
        except Exception as e:
            print(f"An error occurred while fetching the data from {file_path}: {e}")
            return None
        return combined_values

    def perform_calculations_and_save(values):
        sco_values = {
            'SCO1': values['A1CO1'],
            'SCO2': (values['A1CO2'] + values['A3CO3']) / 3 if values['A1CO2'] is not None and values['A3CO3'] is not None else None,
            'SCO3': (values['A1CO3'] + values['A2CO3'] + values['A3CO3']) / 3 if values['A1CO3'] is not None and values['A2CO3'] is not None and values['A3CO3'] is not None else None,
            'SCO4': (values['A2CO4'] + values['A3CO4']) / 2 if values['A2CO4'] is not None and values['A3CO4'] is not None else None,
            'SCO5': (values['A2CO5'] + values['A3CO5']) / 2 if values['A2CO5'] is not None and values['A3CO5'] is not None else None,
            'SCO6': (values['A2CO6'] + values['A3CO6']) / 2 if values['A2CO6'] is not None and values['A3CO6'] is not None else None
        }

        # Convert the dictionary to DataFrame for easier export to Excel
        results_df = pd.DataFrame([sco_values])

        # File path where the Excel will be saved
        output_excel_path = 'SCO_Results.xlsx'

        # Write the DataFrame to an Excel file
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            results_df.to_excel(writer, index=False, sheet_name='Calculated Scores')

        print(f"Calculated scores have been saved to {output_excel_path}.")
        return sco_values

    # File paths and the keys we expect to find in each

    file_paths = [
        ('mst-1 calculations.xlsx', ['A1CO1', 'A1CO2', 'A1CO3']),
        ('mst-2 calculations.xlsx', ['A2CO3', 'A2CO4', 'A2CO5', 'A2CO6']),
        ('mst-3 calculations.xlsx', ['A3CO2', 'A3CO3', 'A3CO4', 'A3CO5', 'A3CO6'])
    ]

    values = fetch_calculated_values(file_paths)
    if values:
        calculated_values = perform_calculations_and_save(values)
        print(calculated_values)
