import pandas as pd
import os

def convert_xlsx_to_csv(config):
    """
    Convert an Excel (XLSX) file to CSV format using a configuration variable
    
    Parameters:
        config (dict or str): Either a dictionary with input_path and optionally output_path,
                             or a string representing the input path
    
    Returns:
        str: Path to the created CSV file
    """
    # Handle the config parameter
    if isinstance(config, str):
        xlsx_file_path = config
        csv_file_path = None
    elif isinstance(config, dict):
        xlsx_file_path = config.get('input_path')
        csv_file_path = config.get('output_path')
    else:
        raise ValueError("Config must be either a string (file path) or a dictionary with 'input_path'")
    
    # Check if input file exists
    if not os.path.exists(xlsx_file_path):
        print(f"Error: Input file does not exist: {xlsx_file_path}")
        return None
    
    # Generate output path if not provided
    if csv_file_path is None:
        file_name = os.path.splitext(xlsx_file_path)[0]
        csv_file_path = f"{file_name}.csv"
    
    try:
        # Read the Excel file
        df = pd.read_excel(xlsx_file_path)
        
        # Write to CSV
        df.to_csv(csv_file_path, index=False)
        
        print(f"Successfully converted {xlsx_file_path} to {csv_file_path}")
        return csv_file_path
    
    except Exception as e:
        print(f"Error converting XLSX to CSV: {str(e)}")
        return None

# Example usage:
if __name__ == "__main__":
    # Example 1: Using a string config
    file_path = "cable_sizing_output.xlsx"
    convert_xlsx_to_csv(file_path)
    
    # Example 2: Using a dictionary config
    # config = {
    #     "input_path": "path/to/your/file.xlsx",
    #     "output_path": "path/to/output.csv"
    # }
    # convert_xlsx_to_csv(config)