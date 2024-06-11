import os
import pandas as pd

def count_records_in_files(directory_path):
    """
    Count the total number of records in text and Excel files within a specified directory.

    Args:
    directory_path (str): Path to the directory containing the files.

    Returns:
    int: Total number of records across all files.
    """
    total_records = 0
    # List all files in the given directory
    for filename in os.listdir(directory_path):
        # Construct full file path
        file_path = os.path.join(directory_path, filename)
        # Check if it's a file
        if os.path.isfile(file_path):
            try:
                if file_path.endswith('.txt'):
                    with open(file_path, 'r', encoding='utf-8') as file:
                        total_records += sum(1 for line in file)
                elif file_path.endswith(('.xlsx', '.xls', '.xlsm')):
                    df = pd.read_excel(file_path)
                    total_records += len(df)
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
    return total_records

# Example usage
directory = 'C:\\Users\\lmonroy\\Tema\\archivos_red\\Origen\\Origen'
print(count_records_in_files(directory))
