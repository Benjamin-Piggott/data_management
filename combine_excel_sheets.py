import pandas as pd
import os

def load_excel_files(directory: str, prefix: str, start: int, end: int) -> pd.DataFrame:
    """
    Load Excel files and all their sheets from a specified directory and combine them into a single DataFrame.
    
    This function iterates over a range of Excel files whose names are composed of a given prefix 
    followed by a zero-padded numerical index and the '.xlsx' extension. File Traces-015.xlsx is skipped 
    because it is known to be corrupted. For each valid file, the function reads all sheets at once into a
    dictionary using the 'sheet_name=None' parameter and then iterates over this dictionary. For each sheet, 
    if the first row appears to contain unit information (e.g. the value in the 'CA' column is "deg"), that row 
    is removed. Additional columns are added to record the source file and sheet name for traceability. 
    Finally, all the individual DataFrames are concatenated into one master DataFrame.
    
    Parameters:
        directory (str): The path to the directory containing the Excel files.
        prefix (str): The common prefix of the Excel files (e.g. "Traces-").
        start (int): The starting index for the files.
        end (int): The ending index for the files (inclusive).
        
    Returns:
        pd.DataFrame: A single DataFrame containing the data from all sheets across all valid Excel files.
    """
    dataframes = []
    
    for i in range(start, end + 1):
        # Skip the known corrupted file, e.g. Traces-015.xlsx.
        if i == 15:
            print(f"Skipping corrupted file: {prefix}{i:03d}.xlsx")
            continue
        
        file_name = f"{prefix}{i:03d}.xlsx"
        file_path = os.path.join(directory, file_name)
        print(f"Processing file: {file_name}...")
        
        try:
            # Open the Excel file using the openpyxl engine.
            xl = pd.ExcelFile(file_path, engine='openpyxl')
        except Exception as e:
            print(f"Error reading {file_name}: {e}")
            continue
        
        # Read all sheets at once into a dictionary.
        try:
            sheets_dict = xl.parse(sheet_name=None, engine='openpyxl')
        except Exception as e:
            print(f"Error reading sheets from {file_name}: {e}")
            continue
        
        print(f"  Found {len(sheets_dict)} sheet(s) in {file_name}.")
        
        # Iterate over the dictionary of DataFrames
        for sheet_name, df in sheets_dict.items():
            print(f"    Processing sheet: {sheet_name}...")
            # Check if the first row of data contains unit information.
            if not df.empty and 'CA' in df.columns:
                first_val = str(df.iloc[0]['CA']).strip().lower()
                if first_val == "deg":
                    print(f"    Removing units row from sheet: {sheet_name} in file {file_name}.")
                    df = df.iloc[1:].reset_index(drop=True)
            
            # Add columns for source file and sheet name for traceability.
            df['source_file'] = file_name
            df['sheet_name'] = sheet_name
            dataframes.append(df)
            
            print(f"    Finished processing sheet: {sheet_name}.")
    
    if dataframes:
        combined_df = pd.concat(dataframes, ignore_index=True)
        print("All files and sheets have been processed and combined.")
    else:
        combined_df = pd.DataFrame()
        print("No data was loaded. The resulting DataFrame is empty.")
    
    return combined_df

def main():
    """
    Main function to load and combine Excel files and all their sheets from a specified directory.
    
    This function first checks whether the 'openpyxl' dependency is installed. It then defines the 
    directory path, file naming scheme and numerical range, and calls the load_excel_files() function. 
    Finally, it outputs basic information and a preview of the combined DataFrame.
    """
    try:
        import openpyxl  # noqa: F401
    except ModuleNotFoundError:
        print("openpyxl is not installed. Please install it using 'pip install openpyxl' or "
              "'conda install openpyxl' and then re-run the script.")
        return

    folder_path = r"C:\Users\171218\Desktop\Uni\Masters\XE703 - Professional Development\Dataset\Traces"
    prefix = "Traces-"
    start_index = 1
    end_index = 20

    combined_df = load_excel_files(folder_path, prefix, start_index, end_index)
    
    print("DataFrame Information:")
    print(combined_df.info())
    print("\nPreview of the data:")
    print(combined_df.head())

if __name__ == "__main__":
    main()
