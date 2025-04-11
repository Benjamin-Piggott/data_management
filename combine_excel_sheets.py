import pandas as pd
import os

def load_excel_files(directory: str, prefix: str, start: int, end: int) -> pd.DataFrame:
    """
    Load Excel files from a specified directory and combine them into a single DataFrame.
    
    This function iterates over a range of Excel files whose names are composed of a given prefix 
    followed by a zero-padded numerical index and the '.xlsx' extension. For each file, the data 
    is read into a DataFrame using openpyxl as the engine. After reading each file, the function 
    checks if the first row of data (after the header row) contains units (e.g. "deg" in the 'CA' column)
    and removes it if present. A new column is added to record the source file name. Finally, all 
    DataFrames are concatenated into one master DataFrame.
    
    Parameters:
        directory (str): The path to the directory containing the Excel files.
        prefix (str): The common prefix of the Excel files (e.g. "Traces-").
        start (int): The starting index for the files.
        end (int): The ending index for the files (inclusive).
        
    Returns:
        pd.DataFrame: A single DataFrame containing the data from all Excel files.
    """
    dataframes = []
    
    for i in range(start, end + 1):
        file_name = f"{prefix}{i:03d}.xlsx"
        file_path = os.path.join(directory, file_name)
        
        try:
            # Read the Excel file using the openpyxl engine
            df = pd.read_excel(file_path, engine='openpyxl')
        except Exception as e:
            print(f"Error reading {file_name}: {e}")
            continue
        
        # Check if the first row of data (row 0 in the DataFrame) contains unit information.
        # For example, if the value in the 'CA' column is 'deg', assume this row contains units.
        if not df.empty and isinstance(df.iloc[0, 0], str) and df.iloc[0, 0].strip().lower() == "deg":
            df = df.iloc[1:].reset_index(drop=True)
        
        # Add a column for the source file name (for traceability)
        df['source_file'] = file_name
        
        dataframes.append(df)
    
    if dataframes:
        combined_df = pd.concat(dataframes, ignore_index=True)
    else:
        combined_df = pd.DataFrame()
    
    return combined_df

def main():
    """
    Main function to load and combine Excel files from a specified directory.
    
    This function first checks whether the 'openpyxl' dependency is installed. It then defines the 
    directory path, file naming scheme and index range, and calls the load_excel_files() function. 
    Finally, it outputs basic information and a preview of the combined DataFrame.
    """
    # Check if openpyxl is installed
    try:
        import openpyxl  # noqa: F401
    except ModuleNotFoundError:
        print("openpyxl is not installed. Please install it using 'pip install openpyxl' or "
              "'conda install openpyxl' and then re-run the script.")
        return

    # Define the directory path (using a raw string to handle backslashes correctly)
    folder_path = r"C:\Users\171218\Desktop\Uni\Masters\XE703 - Professional Development\Dataset\Traces"
    prefix = "Traces-"
    start_index = 1
    end_index = 20

    # Load the Excel files and combine them into one DataFrame
    combined_df = load_excel_files(folder_path, prefix, start_index, end_index)
    
    # Output an initial inspection of the combined data
    print("DataFrame Information:")
    print(combined_df.info())
    print("\nPreview of the data:")
    print(combined_df.head())

if __name__ == "__main__":
    main()
