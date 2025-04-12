import pandas as pd
import os
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

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
            # Check if the first row of data contains unit information (e.g. "deg" in 'CA')
            if not df.empty and 'CA' in df.columns:
                first_val = str(df.iloc[0]['CA']).strip().lower()
                if first_val == "deg":
                    print(f"    Removing units row from sheet: {sheet_name} in file {file_name}.")
                    df = df.iloc[1:].reset_index(drop=True)
            
            # Add metadata columns for traceability.
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

def perform_data_quality_checks(df: pd.DataFrame):
    """
    Perform data quality checks on the DataFrame.
    
    This function prints summary statistics, information about missing values, and the overall
    structure of the DataFrame to ensure data completeness and correct formatting.
    """
    print("Data Summary Statistics:")
    print(df.describe(include='all'))
    print("\nDataframe Info:")
    print(df.info())
    print("\nMissing Values Count:")
    print(df.isnull().sum())

def perform_visualisation_analysis(df: pd.DataFrame):
    """
    Perform basic visualisation and analysis on the data.
    
    This function creates plots to reveal the distribution of key variables and explore potential
    relationships between them. Adjust the column names as per your data structure.
    """
    # Example: Plot histogram of cylinder pressure (assumed to be 'PCYL1')
    if 'PCYL1' in df.columns:
        plt.figure(figsize=(8, 5))
        data_numeric = pd.to_numeric(df['PCYL1'], errors='coerce').dropna()
        plt.hist(data_numeric, bins=30, edgecolor='black')
        plt.title("Distribution of Cylinder Pressure (PCYL1)")
        plt.xlabel("Cylinder Pressure")
        plt.ylabel("Frequency")
        plt.show()
    
    # Example: Scatter plot for two variables (adjust column names as needed)
    if all(col in df.columns for col in ['Cylinde~', 'PCYL1']):
        plt.figure(figsize=(8, 5))
        x = pd.to_numeric(df['Cylinde~'], errors='coerce')
        y = pd.to_numeric(df['PCYL1'], errors='coerce')
        plt.scatter(x, y, alpha=0.7)
        plt.xlabel("Cylinde~")
        plt.ylabel("Cylinder Pressure (PCYL1)")
        plt.title("Relationship Between Cylinde~ and Cylinder Pressure")
        plt.show()
    
    # Example: Correlation Matrix Heatmap (using numeric columns only)
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
    if len(numeric_cols) > 1:
        plt.figure(figsize=(10, 8))
        corr = df[numeric_cols].corr()
        sns.heatmap(corr, annot=True, cmap='coolwarm')
        plt.title("Correlation Matrix")
        plt.show()

def perform_feature_engineering(df: pd.DataFrame) -> pd.DataFrame:
    """
    Create and transform features in the DataFrame.
    
    This function demonstrates simple feature engineering by creating new columns
    based on existing sensor data. Adjust the operations and column names as required.
    
    Returns:
        pd.DataFrame: The DataFrame with new engineered features.
    """
    # Example: Create a new feature that is the difference between 'Cylinde~' and 'RockerN~'
    if all(col in df.columns for col in ['Cylinde~', 'RockerN~']):
        df['sensor_diff'] = pd.to_numeric(df['Cylinde~'], errors='coerce') - pd.to_numeric(df['RockerN~'], errors='coerce')
    
    # Example: Log-transform cylinder pressure ('PCYL1') to reduce skew (ensure no negatives or zeros)
    if 'PCYL1' in df.columns:
        df['log_pressure'] = np.log(pd.to_numeric(df['PCYL1'], errors='coerce') + 1e-6)
    
    print("Feature engineering complete. New features added: 'sensor_diff', 'log_pressure'")
    return df

def main():
    """
    Main function to load, process, and analyse Excel files.
    
    This function checks the required dependencies, loads all Excel files and their sheets from a specified
    directory, performs data quality checks, visualises the data, and applies feature engineering. It then
    outputs a summary of the combined DataFrame.
    """
    try:
        import openpyxl  # noqa: F401
    except ModuleNotFoundError:
        print("openpyxl is not installed. Please install it using 'pip install openpyxl' or 'conda install openpyxl' and then re-run the script.")
        return

    # Define the directory path (handle backslashes correctly with a raw string)
    folder_path = r"C:\Users\171218\Desktop\Uni\Masters\XE703 - Professional Development\Dataset\Traces"
    prefix = "Traces-"
    start_index = 1
    end_index = 20

    # Load and combine data from Excel files
    combined_df = load_excel_files(folder_path, prefix, start_index, end_index)
    
    # Perform data quality checks
    perform_data_quality_checks(combined_df)
    
    # Visualise and analyse the data
    perform_visualisation_analysis(combined_df)
    
    # Apply feature engineering to create new variables
    combined_df = perform_feature_engineering(combined_df)
    
    # Output an overview of the final DataFrame
    print("\nFinal DataFrame Information:")
    print(combined_df.info())
    print("\nPreview of the final data:")
    print(combined_df.head())

if __name__ == "__main__":
    main()
