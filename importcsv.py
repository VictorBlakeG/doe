import pandas as pd


def load_csv_data():
    """
    Open the CSV file and import data into a pandas dataframe.
    
    Returns:
        pd.DataFrame: The raw data from the CSV file
    """
    filepath = "data"
    filename = "iad12_clean.csv"
    fullname = f"{filepath}/{filename}"
    csv_raw_df = pd.read_csv(fullname)
    return csv_raw_df


def show_data_summary(csv_raw_df):
    """
    Display the head and tail of the dataframe with all column names
    and provide a summary of total lines imported.
    
    Args:
        csv_raw_df (pd.DataFrame): The dataframe to summarize
    """
    print("\n" + "="*80)
    print("DATA SUMMARY")
    print("="*80)
    
    print(f"\nTotal rows imported: {len(csv_raw_df)}")
    print(f"Total columns: {len(csv_raw_df.columns)}")
    
    print("\nColumn names:")
    for i, col in enumerate(csv_raw_df.columns, 1):
        print(f"  {i}. {col}")
    
    print("\n" + "-"*80)
    print("HEAD (First 5 rows):")
    print("-"*80)
    print(csv_raw_df.head())
    
    print("\n" + "-"*80)
    print("TAIL (Last 5 rows):")
    print("-"*80)
    print(csv_raw_df.tail())


def remove_missing_data(csv_raw_df):
    """
    Remove all rows that have missing data (NaN values).
    Display the number of rows before and after removal.
    
    Args:
        csv_raw_df (pd.DataFrame): The dataframe to clean
        
    Returns:
        pd.DataFrame: The cleaned dataframe without missing values
    """
    rows_before = len(csv_raw_df)
    
    # Remove rows with any missing data
    csv_clean_df = csv_raw_df.dropna()
    
    rows_after = len(csv_clean_df)
    rows_removed = rows_before - rows_after
    
    print("\n" + "="*80)
    print("DATA CLEANING SUMMARY")
    print("="*80)
    print(f"Rows before cleaning: {rows_before}")
    print(f"Rows after cleaning:  {rows_after}")
    print(f"Rows removed:         {rows_removed}")
    print("="*80 + "\n")
    
    return csv_clean_df
