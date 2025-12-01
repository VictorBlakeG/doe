"""
Export module for saving dataframes to files.
Handles exporting processed data to various formats.
"""
from pathlib import Path


def export_fan_dfs_to_csv(fan_low_df, fan_high_df, output_dir='outputs'):
    """
    Export fan speed dataframes to CSV files.
    
    Saves fan_low_df and fan_high_df to CSV files in the specified output directory.
    
    Args:
        fan_low_df (pd.DataFrame): DataFrame with low speed fans (< 11999)
        fan_high_df (pd.DataFrame): DataFrame with high speed fans (>= 12000)
        output_dir (str): Directory where the CSV files will be saved
        
    Returns:
        tuple: (low_csv_path, high_csv_path) - Paths to the exported CSV files
    """
    # Create output directory if it doesn't exist
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    # Define output file paths
    low_csv_path = output_path / 'fan_low_df.csv'
    high_csv_path = output_path / 'fan_high_df.csv'
    
    # Export dataframes to CSV
    fan_low_df.to_csv(str(low_csv_path), index=False)
    fan_high_df.to_csv(str(high_csv_path), index=False)
    
    return str(low_csv_path), str(high_csv_path)
