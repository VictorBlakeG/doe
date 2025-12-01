"""
Main module that orchestrates data import and cleaning workflow.
"""
from importcsv import load_csv_data, show_data_summary, remove_missing_data
from prep import calculate_fan_speed_mean, create_fan_speed_histogram
from split import split_fan
from clean import analyze_device_vendors, clean_tman
from export import export_fan_dfs_to_csv


def main():
    """
    Main function to execute the data import and cleaning pipeline.
    """
    print("\nStarting data processing pipeline...\n")
    
    # Step 1: Load CSV data
    print("Step 1: Loading CSV data...")
    csv_raw_df = load_csv_data()
    print("✓ CSV data loaded successfully\n")
    
    # Step 2: Show data summary
    print("Step 2: Displaying data summary...")
    show_data_summary(csv_raw_df)
    
    # Step 3: Remove missing data
    print("Step 3: Removing rows with missing data...")
    csv_clean_df = remove_missing_data(csv_raw_df)
    
    # Step 4: Calculate fan speed mean
    print("Step 4: Calculating fan speed mean...")
    mean_fan_df = calculate_fan_speed_mean(csv_clean_df)
    print("✓ Fan speed mean calculated and added to dataframe\n")
    
    # Step 5: Platform Manufacturer clean
    print("Step 5: Platform Manufacturer clean...")
    vendor_counts = analyze_device_vendors(mean_fan_df)
    print("✓ Device vendor analysis completed\n")
    
    # Step 6: SFP Manufacturer clean
    print("Step 6: Transceiver Manufacturer clean...")
    sfp_counts = clean_tman(mean_fan_df)
    print("✓ Transceiver manufacturer analysis completed\n")
    
    # Step 7: Generate fan speed mean histogram
    print("Step 7: Generating fan speed mean histogram...")
    html_file = create_fan_speed_histogram(mean_fan_df)
    print(f"✓ Histogram generated: {html_file}\n")
    
    # Step 8: Split fan data by speed threshold
    print("Step 8: Splitting fan data by speed threshold...")
    fan_low_df, fan_high_df = split_fan(mean_fan_df)
    print(f"✓ Fan data split successfully\n")
    
    # Display counts
    print("================================================================================")
    print("FAN SPEED DISTRIBUTION")
    print("================================================================================")
    print(f"Low speed fans (< 11999):    {len(fan_low_df):,} rows")
    print(f"High speed fans (>= 12000):  {len(fan_high_df):,} rows")
    print(f"Total:                       {len(fan_low_df) + len(fan_high_df):,} rows")
    print("================================================================================\n")
    
    # Step 9: Export fan dataframes to CSV
    print("Step 9: Exporting fan dataframes to CSV...")
    low_csv_path, high_csv_path = export_fan_dfs_to_csv(fan_low_df, fan_high_df)
    print(f"✓ Low speed fan data exported: {low_csv_path}")
    print(f"✓ High speed fan data exported: {high_csv_path}\n")
    
    print("Data processing pipeline completed!")
    print(f"Final dataframe has {len(mean_fan_df)} rows\n")
    
    return mean_fan_df


if __name__ == "__main__":
    main()
