"""
Main module that orchestrates data import and cleaning workflow.
"""
from importcsv import load_csv_data, show_data_summary, remove_missing_data
from prep import calculate_fan_speed_mean, create_fan_speed_histogram
from split import split_fan
from clean import analyze_device_vendors, clean_tman
from export import export_fan_dfs_to_csv
from balance import balance_dataframes
from viz import create_fan_hl_histogram, create_ttemp_hl_histogram
from doe import setup_doe_design, create_full_factorial_design, fit_doe_model


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

    print(f"Starting dataframe has {len(csv_clean_df)} rows\n")

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
    
    # Step 10: Balance dataframes to equal size
    print("Step 10: Balancing dataframes to equal size...")
    balanced_low_df, balanced_high_df = balance_dataframes(fan_low_df, fan_high_df)
    print("✓ Dataframes balanced successfully\n")
    
    # Step 11: Generate comparative histograms
    print("Step 11: Generating comparative fan speed histograms...")
    hl_html_file = create_fan_hl_histogram(balanced_low_df, balanced_high_df)
    print(f"✓ Comparative histogram generated: {hl_html_file}\n")
    
    # Step 12: Generate interface temperature comparative histograms
    print("Step 12: Generating interface temperature comparative histograms...")
    ttemp_html_file = create_ttemp_hl_histogram(balanced_low_df, balanced_high_df)
    print(f"✓ Temperature histogram generated: {ttemp_html_file}\n")
    
    # Step 13: Set up Design of Experiments
    print("Step 13: Setting up Design of Experiments...")
    doe_df = setup_doe_design(balanced_low_df, balanced_high_df)
    print("✓ DOE design setup completed\n")
    
    # Step 14: Create full-factorial design
    print("Step 14: Creating full-factorial design...")
    design_table = create_full_factorial_design(doe_df)
    print("✓ Full-factorial design created\n")
    
    # Step 15: Fit DOE model
    print("Step 15: Fitting Design of Experiments model...")
    model, results, summary_stats = fit_doe_model(doe_df)
    print("✓ DOE model fit completed\n")
    
    print("Data processing pipeline completed!")



if __name__ == "__main__":
    main()
