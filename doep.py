"""
Main module that orchestrates data import and cleaning workflow.
"""
from importcsv import load_csv_data, show_data_summary, remove_missing_data
from prep import calculate_fan_speed_mean, create_fan_speed_histogram
from split import split_fan
from clean import clean_tman
from export import export_fan_dfs_to_csv
from balance import balance_dataframes
from viz import create_fan_hl_histogram, create_ttemp_hl_histogram
from doe import setup_doe_design, create_full_factorial_design, fit_doe_model, fit_reduced_doe_model, convert_html_to_pdf
from powerpoint_generator import convert_html_to_powerpoint


def main():
    """
    Main function to execute the data import and cleaning pipeline.
    """
    print("\nStarting data processing pipeline...\n")
    
    # Step 1: Load fan speed dataframes from CSV
    print("Step 1: Loading fan_low_df.csv and fan_high_df.csv...")
    import pandas as pd
    fan_low_df = pd.read_csv('outputs/fan_low_df.csv')
    fan_high_df = pd.read_csv('outputs/fan_high_df.csv')
    print(f"✓ fan_low_df loaded: {len(fan_low_df):,} rows")
    print(f"✓ fan_high_df loaded: {len(fan_high_df):,} rows\n")
    
    # Step 2: Show data summary for both dataframes
    print("Step 2: Displaying data summary...")
    print("\n--- Low Speed Fans ---")
    show_data_summary(fan_low_df)
    print("\n--- High Speed Fans ---")
    show_data_summary(fan_high_df)
    print()
    
    # Step 3: Remove missing data for both dataframes
    print("Step 3: Removing rows with missing data...")
    fan_low_clean_df = remove_missing_data(fan_low_df)
    fan_high_clean_df = remove_missing_data(fan_high_df)
    print(f"Low speed dataframe after cleaning: {len(fan_low_clean_df):,} rows")
    print(f"High speed dataframe after cleaning: {len(fan_high_clean_df):,} rows\n")

    # Step 4: Check if fan_speed_mean column exists
    if 'fan_speed_mean' in fan_low_clean_df.columns and 'fan_speed_mean' in fan_high_clean_df.columns:
        print("Step 4: fan_speed_mean column already exists in both dataframes - skipping\n")
        fan_low_df_final = fan_low_clean_df
        fan_high_df_final = fan_high_clean_df
    else:
        print("Step 4: Calculating fan speed mean for both dataframes...")
        fan_low_df_final = calculate_fan_speed_mean(fan_low_clean_df)
        fan_high_df_final = calculate_fan_speed_mean(fan_high_clean_df)
        print("✓ Fan speed mean calculated for both dataframes\n")
    
    # Step 5: Transceiver Manufacturer clean for both dataframes
    print("Step 5: Transceiver Manufacturer clean...")
    fan_low_clean_tman = clean_tman(fan_low_df_final)
    fan_high_clean_tman = clean_tman(fan_high_df_final)
    print("✓ Transceiver manufacturer analysis and consolidation completed for both dataframes\n")
    
    # Step 6: Generate fan speed mean histogram for both dataframes
    print("Step 6: Generating fan speed mean histograms...")
    html_file_low = create_fan_speed_histogram(fan_low_clean_tman)
    html_file_high = create_fan_speed_histogram(fan_high_clean_tman)
    print(f"✓ Low speed histogram generated: {html_file_low}")
    print(f"✓ High speed histogram generated: {html_file_high}\n")
    
    # Continue with the split data
    print("Step 7: Preparing fan data for subsequent analysis...")
    
    # Display counts
    print("================================================================================")
    print("FAN SPEED DISTRIBUTION")
    print("================================================================================")
    print(f"Low speed fans (< 11999):    {len(fan_low_clean_tman):,} rows")
    print(f"High speed fans (>= 12000):  {len(fan_high_clean_tman):,} rows")
    print(f"Total:                       {len(fan_low_clean_tman) + len(fan_high_clean_tman):,} rows")
    print("================================================================================\n")
    
    # Step 8: Balance dataframes to equal size
    print("Step 8: Balancing dataframes to equal size...")
    balanced_low_df, balanced_high_df = balance_dataframes(fan_low_clean_tman, fan_high_clean_tman)
    print("✓ Dataframes balanced successfully\n")
    
    # Step 9: Generate comparative histograms
    print("Step 9: Generating comparative fan speed histograms...")
    hl_html_file = create_fan_hl_histogram(balanced_low_df, balanced_high_df)
    print(f"✓ Comparative histogram generated: {hl_html_file}\n")
    
    # Step 10: Generate interface temperature comparative histograms
    print("Step 10: Generating interface temperature comparative histograms...")
    ttemp_html_file = create_ttemp_hl_histogram(balanced_low_df, balanced_high_df)
    print(f"✓ Temperature histogram generated: {ttemp_html_file}\n")
    
    # Step 11: Set up Design of Experiments
    print("Step 11: Setting up Design of Experiments...")
    doe_df = setup_doe_design(balanced_low_df, balanced_high_df)
    print("✓ DOE design setup completed\n")
    
    # Step 12: Create full-factorial design
    print("Step 12: Creating full-factorial design...")
    design_table = create_full_factorial_design(doe_df)
    print("✓ Full-factorial design created\n")
    
    # Step 13: Fit DOE model
    print("Step 13: Fitting Design of Experiments model...")
    model, results, summary_stats = fit_doe_model(doe_df)
    print("✓ DOE model fit completed\n")
    
    # Step 14: Fit reduced DOE model (removing non-significant terms)
    print("Step 14: Fitting Reduced Design of Experiments model...")
    reduced_model, reduced_results, reduced_summary_stats = fit_reduced_doe_model(doe_df, results)
    print("✓ Reduced DOE model fit completed\n")
    
    # Step 15: Convert HTML reports to PDF
    print("Step 15: Converting HTML reports to PDF...")
    pdf_status = convert_html_to_pdf()
    print("✓ HTML to PDF conversion completed\n")
    
    # Step 16: Convert HTML reports to PowerPoint
    print("Step 16: Converting HTML reports to PowerPoint...")
    pptx_status = convert_html_to_powerpoint()
    print("✓ HTML to PowerPoint conversion completed\n")
    
    print("Data processing pipeline completed!")



if __name__ == "__main__":
    main()
