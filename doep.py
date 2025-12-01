"""
Main module that orchestrates data import and cleaning workflow.
"""
from import_module import load_csv_data, show_data_summary, remove_missing_data
from prep import calculate_fan_speed_mean, create_fan_speed_histogram


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
    
    # Step 5: Generate fan speed mean histogram
    print("Step 5: Generating fan speed mean histogram...")
    html_file = create_fan_speed_histogram(mean_fan_df)
    print(f"✓ Histogram generated: {html_file}\n")
    
    print("Data processing pipeline completed!")
    print(f"Final dataframe has {len(mean_fan_df)} rows\n")
    
    return mean_fan_df


if __name__ == "__main__":
    main()
