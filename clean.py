"""
Data cleaning and analysis module for device platform information.
Handles analysis of device vendors and manufacturers.
"""


def analyze_device_vendors(mean_fan_df):
    """
    Analyze and display a summary of device vendor/manufacturer distribution.
    
    Counts the number of instances (rows) for each unique device vendor
    and displays the results in a formatted table.
    
    Args:
        mean_fan_df (pd.DataFrame): DataFrame with 'Device Vendor' column
        
    Returns:
        pd.Series: Value counts of device vendors
    """
    # Get value counts for each device vendor
    vendor_counts = mean_fan_df['Device Vendor'].value_counts()
    
    # Display the analysis
    print("\n" + "="*80)
    print("PLATFORM MANUFACTURER DISTRIBUTION")
    print("="*80)
    print(f"\nTotal unique manufacturers: {len(vendor_counts)}\n")
    
    # Display each vendor with its count
    for vendor, count in vendor_counts.items():
        percentage = (count / len(mean_fan_df)) * 100
        print(f"{vendor:.<40} {count:>6} rows ({percentage:>5.1f}%)")
    
    print("\n" + "="*80 + "\n")
    
    return vendor_counts


def clean_tman(mean_fan_df):
    """
    Analyze and display a summary of SFP manufacturer distribution.
    
    Counts the number of instances (rows) for each unique SFP manufacturer
    and displays the results in a formatted table. Consolidates FINISAR variants
    into a single "Finisar" entry.
    
    Args:
        mean_fan_df (pd.DataFrame): DataFrame with 'SFP_manufacturer' column
        
    Returns:
        pd.Series: Value counts of SFP manufacturers (cleaned)
    """
    # Create a copy to avoid modifying the original dataframe
    df_copy = mean_fan_df.copy()
    
    # Consolidate FINISAR variants into "Finisar"
    df_copy['SFP_manufacturer'] = df_copy['SFP_manufacturer'].replace({
        'FINISAR CORP': 'Finisar',
        'FINISAR CORP.': 'Finisar',
        'FINISAR': 'Finisar'
    })
    
    # Get value counts for each SFP manufacturer
    sfp_counts = df_copy['SFP_manufacturer'].value_counts()
    
    # Display the analysis
    print("\n" + "="*80)
    print("TRANSCEIVER MANUFACTURER DISTRIBUTION")
    print("="*80)
    print(f"\nTotal unique transceiver manufacturers: {len(sfp_counts)}\n")
    
    # Display each manufacturer with its count (no dots separator)
    for manufacturer, count in sfp_counts.items():
        percentage = (count / len(df_copy)) * 100
        print(f"{manufacturer:<40} {count:>6} rows ({percentage:>5.1f}%)")
    
    print("\n" + "="*80 + "\n")
    
    return sfp_counts
