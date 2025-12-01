"""
Data balancing module for equalizing dataframe sizes.
Handles random deletion of rows to match dataframe sizes.
"""


def balance_dataframes(fan_low_df, fan_high_df):
    """
    Balance two dataframes by randomly deleting rows so they have the same size.
    
    The target size is the smaller of the two dataframe lengths, rounded to the 
    nearest 100. Both dataframes are reduced to this size.
    
    Args:
        fan_low_df (pd.DataFrame): DataFrame with low speed fans
        fan_high_df (pd.DataFrame): DataFrame with high speed fans
        
    Returns:
        tuple: (balanced_low_df, balanced_high_df) - Two balanced dataframes with equal rows
    """
    print("\n" + "="*80)
    print("DATAFRAME BALANCING")
    print("="*80)
    print(f"\nInitial row counts:")
    print(f"  Low speed fans:   {len(fan_low_df):,} rows")
    print(f"  High speed fans:  {len(fan_high_df):,} rows")
    
    # Use the smaller dataframe size, rounded to nearest 100
    min_rows = min(len(fan_low_df), len(fan_high_df))
    target_rows = round(min_rows / 100) * 100
    
    print(f"\nMinimum rows: {min_rows:,}")
    print(f"Target rows (rounded to nearest 100): {target_rows:,}\n")
    
    # Create copies to avoid modifying original dataframes
    balanced_low_df = fan_low_df.copy()
    balanced_high_df = fan_high_df.copy()
    
    # Randomly sample rows from each dataframe to reach target size
    # Ensure we don't try to sample more rows than available
    actual_target = min(target_rows, len(balanced_low_df), len(balanced_high_df))
    balanced_low_df = balanced_low_df.sample(n=actual_target, random_state=42)
    balanced_high_df = balanced_high_df.sample(n=actual_target, random_state=42)
    
    # Display results
    print(f"Final row counts after balancing:")
    print(f"  Low speed fans:   {len(balanced_low_df):,} rows")
    print(f"  High speed fans:  {len(balanced_high_df):,} rows")
    print(f"  Total:            {len(balanced_low_df) + len(balanced_high_df):,} rows")
    print("="*80 + "\n")
    
    return balanced_low_df, balanced_high_df
