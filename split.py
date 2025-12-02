"""
Data splitting module for fan speed analysis.
Handles partitioning dataframes based on fan speed thresholds.
"""


def split_fan(mean_fan_df):
    """
    Split dataframe into two groups based on fan speed mean threshold.
    
    Rows with fan_speed_mean < 999 go to fan_low_df
    Rows with fan_speed_mean >= 10000 go to fan_high_df
    
    Args:
        mean_fan_df (pd.DataFrame): DataFrame with 'fan_speed_mean' column
        
    Returns:
        tuple: (fan_low_df, fan_high_df) - Two dataframes split by threshold
    """
    # Split dataframe based on fan speed threshold
    fan_low_df = mean_fan_df[mean_fan_df['fan_speed_mean'] < 9999].copy()
    fan_high_df = mean_fan_df[mean_fan_df['fan_speed_mean'] >= 10000].copy()
    
    return fan_low_df, fan_high_df
