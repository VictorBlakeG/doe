"""
Data preparation module for processing cleaned dataframes.
Handles feature engineering and data transformations.
"""
import plotly.graph_objects as go
from pathlib import Path


def calculate_fan_speed_mean(csv_clean_df):
    """
    Process fan sensor data and calculate mean fan speed.
    
    Copies the input dataframe and calculates the mean across four fan trays,
    storing the result in a new 'fan_speed_mean' column.
    
    Args:
        csv_clean_df (pd.DataFrame): Cleaned dataframe with fan sensor columns
        
    Returns:
        pd.DataFrame: DataFrame with added 'fan_speed_mean' column
    """
    # Copy the dataframe to avoid modifying the original
    mean_fan_df = csv_clean_df.copy()
    
    # Define the fan sensor columns (from different fan trays)
    fan_sensor_columns = [
        'Fan_FanTray1Fan1Sensor1',
        'Fan_FanTray2Fan1Sensor1',
        'Fan_FanTray3Fan1Sensor1',
        'Fan_FanTray4Fan1Sensor1'
    ]
    
    # Calculate the mean across the four fan trays and add to new column
    mean_fan_df['fan_speed_mean'] = mean_fan_df[fan_sensor_columns].mean(axis=1)
    
    return mean_fan_df


def create_fan_speed_histogram(mean_fan_df, output_dir='outputs'):
    """
    Create an interactive HTML histogram of the fan_speed_mean distribution.
    
    Args:
        mean_fan_df (pd.DataFrame): DataFrame with 'fan_speed_mean' column
        output_dir (str): Directory where the HTML file will be saved
        
    Returns:
        str: Path to the generated HTML file
    """
    # Create output directory if it doesn't exist
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    # Extract fan speed mean data
    fan_speed_data = mean_fan_df['fan_speed_mean'].dropna()
    
    # Create histogram
    fig = go.Figure(data=[
        go.Histogram(
            x=fan_speed_data,
            nbinsx=50,
            name='Fan Speed Mean',
            marker=dict(color='#1f77b4', line=dict(color='#0d47a1', width=1.5))
        )
    ])
    
    # Update layout
    fig.update_layout(
        title='Distribution of Fan Speed Mean',
        xaxis_title='Fan Speed Mean (RPM)',
        yaxis_title='Frequency',
        hovermode='x unified',
        template='plotly_white',
        showlegend=False,
        height=600,
        margin=dict(l=100, r=50, t=100, b=50)
    )
    
    # Add statistics to the plot
    mean_val = fan_speed_data.mean()
    median_val = fan_speed_data.median()
    std_val = fan_speed_data.std()
    min_val = fan_speed_data.min()
    max_val = fan_speed_data.max()
    
    stats_text = (
        f'<b>Statistics</b><br>'
        f'Mean: {mean_val:.2f}<br>'
        f'Median: {median_val:.2f}<br>'
        f'Std Dev: {std_val:.2f}<br>'
        f'Min: {min_val:.2f}<br>'
        f'Max: {max_val:.2f}'
    )
    
    fig.add_annotation(
        text=stats_text,
        xref='paper', yref='paper',
        x=0.98, y=0.97,
        showarrow=False,
        bgcolor='rgba(255, 255, 255, 0.8)',
        bordercolor='#1f77b4',
        borderwidth=2,
        align='left',
        xanchor='right',
        yanchor='top',
        font=dict(size=12)
    )
    
    # Save HTML file
    html_file = output_path / 'fan_speed_mean.html'
    fig.write_html(str(html_file))
    
    return str(html_file)
