"""
Visualization module for comparative histogram analysis.
Handles creating side-by-side distribution comparisons with statistics.
"""
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path


def create_fan_hl_histogram(balanced_low_df, balanced_high_df, output_dir='outputs'):
    """
    Create side-by-side histograms comparing low and high speed fan distributions.
    
    Creates an interactive HTML visualization with separate histograms for low and
    high speed fans, including statistics (mean, std dev, variance, count).
    
    Args:
        balanced_low_df (pd.DataFrame): Balanced dataframe with low speed fans
        balanced_high_df (pd.DataFrame): Balanced dataframe with high speed fans
        output_dir (str): Directory where the HTML file will be saved
        
    Returns:
        str: Path to the generated HTML file
    """
    # Create output directory if it doesn't exist
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    # Extract fan speed data
    low_speed_data = balanced_low_df['fan_speed_mean'].dropna()
    high_speed_data = balanced_high_df['fan_speed_mean'].dropna()
    
    # Calculate statistics
    low_stats = {
        'mean': low_speed_data.mean(),
        'std': low_speed_data.std(),
        'var': low_speed_data.var(),
        'count': len(low_speed_data)
    }
    
    high_stats = {
        'mean': high_speed_data.mean(),
        'std': high_speed_data.std(),
        'var': high_speed_data.var(),
        'count': len(high_speed_data)
    }
    
    # Create subplots
    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=('Low (< 9999 rpm)', 'High (>= 10,000 rpm)'),
        specs=[[{'secondary_y': False}, {'secondary_y': False}]]
    )
    
    # Add low speed histogram
    fig.add_trace(
        go.Histogram(
            x=low_speed_data,
            nbinsx=40,
            name='Low Speed',
            marker=dict(color='#1f77b4', line=dict(color='#0d47a1', width=1.5)),
            showlegend=False
        ),
        row=1, col=1
    )
    
    # Add high speed histogram
    fig.add_trace(
        go.Histogram(
            x=high_speed_data,
            nbinsx=40,
            name='High Speed',
            marker=dict(color='#ff7f0e', line=dict(color='#d62728', width=1.5)),
            showlegend=False
        ),
        row=1, col=2
    )
    
    # Create statistics text for low speed
    low_stats_text = (
        f'<b>Low Speed Statistics</b><br>'
        f'Mean: {low_stats["mean"]:.2f}<br>'
        f'Std Dev: {low_stats["std"]:.2f}<br>'
        f'Variance: {low_stats["var"]:.2f}<br>'
        f'Count: {low_stats["count"]:,}'
    )
    
    # Create statistics text for high speed
    high_stats_text = (
        f'<b>High Speed Statistics</b><br>'
        f'Mean: {high_stats["mean"]:.2f}<br>'
        f'Std Dev: {high_stats["std"]:.2f}<br>'
        f'Variance: {high_stats["var"]:.2f}<br>'
        f'Count: {high_stats["count"]:,}'
    )
    
    # Add annotations for statistics
    fig.add_annotation(
        text=low_stats_text,
        xref='x domain', yref='y domain',
        x=0.98, y=0.97,
        showarrow=False,
        bgcolor='rgba(255, 255, 255, 0.8)',
        bordercolor='#1f77b4',
        borderwidth=2,
        align='left',
        xanchor='right',
        yanchor='top',
        font=dict(size=11),
        row=1, col=1
    )
    
    fig.add_annotation(
        text=high_stats_text,
        xref='x domain', yref='y domain',
        x=0.98, y=0.97,
        showarrow=False,
        bgcolor='rgba(255, 255, 255, 0.8)',
        bordercolor='#ff7f0e',
        borderwidth=2,
        align='left',
        xanchor='right',
        yanchor='top',
        font=dict(size=11),
        row=1, col=2
    )
    
    # Update layout
    fig.update_layout(
        title_text='Fan Speed Distribution Comparison: Low vs High Speed Fans',
        height=600,
        showlegend=False,
        template='plotly_white',
        margin=dict(l=100, r=100, t=100, b=50)
    )
    
    # Update x and y axes labels
    fig.update_xaxes(title_text='Fan Speed Mean (RPM)', row=1, col=1)
    fig.update_xaxes(title_text='Fan Speed Mean (RPM)', row=1, col=2)
    fig.update_yaxes(title_text='Frequency', row=1, col=1)
    fig.update_yaxes(title_text='Frequency', row=1, col=2)
    
    # Save HTML file
    html_file = output_path / 'fan_hl_histogram.html'
    fig.write_html(str(html_file))
    
    return str(html_file)


def create_ttemp_hl_histogram(balanced_low_df, balanced_high_df, output_dir='outputs'):
    """
    Create side-by-side histograms comparing interface temperature distributions.
    
    Creates an interactive HTML visualization with separate histograms for low and
    high speed fans' interface temperatures, including statistics.
    
    Args:
        balanced_low_df (pd.DataFrame): Balanced dataframe with low speed fans
        balanced_high_df (pd.DataFrame): Balanced dataframe with high speed fans
        output_dir (str): Directory where the HTML file will be saved
        
    Returns:
        str: Path to the generated HTML file
    """
    # Create output directory if it doesn't exist
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    # Extract interface temperature data
    low_temp_data = balanced_low_df['Interface_Temp'].dropna()
    high_temp_data = balanced_high_df['Interface_Temp'].dropna()
    
    # Calculate statistics
    low_stats = {
        'mean': low_temp_data.mean(),
        'std': low_temp_data.std(),
        'var': low_temp_data.var(),
        'count': len(low_temp_data)
    }
    
    high_stats = {
        'mean': high_temp_data.mean(),
        'std': high_temp_data.std(),
        'var': high_temp_data.var(),
        'count': len(high_temp_data)
    }
    
    # Create subplots
    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=('Low (< 9999 rpm)', 'High (>= 10,000 rpm)'),
        specs=[[{'secondary_y': False}, {'secondary_y': False}]]
    )
    
    # Add low speed histogram
    fig.add_trace(
        go.Histogram(
            x=low_temp_data,
            nbinsx=40,
            name='Low Speed',
            marker=dict(color='#2ca02c', line=dict(color='#1f8f1f', width=1.5)),
            showlegend=False
        ),
        row=1, col=1
    )
    
    # Add high speed histogram
    fig.add_trace(
        go.Histogram(
            x=high_temp_data,
            nbinsx=40,
            name='High Speed',
            marker=dict(color='#d62728', line=dict(color='#8b0000', width=1.5)),
            showlegend=False
        ),
        row=1, col=2
    )
    
    # Create statistics text for low speed
    low_stats_text = (
        f'<b>Low Speed Statistics</b><br>'
        f'Mean: {low_stats["mean"]:.2f}°C<br>'
        f'Std Dev: {low_stats["std"]:.2f}<br>'
        f'Variance: {low_stats["var"]:.2f}<br>'
        f'Count: {low_stats["count"]:,}'
    )
    
    # Create statistics text for high speed
    high_stats_text = (
        f'<b>High Speed Statistics</b><br>'
        f'Mean: {high_stats["mean"]:.2f}°C<br>'
        f'Std Dev: {high_stats["std"]:.2f}<br>'
        f'Variance: {high_stats["var"]:.2f}<br>'
        f'Count: {high_stats["count"]:,}'
    )
    
    # Add annotations for statistics
    fig.add_annotation(
        text=low_stats_text,
        xref='x domain', yref='y domain',
        x=0.98, y=0.97,
        showarrow=False,
        bgcolor='rgba(255, 255, 255, 0.8)',
        bordercolor='#2ca02c',
        borderwidth=2,
        align='left',
        xanchor='right',
        yanchor='top',
        font=dict(size=11),
        row=1, col=1
    )
    
    fig.add_annotation(
        text=high_stats_text,
        xref='x domain', yref='y domain',
        x=0.98, y=0.97,
        showarrow=False,
        bgcolor='rgba(255, 255, 255, 0.8)',
        bordercolor='#d62728',
        borderwidth=2,
        align='left',
        xanchor='right',
        yanchor='top',
        font=dict(size=11),
        row=1, col=2
    )
    
    # Update layout
    fig.update_layout(
        title_text='Interface Temperature Distribution Comparison: Low vs High Speed Fans',
        height=600,
        showlegend=False,
        template='plotly_white',
        margin=dict(l=100, r=100, t=100, b=50)
    )
    
    # Update x and y axes labels
    fig.update_xaxes(title_text='Interface Temperature (°C)', row=1, col=1)
    fig.update_xaxes(title_text='Interface Temperature (°C)', row=1, col=2)
    fig.update_yaxes(title_text='Frequency', row=1, col=1)
    fig.update_yaxes(title_text='Frequency', row=1, col=2)
    
    # Save HTML file
    html_file = output_path / 'ttemp_hl_histogram.html'
    fig.write_html(str(html_file))
    
    return str(html_file)
