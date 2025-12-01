"""
Design of Experiments (DOE) module for statistical analysis.
Handles factorial design setup, execution, and visualization.
"""
import pandas as pd
import numpy as np
from pathlib import Path
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import statsmodels.api as sm
from statsmodels.formula.api import ols
import warnings
warnings.filterwarnings('ignore')


def setup_doe_design(balanced_low_df, balanced_high_df):
    """
    Set up a design of experiments with three factors.
    
    Factors:
    - Device_Vendor: Nominal (device vendor names)
    - Transceiver_Manufacturer: Nominal (transceiver manufacturer names)
    - Fan_Speed_Range: Categorical (L for low, H for high)
    
    Response: Interface_Temp
    
    Args:
        balanced_low_df (pd.DataFrame): Balanced low speed fan dataframe
        balanced_high_df (pd.DataFrame): Balanced high speed fan dataframe
        
    Returns:
        pd.DataFrame: Full-factorial design dataframe
    """
    # Combine dataframes with speed range indicator
    low_df = balanced_low_df.copy()
    low_df['Fan_Speed_Range'] = 'L'
    
    high_df = balanced_high_df.copy()
    high_df['Fan_Speed_Range'] = 'H'
    
    doe_df = pd.concat([low_df, high_df], ignore_index=True)
    
    # Extract unique levels for each factor
    device_vendors = doe_df['Device Vendor'].unique()
    transceiver_mfrs = doe_df['SFP_manufacturer'].unique()
    speed_ranges = ['L', 'H']
    
    print("\n" + "="*80)
    print("DESIGN OF EXPERIMENTS SETUP")
    print("="*80)
    print(f"\nFactor Levels:")
    print(f"  Device Vendor: {len(device_vendors)} levels")
    print(f"    {device_vendors}")
    print(f"  Transceiver Manufacturer: {len(transceiver_mfrs)} levels")
    print(f"    {list(transceiver_mfrs[:5])}... (+{len(transceiver_mfrs)-5} more)")
    print(f"  Fan Speed Range: {len(speed_ranges)} levels")
    print(f"    {speed_ranges}")
    
    print(f"\nResponse Variable: Interface_Temp")
    print(f"Total observations: {len(doe_df)}")
    print("="*80 + "\n")
    
    return doe_df


def create_full_factorial_design(doe_df, output_dir='outputs'):
    """
    Create a full-factorial design and display design table.
    
    Args:
        doe_df (pd.DataFrame): Design dataframe from setup_doe_design
        output_dir (str): Directory for output files
        
    Returns:
        pd.DataFrame: Full-factorial design table
    """
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    # Get unique levels for each factor
    device_vendors = sorted(doe_df['Device Vendor'].unique())
    transceiver_mfrs = sorted(doe_df['SFP_manufacturer'].unique())
    speed_ranges = ['L', 'H']
    
    # Create full factorial design
    design_rows = []
    run_number = 1
    
    for vendor in device_vendors:
        for tman in transceiver_mfrs:
            for speed in speed_ranges:
                design_rows.append({
                    'Run': run_number,
                    'Device_Vendor': vendor,
                    'Transceiver_Manufacturer': tman,
                    'Fan_Speed_Range': speed
                })
                run_number += 1
    
    design_table = pd.DataFrame(design_rows)
    
    # Display design table
    print("\n" + "="*80)
    print("FULL FACTORIAL DESIGN TABLE")
    print("="*80)
    print(f"\nTotal Design Runs: {len(design_table)}")
    print(f"Factors: {len(device_vendors)} × {len(transceiver_mfrs)} × {len(speed_ranges)}")
    print(f"\nFirst 20 runs:")
    print(design_table.head(20).to_string(index=False))
    print(f"\n... ({len(design_table)-20} more runs)")
    print("="*80 + "\n")
    
    # Save design table as HTML
    html_table = design_table.to_html(index=False)
    html_content = f"""
    <html>
    <head>
        <title>DOE Design Table</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; }}
            h1 {{ color: #333; }}
            table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
            th {{ background-color: #4CAF50; color: white; padding: 10px; text-align: left; }}
            td {{ border: 1px solid #ddd; padding: 8px; }}
            tr:nth-child(even) {{ background-color: #f2f2f2; }}
        </style>
    </head>
    <body>
        <h1>Full Factorial Design of Experiments</h1>
        <p><strong>Total Runs:</strong> {len(design_table)}</p>
        <p><strong>Factors:</strong> Device Vendor × Transceiver Manufacturer × Fan Speed Range</p>
        <p><strong>Response:</strong> Interface Temperature</p>
        {html_table}
    </body>
    </html>
    """
    
    html_file = output_path / 'doe_design.html'
    with open(html_file, 'w') as f:
        f.write(html_content)
    
    print(f"Design table saved to: {html_file}\n")
    
    return design_table


def fit_doe_model(doe_df, output_dir='outputs'):
    """
    Fit a full-factorial model with main effects and two-way interactions.
    
    Args:
        doe_df (pd.DataFrame): Design dataframe
        output_dir (str): Directory for output files
        
    Returns:
        tuple: (model, results, summary_stats)
    """
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    # Prepare data for model - encode categorical variables
    model_df = doe_df[['Device Vendor', 'SFP_manufacturer', 'Fan_Speed_Range', 'Interface_Temp']].copy()
    model_df.columns = ['Device_Vendor', 'Transceiver_Manufacturer', 'Fan_Speed_Range', 'ttemp']
    
    # Fit model: response ~ all factors + two-way interactions
    formula = 'ttemp ~ C(Transceiver_Manufacturer) + C(Fan_Speed_Range) + C(Transceiver_Manufacturer):C(Fan_Speed_Range)'
    
    model = ols(formula, data=model_df)
    results = model.fit()
    
    # Extract ANOVA table using Type I (sequential)
    anova_table = sm.stats.anova_lm(results, typ=1)
    
    print("\n" + "="*80)
    print("DESIGN OF EXPERIMENTS MODEL FIT RESULTS")
    print("="*80)
    
    print("\nModel Summary:")
    print(f"  R-squared: {results.rsquared:.4f}")
    print(f"  Adjusted R-squared: {results.rsquared_adj:.4f}")
    print(f"  F-statistic: {results.fvalue:.4f}")
    print(f"  Prob (F-statistic): {results.f_pvalue:.6f}")
    print(f"  Residual Std Error: {np.sqrt(results.mse_resid):.4f}")
    print(f"  Degrees of Freedom: {results.df_resid}")
    
    print("\nANOVA Table (Type I - Sequential):")
    print(anova_table.to_string())
    
    print("\n" + "="*80)
    print("PARAMETER ESTIMATES (95% Confidence Interval)")
    print("="*80)
    
    param_summary = pd.DataFrame({
        'Coefficient': results.params,
        'Std Error': results.bse,
        't-value': results.tvalues,
        'p-value': results.pvalues,
        'Lower CI': results.conf_int()[0],
        'Upper CI': results.conf_int()[1]
    })
    
    print(param_summary.to_string())
    print("="*80 + "\n")
    
    # Create summary statistics
    summary_stats = {
        'r_squared': results.rsquared,
        'adj_r_squared': results.rsquared_adj,
        'f_statistic': results.fvalue,
        'f_pvalue': results.f_pvalue,
        'anova_table': anova_table,
        'params': results.params,
        'pvalues': results.pvalues,
        'param_summary': param_summary,
        'conf_int': results.conf_int(alpha=0.05)
    }
    
    # Create HTML report
    create_doe_report(results, anova_table, param_summary, output_path)
    
    return model, results, summary_stats


def _clean_label(label):
    """
    Clean parameter labels: remove C() notation and format interactions as a*b.
    
    Args:
        label (str): Raw label from statsmodels
        
    Returns:
        str: Cleaned label
    """
    # Handle interaction terms: C(A)[T.x]:C(B)[T.y] -> A*B[x,y]
    if ':' in label:
        parts = label.split(':')
        factor1 = parts[0].split('[')[0].replace('C(', '').replace(')', '')
        factor2 = parts[1].split('[')[0].replace('C(', '').replace(')', '')
        
        # Extract level names if present
        level1 = parts[0].split('[T.')[-1].rstrip(']') if '[T.' in parts[0] else ''
        level2 = parts[1].split('[T.')[-1].rstrip(']') if '[T.' in parts[1] else ''
        
        if level1 and level2:
            return f"{factor1}*{factor2}[{level1},{level2}]"
        else:
            return f"{factor1}*{factor2}"
    else:
        # Handle main effect: C(A)[T.x] -> A[x]
        factor = label.split('[')[0].replace('C(', '').replace(')', '')
        level = label.split('[T.')[-1].rstrip(']') if '[T.' in label else ''
        
        if level:
            return f"{factor}[{level}]"
        else:
            return factor


def _clean_formula(formula_str):
    """
    Clean model formula: remove C() notation and format interactions as a*b.
    
    Args:
        formula_str (str): Raw formula from statsmodels
        
    Returns:
        str: Cleaned formula
    """
    # Replace C(x) with x
    import re
    cleaned = re.sub(r'C\(([^)]+)\)', r'\1', formula_str)
    # Replace : with * for interactions
    cleaned = cleaned.replace(':', '*')
    return cleaned


def create_doe_report(results, anova_table, param_summary, output_path):
    """
    Create an HTML report with model results, tables, visualizations, and model formula.
    
    Args:
        results: Statsmodels regression results
        anova_table: ANOVA results
        param_summary: Parameter summary table
        output_path: Path to output directory
    """
    # Extract and clean model formula
    model_formula = _clean_formula(str(results.model.formula))
    
    # Calculate predicted response by Fan Speed Range
    # Extract unique transceiver manufacturers
    tman_levels = [param_name.split('[T.')[-1].rstrip(']') 
                   for param_name in results.params.index 
                   if 'Transceiver_Manufacturer' in param_name and ':' not in param_name]
    
    # Get the model predictions for different scenarios
    speed_responses = {'L': [], 'H': []}
    speed_responses_mean = {}
    
    for speed in ['L', 'H']:
        # Calculate average response across transceiver manufacturers
        intercept = results.params['Intercept']
        speed_effect = results.params.get(f'C(Fan_Speed_Range)[T.{speed}]', 0)
        avg_response = intercept + speed_effect
        speed_responses_mean[speed] = avg_response
        
        # Calculate response for each transceiver manufacturer
        for tman in tman_levels:
            tman_effect = results.params.get(f'C(Transceiver_Manufacturer)[T.{tman}]', 0)
            interaction_effect = results.params.get(f'C(Transceiver_Manufacturer)[T.{tman}]:C(Fan_Speed_Range)[T.{speed}]', 0)
            response = intercept + tman_effect + speed_effect + interaction_effect
            speed_responses[speed].append(response)
    
    # Create Plotly figure for fan speed response
    fig = go.Figure()
    
    # Add box plot showing response distribution by fan speed
    speeds = ['L', 'H']
    speed_labels = {'L': 'Low Speed', 'H': 'High Speed'}
    colors = {'L': '#FF6B6B', 'H': '#4ECDC4'}
    
    for speed in speeds:
        fig.add_trace(go.Box(
            y=speed_responses[speed],
            name=speed_labels[speed],
            marker=dict(color=colors[speed]),
            boxmean='sd',
            showlegend=True
        ))
    
    fig.update_layout(
        title='Model Response by Fan Speed Range',
        yaxis_title='Predicted Interface Temperature (°C)',
        xaxis_title='Fan Speed Range',
        template='plotly_white',
        height=500,
        showlegend=True
    )
    
    # Convert figure to HTML
    response_plot_html = fig.to_html(full_html=False, include_plotlyjs='cdn')
    
    # Create Actual by Predicted plot with confidence interval
    # Get actual values and predictions
    actual = results.model.endog
    predicted = results.fittedvalues
    residuals = results.resid
    std_error = np.sqrt(results.mse_resid)
    
    # Calculate 95% confidence interval
    ci_upper = predicted + 1.96 * std_error
    ci_lower = predicted - 1.96 * std_error
    
    # Sort by predicted values for smooth confidence interval
    sort_idx = np.argsort(predicted)
    predicted_sorted = predicted.iloc[sort_idx]
    ci_lower_sorted = ci_lower.iloc[sort_idx]
    ci_upper_sorted = ci_upper.iloc[sort_idx]
    
    # Create Actual by Predicted plot
    fig_ap = go.Figure()
    
    # Add lower confidence interval line
    fig_ap.add_trace(go.Scatter(
        x=predicted_sorted,
        y=ci_lower_sorted,
        fill=None,
        mode='lines',
        line=dict(color='red', width=1),
        name='95% CI Lower',
        showlegend=False
    ))
    
    # Add upper confidence interval line with fill (red shaded area)
    fig_ap.add_trace(go.Scatter(
        x=predicted_sorted,
        y=ci_upper_sorted,
        fill='tonexty',
        mode='lines',
        line=dict(color='red', width=1),
        fillcolor='rgba(255,0,0,0.2)',
        name='95% Confidence Interval',
        showlegend=True
    ))
    
    # Add actual data points (black)
    fig_ap.add_trace(go.Scatter(
        x=predicted,
        y=actual,
        mode='markers',
        marker=dict(color='black', size=4, opacity=0.6),
        name='Actual Data Points',
        showlegend=True
    ))
    
    # Add diagonal reference line (perfect prediction)
    min_val = min(predicted.min(), actual.min())
    max_val = max(predicted.max(), actual.max())
    fig_ap.add_trace(go.Scatter(
        x=[min_val, max_val],
        y=[min_val, max_val],
        mode='lines',
        line=dict(color='blue', dash='dash', width=2),
        name='Perfect Prediction',
        showlegend=True
    ))
    
    fig_ap.update_layout(
        title='Actual by Predicted Plot',
        xaxis_title='Predicted Interface Temperature (°C)',
        yaxis_title='Actual Interface Temperature (°C)',
        template='plotly_white',
        height=600,
        width=700,
        hovermode='closest'
    )
    
    # Convert figure to HTML
    actual_predicted_plot_html = fig_ap.to_html(full_html=False, include_plotlyjs=False)
    
    # Clean ANOVA table index and columns
    anova_table_clean = anova_table.copy()
    anova_table_clean.index = [_clean_label(idx) for idx in anova_table_clean.index]
    anova_table_html = anova_table_clean.to_html()
    
    # Clean parameter summary
    param_summary_clean = param_summary.copy()
    param_summary_clean.index = [_clean_label(idx) for idx in param_summary_clean.index]
    param_summary_html = param_summary_clean.to_html()
    
    # Create leverage plots for each term
    leverage_plots_html = ""
    try:
        from statsmodels.graphics.api import plot_partregress_grid
        import matplotlib.pyplot as plt
        
        # Get exog data and model terms
        exog_data = results.model.exog
        exog_names = results.model.exog_names
        
        # Filter out the intercept and create individual leverage plots
        terms_to_plot = [name for name in exog_names if name != 'const']
        
        if terms_to_plot:
            # Create cleaned names for display
            cleaned_names = [_clean_label(name) for name in terms_to_plot]
            
            # Create a grid of leverage plots (2x2 or 2x3 depending on number of terms)
            fig_leverage = plot_partregress_grid(
                results,
                fig=None,
                exog_idx=list(range(1, min(len(terms_to_plot) + 1, 7)))  # Up to 6 plots
            )
            
            # Clean up labels and reduce font sizes for readability
            for idx, ax in enumerate(fig_leverage.get_axes()):
                ax.tick_params(labelsize=7)
                ax.xaxis.label.set_fontsize(7)
                ax.yaxis.label.set_fontsize(7)
                ax.title.set_fontsize(7)
                
                # Set the cleaned label
                if idx < len(cleaned_names):
                    ax.set_title(cleaned_names[idx])
            
            # Convert matplotlib figure to HTML
            import io
            import base64
            from matplotlib.backends.backend_agg import FigureCanvasAgg
            
            buf = io.BytesIO()
            canvas = FigureCanvasAgg(fig_leverage)
            canvas.print_png(buf)
            buf.seek(0)
            img_base64 = base64.b64encode(buf.read()).decode('utf-8')
            leverage_plots_html = f'<img src="data:image/png;base64,{img_base64}" style="width: 100%; max-width: 900px; height: auto;" alt="Leverage Plots">'
            
            plt.close(fig_leverage)
    except Exception as e:
        print(f"Warning: Could not generate leverage plots: {e}")
        leverage_plots_html = f'<p style="color: #666;">Leverage plots could not be generated: {str(e)}</p>'
    
    # Create summary table
    summary_html = f"""
    <html>
    <head>
        <title>DOE Analysis Report</title>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }}
            h1, h2 {{ color: #333; }}
            .section {{ background-color: white; padding: 20px; margin: 20px 0; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
            table {{ border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 12px; }}
            th {{ background-color: #4CAF50; color: white; padding: 10px; text-align: left; }}
            td {{ border: 1px solid #ddd; padding: 8px; }}
            tr:nth-child(even) {{ background-color: #f9f9f9; }}
            .metric {{ display: inline-block; margin: 10px 20px 10px 0; }}
            .metric-value {{ font-size: 24px; font-weight: bold; color: #4CAF50; }}
            .metric-label {{ font-size: 12px; color: #666; }}
            .formula {{ background-color: #f0f0f0; padding: 15px; border-left: 4px solid #4CAF50; margin: 10px 0; font-family: monospace; font-size: 14px; }}
            .plot-container {{ margin: 20px 0; }}
        </style>
    </head>
    <body>
        <h1>Design of Experiments Analysis Report</h1>
        
        <div class="section">
            <h2>Actual by Predicted Plot</h2>
            <div class="plot-container">
                {actual_predicted_plot_html}
            </div>
        </div>
        
        <div class="section">
            <h2>Model Formula</h2>
            <div class="formula">{model_formula}</div>
        </div>
        
        <div class="section">
            <h2>Model Summary</h2>
            <div class="metric">
                <div class="metric-value">{results.rsquared:.4f}</div>
                <div class="metric-label">R-squared</div>
            </div>
            <div class="metric">
                <div class="metric-value">{results.rsquared_adj:.4f}</div>
                <div class="metric-label">Adjusted R-squared</div>
            </div>
            <div class="metric">
                <div class="metric-value">{results.fvalue:.4f}</div>
                <div class="metric-label">F-statistic</div>
            </div>
            <div class="metric">
                <div class="metric-value">{results.f_pvalue:.6f}</div>
                <div class="metric-label">Prob (F-statistic)</div>
            </div>
        </div>
        
        <div class="section">
            <h2>Model Response by Fan Speed</h2>
            <div class="plot-container">
                {response_plot_html}
            </div>
            <table>
                <tr>
                    <th>Fan Speed Range</th>
                    <th>Average Predicted Temperature (°C)</th>
                </tr>
                <tr>
                    <td>Low Speed (L)</td>
                    <td>{speed_responses_mean['L']:.4f}</td>
                </tr>
                <tr>
                    <td>High Speed (H)</td>
                    <td>{speed_responses_mean['H']:.4f}</td>
                </tr>
            </table>
        </div>
        
        <div class="section">
            <h2>ANOVA Table (Type I - Sequential)</h2>
            {anova_table_html}
        </div>
        
        <div class="section">
            <h2>Parameter Estimates (95% Confidence Interval)</h2>
            {param_summary_html}
        </div>
        
        <div class="section">
            <h2>Leverage Plots</h2>
            <p>Leverage plots show the relationship between each variable and the response, controlling for other variables in the model.</p>
            <div class="plot-container">
                {leverage_plots_html}
            </div>
        </div>
    </body>
    </html>
    """
    
    html_file = output_path / 'doe_analysis_report.html'
    with open(html_file, 'w') as f:
        f.write(summary_html)
    
    print(f"Analysis report saved to: {html_file}\n")
