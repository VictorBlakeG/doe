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
    - Transceiver_Manufacturer: Nominal (transceiver manufacturer names)
    - Fan_Speed_Range: Categorical (L for low, H for high)
    - Rack_Unit: Continuous (rack unit position)
    
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
    transceiver_mfrs = doe_df['SFP_manufacturer'].unique()
    speed_ranges = ['L', 'H']
    rack_unit_min = doe_df['rack_unit'].min()
    rack_unit_max = doe_df['rack_unit'].max()
    rack_unit_mean = doe_df['rack_unit'].mean()
    
    print("\n" + "="*80)
    print("DESIGN OF EXPERIMENTS SETUP")
    print("="*80)
    print(f"\nFactor Levels:")
    print(f"  Transceiver Manufacturer: {len(transceiver_mfrs)} levels (categorical)")
    print(f"    {list(transceiver_mfrs[:5])}... (+{len(transceiver_mfrs)-5} more)")
    print(f"  Fan Speed Range: {len(speed_ranges)} levels (categorical)")
    print(f"    {speed_ranges}")
    print(f"  Rack Unit: Continuous")
    print(f"    Range: {rack_unit_min:.1f} to {rack_unit_max:.1f}")
    print(f"    Mean: {rack_unit_mean:.2f}")
    
    print(f"\nResponse Variable: Interface_Temp")
    print(f"Total observations: {len(doe_df)}")
    print("="*80 + "\n")
    
    return doe_df


def create_full_factorial_design(doe_df, output_dir='outputs'):
    """
    Create a full-factorial design with discrete rack_unit factor (integer increments 1-42).
    
    Design includes all combinations of:
    - Transceiver Manufacturer (10 levels)
    - Rack Unit (42 discrete values: 1-42)
    - Fan Speed Range (L, H)
    
    Args:
        doe_df (pd.DataFrame): Design dataframe from setup_doe_design
        output_dir (str): Directory for output files
        
    Returns:
        pd.DataFrame: Full-factorial design table
    """
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    # Get unique levels for each factor
    transceiver_mfrs = sorted(doe_df['SFP_manufacturer'].unique())
    speed_ranges = ['L', 'H']
    
    # Create discrete rack_unit levels as integers from 1 to 42
    rack_units = list(range(1, 43))  # Integer increments 1 through 42
    
    # Create full factorial design with all combinations
    design_rows = []
    run_number = 1
    
    for tman in transceiver_mfrs:
        for rack_unit in rack_units:
            for speed in speed_ranges:
                design_rows.append({
                    'Run': run_number,
                    'Transceiver_Manufacturer': tman,
                    'Rack_Unit': rack_unit,
                    'Fan_Speed_Range': speed
                })
                run_number += 1
    
    design_table = pd.DataFrame(design_rows)
    
    # Display design summary
    total_runs = len(transceiver_mfrs) * len(rack_units) * len(speed_ranges)
    print("\n" + "="*80)
    print("FULL FACTORIAL DESIGN TABLE")
    print("="*80)
    print(f"\nDesign Factors (All Discrete/Categorical):")
    print(f"  Transceiver Manufacturers: {len(transceiver_mfrs)} levels")
    print(f"  Rack Unit: {len(rack_units)} discrete levels (integer: 1-42)")
    print(f"  Fan Speed Range: {len(speed_ranges)} levels (L, H)")
    print(f"\nTotal Design Runs: {total_runs}")
    print(f"  ({len(transceiver_mfrs)} × {len(rack_units)} × {len(speed_ranges)})")
    print(f"\nDesign Runs (First 30):")
    print(design_table.head(30).to_string(index=False))
    if len(design_table) > 30:
        print(f"\n... ({len(design_table)-30} more runs)")
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
            h2 {{ color: #666; margin-top: 30px; }}
            table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
            th {{ background-color: #4CAF50; color: white; padding: 10px; text-align: left; }}
            td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
            tr:nth-child(even) {{ background-color: #f2f2f2; }}
            .factor-info {{ background-color: #e8f5e9; padding: 10px; border-radius: 5px; margin: 10px 0; }}
        </style>
    </head>
    <body>
        <h1>Design of Experiments (DOE) - Full Factorial Design</h1>
        
        <h2>Design Structure</h2>
        <div class="factor-info">
            <p><strong>All factors are treated as discrete/categorical:</strong></p>
            <ul>
                <li><strong>Transceiver Manufacturer:</strong> {len(transceiver_mfrs)} discrete levels</li>
                <li><strong>Rack Unit:</strong> {len(rack_units)} discrete levels (integer values: 1 through 42)</li>
                <li><strong>Fan Speed Range:</strong> {len(speed_ranges)} discrete levels (L=Low, H=High)</li>
            </ul>
            <p><strong>Total Design Combinations:</strong> {total_runs} factorial runs</p>
            <p><strong>Response Variable:</strong> Interface Temperature</p>
            <p><strong>Model Terms:</strong> Main effects + 2-way and 3-way interactions</p>
        </div>
        
        <h2>Design Table</h2>
        <p><em>Showing all {total_runs} runs. Each row represents one test condition.</em></p>
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
    Fit a full-factorial model with discrete factors and lack-of-fit test using duplicate observations.
    
    Treats all factors as categorical:
    - Transceiver_Manufacturer: 10 levels
    - Rack_Unit: 42 discrete levels (1-42)
    - Fan_Speed_Range: 2 levels (L, H)
    
    Args:
        doe_df (pd.DataFrame): Design dataframe with observations
        output_dir (str): Directory for output files
        
    Returns:
        tuple: (model, results, summary_stats)
    """
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    # Prepare data for model - encode categorical variables
    model_df = doe_df[['SFP_manufacturer', 'Fan_Speed_Range', 'rack_unit', 'Interface_Temp']].copy()
    model_df.columns = ['Transceiver_Manufacturer', 'Fan_Speed_Range', 'Rack_Unit', 'ttemp']
    
    # Convert Rack_Unit to integer for categorical treatment
    model_df['Rack_Unit'] = model_df['Rack_Unit'].astype(int)
    
    # Fit model with all factors as categorical
    # Formula: response ~ manufacturer + rack_unit + speed + all interactions
    formula = 'ttemp ~ C(Transceiver_Manufacturer) + C(Rack_Unit) + C(Fan_Speed_Range) + C(Transceiver_Manufacturer):C(Rack_Unit) + C(Transceiver_Manufacturer):C(Fan_Speed_Range) + C(Rack_Unit):C(Fan_Speed_Range) + C(Transceiver_Manufacturer):C(Rack_Unit):C(Fan_Speed_Range)'
    
    model = ols(formula, data=model_df)
    results = model.fit()
    
    # Extract ANOVA table using Type I (sequential)
    anova_table = sm.stats.anova_lm(results, typ=1)
    
    # Calculate lack-of-fit test using replicate observations
    lof_table = _calculate_lack_of_fit(model_df, results, anova_table)
    
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
    print("LACK-OF-FIT TEST (Using Duplicate Observations)")
    print("="*80)
    print(lof_table.to_string())
    
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
        'lof_table': lof_table,
        'params': results.params,
        'pvalues': results.pvalues,
        'param_summary': param_summary,
        'conf_int': results.conf_int(alpha=0.05)
    }
    
    # Create HTML report
    create_doe_report(results, anova_table, param_summary, lof_table, output_path)
    
    return model, results, summary_stats


def _calculate_lack_of_fit(model_df, results, anova_table):
    """
    Calculate lack-of-fit test using pure error from replicate observations.
    
    Partitions residual sum of squares into:
    - Pure Error: variation among replicates at same factor level combinations
    - Lack of Fit: deviation from model beyond pure error
    
    Args:
        model_df (pd.DataFrame): Model data with predictors and response
        results: Fitted regression results
        anova_table: Original ANOVA table
        
    Returns:
        pd.DataFrame: Lack-of-fit ANOVA table
    """
    # Group by factor combinations to identify replicates
    group_cols = ['Transceiver_Manufacturer', 'Rack_Unit', 'Fan_Speed_Range']
    grouped = model_df.groupby(group_cols, observed=True)
    
    # Calculate pure error (variation within factor combinations with replicates)
    pure_error_ss = 0
    pure_error_df = 0
    num_replicates = 0
    
    for name, group in grouped:
        n_group = len(group)
        if n_group > 1:  # Only consider groups with replicates
            group_mean = group['ttemp'].mean()
            pure_error_ss += ((group['ttemp'] - group_mean) ** 2).sum()
            pure_error_df += (n_group - 1)
            num_replicates += n_group
    
    # Get residual information from fitted model
    residual_ss = results.ssr  # Sum of squared residuals
    residual_df = results.df_resid
    
    # Calculate lack-of-fit
    lof_ss = residual_ss - pure_error_ss if pure_error_ss > 0 else residual_ss
    lof_df = residual_df - pure_error_df if pure_error_ss > 0 else residual_df
    
    # Calculate mean squares
    pure_error_ms = pure_error_ss / pure_error_df if pure_error_df > 0 else 0
    lof_ms = lof_ss / lof_df if lof_df > 0 else 0
    
    # F-ratio and p-value for lack-of-fit
    lof_f = lof_ms / pure_error_ms if pure_error_ms > 0 else 0
    lof_pvalue = 1 - sm.stats.f.cdf(lof_f, lof_df, pure_error_df) if lof_f > 0 else 1.0
    
    # Create LOF table
    lof_data = {
        'Source': ['Lack of Fit', 'Pure Error', 'Total Error'],
        'df': [lof_df, pure_error_df, residual_df],
        'sum_sq': [lof_ss, pure_error_ss, residual_ss],
        'mean_sq': [lof_ms, pure_error_ms, residual_ss/residual_df],
        'F': [lof_f, np.nan, np.nan],
        'PR(>F)': [lof_pvalue, np.nan, np.nan]
    }
    
    lof_table = pd.DataFrame(lof_data)
    
    # Add summary statistics
    print(f"\nDuplicate Observations Analysis:")
    print(f"  Total observations: {len(model_df)}")
    print(f"  Factor combinations with replicates: {num_replicates}")
    print(f"  Pure Error df: {pure_error_df}")
    print(f"  Pure Error SS: {pure_error_ss:.4f}")
    if lof_f > 0:
        print(f"  LOF F-statistic: {lof_f:.4f}")
        print(f"  LOF p-value: {lof_pvalue:.6f}")
    
    return lof_table
    
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
    Clean model formula: format interactions as * instead of :, keep C() notation for categorical variables.
    
    Args:
        formula_str (str): Raw formula from statsmodels
        
    Returns:
        str: Cleaned formula
    """
    # Keep C() notation but replace : with * for interactions
    import re
    cleaned = formula_str.replace(':', '*')
    return cleaned


def create_doe_report(results, anova_table, param_summary, lof_table, output_path):
    """
    Create an HTML report with model results, tables, visualizations, and model formula.
    
    Args:
        results: Statsmodels regression results
        anova_table: ANOVA results
        param_summary: Parameter summary table
        lof_table: Lack-of-fit ANOVA table
        output_path: Path to output directory
    """
    # Extract and clean model formula (symbolic form)
    symbolic_formula = _clean_formula(str(results.model.formula))
    
    # Create expanded formula showing all parameters with coefficients
    exog_names = results.model.exog_names
    cleaned_names = [_clean_label(name) for name in exog_names]
    coefficients = results.params.values
    
    # Build the formula with coefficients
    formula_terms = []
    for name, coef in zip(cleaned_names, coefficients):
        if name == 'Intercept':
            formula_terms.append(f"{coef:.6f}")
        else:
            sign = "+" if coef >= 0 else "-"
            formula_terms.append(f"{sign} ({abs(coef):.6f})*{name}")
    
    # Create the expanded formula
    model_formula = "ttemp = " + " ".join(formula_terms)
    
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
    
    # Clean parameter summary and format p-values
    param_summary_clean = param_summary.copy()
    param_summary_clean.index = [_clean_label(idx) for idx in param_summary_clean.index]
    
    # Sort by p-value FIRST (before any formatting)
    param_summary_clean = param_summary_clean.sort_values('p-value')
    
    # Add Significance column BEFORE converting p-values to strings
    alpha_risk = 0.05
    param_summary_clean['Significance'] = param_summary_clean['p-value'].apply(
        lambda x: 'Significant' if x < alpha_risk else 'Not Significant'
    )
    
    # Format p-values: convert to decimals without scientific notation, standardize decimal places
    pvalue_decimals = 6  # Number of decimal places
    param_summary_clean['p-value'] = param_summary_clean['p-value'].apply(
        lambda x: f"{x:.{pvalue_decimals}f}"
    )
    
    # Reorder columns to place Significance at the end
    cols = list(param_summary_clean.columns)
    cols = cols[:-1] + ['Significance']  # Move Significance to end
    param_summary_clean = param_summary_clean[cols]
    
    param_summary_html = param_summary_clean.to_html()
    
    # Create leverage plots for each term
    leverage_plots_html = ""
    try:
        import matplotlib.pyplot as plt
        import io
        import base64
        from matplotlib.backends.backend_agg import FigureCanvasAgg
        
        # Get model terms and data
        exog_names = results.model.exog_names
        terms_to_plot = [name for name in exog_names if name != 'const']
        
        if terms_to_plot:
            # Create cleaned names for display
            cleaned_names = [_clean_label(name) for name in terms_to_plot]
            
            # For many terms, create multiple plots in pages (6 plots per page)
            total_terms = len(terms_to_plot)
            plots_per_page = 6
            num_pages = (total_terms + plots_per_page - 1) // plots_per_page
            
            leverage_html_pages = []
            
            # Get model data
            exog = results.model.exog
            residuals = results.resid
            leverage = results.get_influence().hat_matrix_diag
            
            for page in range(num_pages):
                start_idx = page * plots_per_page
                end_idx = min((page + 1) * plots_per_page, total_terms)
                num_plots = end_idx - start_idx
                
                # Create figure with subplots (2 rows x 3 cols)
                fig, axes = plt.subplots(2, 3, figsize=(14, 8))
                axes_flat = axes.flatten()
                
                for plot_idx in range(num_plots):
                    term_idx = start_idx + plot_idx
                    ax = axes_flat[plot_idx]
                    
                    # Get the exog column (add 1 to skip intercept, which is column 0)
                    col_idx = term_idx + 1
                    if col_idx < exog.shape[1]:
                        exog_col = exog[:, col_idx]
                        
                        # Create scatter plot of exog vs residuals
                        ax.scatter(exog_col, residuals, alpha=0.5, s=20, color='steelblue')
                        ax.axhline(y=0, color='r', linestyle='-', linewidth=1)
                        ax.set_xlabel('Predictor', fontsize=7)
                        ax.set_ylabel('Residuals', fontsize=7)
                        ax.set_title(cleaned_names[term_idx], fontsize=7)
                        ax.tick_params(labelsize=7)
                    else:
                        ax.axis('off')
                
                # Hide unused subplots
                for plot_idx in range(num_plots, 6):
                    axes_flat[plot_idx].axis('off')
                
                plt.tight_layout()
                
                # Convert matplotlib figure to HTML
                buf = io.BytesIO()
                canvas = FigureCanvasAgg(fig)
                canvas.print_png(buf)
                buf.seek(0)
                img_base64 = base64.b64encode(buf.read()).decode('utf-8')
                
                page_html = f'<img src="data:image/png;base64,{img_base64}" style="width: 100%; max-width: 1000px; height: auto;" alt="Leverage Plots Page {page+1}">'
                if num_pages > 1:
                    page_html = f'<h4>Plots {start_idx + 1}-{end_idx} of {total_terms}</h4>' + page_html
                leverage_html_pages.append(page_html)
                
                plt.close(fig)
            
            leverage_plots_html = '<br>'.join(leverage_html_pages)
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
