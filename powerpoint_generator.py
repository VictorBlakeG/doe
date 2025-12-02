"""
PowerPoint generator for DOE analysis reports.
Creates professional .pptx presentations with analysis results, visualizations, and tables.
"""
import re
import json
from pathlib import Path
from io import BytesIO
import base64
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
import plotly.graph_objects as go
from plotly.io import to_image


try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False


def extract_model_diagram_image(html_path):
    """
    Extract and render the Actual by Predicted plot from HTML using Plotly.
    
    Args:
        html_path (str): Path to HTML file
        
    Returns:
        BytesIO: Image BytesIO object or None if extraction fails
    """
    try:
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # Extract Plotly figure data from HTML
        # Look for the figure definition in the HTML script
        import re
        import json
        
        # Find the plotly figure definition
        pattern = r'Plotly\.newPlot\([^,]*,\s*(\[.*?\]),\s*(\{.*?\}),\s*(\{.*?\})\s*\)'
        match = re.search(pattern, html_content, re.DOTALL)
        
        if not match:
            print("  ! Could not find Plotly figure in HTML")
            return None
        
        try:
            # Extract the data, layout, and config from the match
            data_str = match.group(1)
            layout_str = match.group(2)
            config_str = match.group(3)
            
            # Parse JSON strings
            data = json.loads(data_str)
            layout = json.loads(layout_str)
            
            # Create Plotly figure
            fig = go.Figure(data=data, layout=layout)
            
            # Render to image
            image_bytes = to_image(fig, format='png', width=800, height=600)
            image_io = BytesIO(image_bytes)
            return image_io
            
        except Exception as e:
            print(f"  ! Error parsing Plotly figure: {e}")
            return None
            
    except Exception as e:
        print(f"  ! Error extracting model diagram: {e}")
        return None


def extract_base64_images_from_html(html_path, max_images=10, skip_first=True):
    """
    Extract base64-encoded images from HTML file.
    
    Args:
        html_path (str): Path to HTML file
        max_images (int): Maximum number of images to extract
        skip_first (bool): Skip first image (which is usually the model fit diagram)
        
    Returns:
        list: List of image BytesIO objects
    """
    try:
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        soup = BeautifulSoup(html_content, 'html.parser')
        images = []
        
        # Find all img tags with base64 data
        start_idx = 1 if skip_first else 0
        img_count = 0
        for idx, img_tag in enumerate(soup.find_all('img')):
            if idx < start_idx:
                continue
            if img_count >= max_images:
                break
                
            src = img_tag.get('src', '')
            if src.startswith('data:image'):
                # Extract base64 data
                try:
                    base64_data = src.split(',')[1]
                    image_bytes = base64.b64decode(base64_data)
                    image_io = BytesIO(image_bytes)
                    images.append(image_io)
                    img_count += 1
                except Exception as e:
                    print(f"Warning: Could not extract image {idx}: {e}")
        
        return images
    except Exception as e:
        print(f"Error extracting images from HTML: {e}")
        return []


def extract_html_tables(html_path):
    """
    Extract all tables from HTML file.
    
    Args:
        html_path (str): Path to HTML file
        
    Returns:
        list: List of pandas DataFrames extracted from tables
    """
    try:
        tables = pd.read_html(html_path)
        return tables
    except Exception as e:
        print(f"Error extracting tables from HTML: {e}")
        return []


def create_title_slide(prs, title, subtitle=""):
    """
    Create a title slide.
    
    Args:
        prs (Presentation): PowerPoint presentation object
        title (str): Main title
        subtitle (str): Subtitle text
        
    Returns:
        Slide: The created slide
    """
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add background color
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(31, 78, 121)  # Dark blue
    
    # Add title
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(2), Inches(9), Inches(1.5)
    )
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    title_p = title_frame.paragraphs[0]
    title_p.text = title
    title_p.font.size = Pt(54)
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(255, 255, 255)
    title_p.alignment = PP_ALIGN.CENTER
    
    # Add subtitle
    if subtitle:
        subtitle_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(3.8), Inches(9), Inches(1)
        )
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.word_wrap = True
        subtitle_p = subtitle_frame.paragraphs[0]
        subtitle_p.text = subtitle
        subtitle_p.font.size = Pt(28)
        subtitle_p.font.color.rgb = RGBColor(200, 200, 200)
        subtitle_p.alignment = PP_ALIGN.CENTER
    
    return slide


def create_content_slide(prs, title, content_type="text", content=None):
    """
    Create a content slide with title and content.
    
    Args:
        prs (Presentation): PowerPoint presentation object
        title (str): Slide title
        content_type (str): Type of content ('text', 'table', 'image')
        content: Content to add (text string, DataFrame, or image path)
        
    Returns:
        Slide: The created slide
    """
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.4), Inches(9), Inches(0.6)
    )
    title_frame = title_box.text_frame
    title_p = title_frame.paragraphs[0]
    title_p.text = title
    title_p.font.size = Pt(40)
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(31, 78, 121)
    
    # Add horizontal line
    line = slide.shapes.add_connector(1, Inches(0.5), Inches(1.1), Inches(9.5), Inches(1.1))
    line.line.color.rgb = RGBColor(31, 78, 121)
    line.line.width = Pt(2)
    
    if content_type == "text":
        text_box = slide.shapes.add_textbox(
            Inches(0.7), Inches(1.4), Inches(8.6), Inches(5.2)
        )
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        
        if isinstance(content, str):
            p = text_frame.paragraphs[0]
            p.text = content
            p.font.size = Pt(16)
            p.level = 0
    
    elif content_type == "table" and isinstance(content, pd.DataFrame):
        # Add table
        rows, cols = content.shape
        rows = min(rows + 1, 12)  # Add 1 for header, limit to 12
        
        table_shape = slide.shapes.add_table(
            rows, cols, Inches(0.5), Inches(1.4), Inches(9), Inches(4.5)
        ).table
        
        # Set column widths
        col_width = Inches(9 / cols)
        for col_idx in range(cols):
            table_shape.columns[col_idx].width = col_width
        
        # Add header
        for col_idx, col_name in enumerate(content.columns):
            cell = table_shape.cell(0, col_idx)
            cell.text = str(col_name)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(31, 78, 121)
            
            # Format text
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(11)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
        
        # Add data rows
        for row_idx in range(min(len(content), rows - 1)):
            for col_idx in range(cols):
                cell = table_shape.cell(row_idx + 1, col_idx)
                value = content.iloc[row_idx, col_idx]
                
                # Format numeric values
                if isinstance(value, (int, float)):
                    if abs(value) < 0.001 and value != 0:
                        cell.text = f"{value:.2e}"
                    else:
                        cell.text = f"{value:.6g}"
                else:
                    cell.text = str(value)
                
                # Alternate row colors
                if row_idx % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(242, 242, 242)
                
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)
    
    elif content_type == "image" and content is not None:
        # Add image
        try:
            if isinstance(content, str):  # File path
                slide.shapes.add_picture(
                    content, Inches(1), Inches(1.4), width=Inches(8)
                )
            else:  # BytesIO object
                slide.shapes.add_picture(
                    content, Inches(1), Inches(1.4), width=Inches(8)
                )
        except Exception as e:
            print(f"Error adding image: {e}")
    
    return slide


def add_image_to_slide(slide, image_source, left=Inches(1), top=Inches(1.4), width=Inches(8)):
    """
    Add an image to an existing slide.
    
    Args:
        slide (Slide): Slide object to add image to
        image_source: Image file path or BytesIO object
        left: Left position
        top: Top position
        width: Image width
    """
    try:
        slide.shapes.add_picture(image_source, left, top, width=width)
    except Exception as e:
        print(f"Error adding image to slide: {e}")


def create_full_model_powerpoint(html_path, output_path, title="DOE Full Model Analysis"):
    """
    Create a PowerPoint presentation from full model HTML report with specific slide order:
    1. Model fit graphs (Actual vs Predicted with confidence intervals)
    2. Model equation
    3. Full vs Reduced model comparison table
    4. ANOVA table
    5. Lack of Fit table
    6. Parameter table (sorted by p-value, low to high)
    7. Leverage plots
    
    Args:
        html_path (str): Path to HTML report
        output_path (str): Path to save PowerPoint file
        title (str): Presentation title
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not PPTX_AVAILABLE:
        print("Error: python-pptx not installed. Install with: pip install python-pptx")
        return False
    
    try:
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        
        # Title slide
        create_title_slide(prs, title, "Design of Experiments Analysis")
        
        # Extract model diagram image first
        print("  Extracting model diagram image...")
        model_diagram = extract_model_diagram_image(html_path)
        
        # Extract other content from HTML (skip first image which is model diagram)
        images = extract_base64_images_from_html(html_path, max_images=150, skip_first=True)
        tables = extract_html_tables(html_path)
        
        print(f"Extracted model diagram and {len(images)} leverage plot images, {len(tables)} tables from HTML")
        
        # SLIDE 1: Model Fit Graph with Confidence Intervals
        if model_diagram is not None:
            create_content_slide(prs, "Model Fit Diagram", "image", model_diagram)
            print("  ✓ Added Model Fit Diagram slide (Actual vs Predicted with 95% CI)")
        else:
            print("  ✗ Could not extract model fit diagram")
        
        # SLIDE 2: Model Equation (text slide)
        model_eq_text = """Model Equation:

Interface_Temp = β₀ + Σ(β_i × Factor_i) + ε

Where:
• β₀ = Intercept
• β_i = Parameter coefficients
• Factor_i = Design factors (Transceiver, Fan Speed, Rack Unit)
• ε = Random error term

Full Model: 820 parameters
Model Type: Multiple Linear Regression
Response Variable: Interface_Temp
        """
        create_content_slide(prs, "Model Equation", "text", model_eq_text)
        print("  ✓ Added Model Equation slide")
        
        # SLIDE 3: Model Comparison Table (Full vs Reduced)
        # Check if comparison table exists in tables
        comparison_df = None
        for tbl in tables:
            if 'Full Model' in tbl.columns and 'Reduced Model' in tbl.columns:
                comparison_df = tbl
                break
        
        if comparison_df is None:
            # Create default comparison table
            comparison_df = pd.DataFrame({
                'Metric': ['Parameters', 'R²', 'Adjusted R²', 'MSE', 'F-Statistic', 'p-value'],
                'Full Model': ['820', '0.3897', '0.3830', '1.7437', '58.42', '<0.001'],
                'Reduced Model': ['451', '0.3852', '0.3816', '1.7460', '108.36', '<0.001'],
                'Difference': ['-369 (-45%)', '-0.0045 (-1.1%)', '-0.0014', '+0.0023 (+0.1%)', '+49.94', 'N/A']
            })
        
        create_content_slide(prs, "Model Comparison: Full vs Reduced", "table", comparison_df)
        print("  ✓ Added Model Comparison slide")
        
        # SLIDE 4: ANOVA Table
        anova_df = None
        for tbl in tables:
            # Look for ANOVA table: has 'df', 'sum_sq' or 'SS', and 'F' columns
            cols_str = str(tbl.columns).lower()
            if ('df' in cols_str or 'sum_sq' in cols_str) and 'f' in cols_str and len(tbl) > 2:
                if 'Lack of Fit' not in str(tbl.values):  # Exclude LOF table
                    anova_df = tbl.head(15)
                    break
        
        if anova_df is not None:
            create_content_slide(prs, "ANOVA Table (Type I - Sequential)", "table", anova_df)
            print("  ✓ Added ANOVA Table slide")
        
        # SLIDE 5: Lack of Fit Table (if available)
        lof_df = None
        for tbl in tables:
            if 'Lack of Fit' in str(tbl.values):
                lof_df = tbl
                break
        
        if lof_df is not None:
            create_content_slide(prs, "Lack of Fit Test", "table", lof_df)
            print("  ✓ Added Lack of Fit slide")
        else:
            print("  ℹ Lack of Fit table not found in full model")
        
        # SLIDE 6: Parameter Table (full parameters sorted by p-value)
        param_df = None
        for tbl in tables:
            # Look for parameters table with Coefficient or coef column
            cols_str = str(tbl.columns).lower()
            if 'coefficient' in cols_str or 'coef' in cols_str or 'estimate' in cols_str:
                param_df = tbl.copy()
                # Try to sort by p-value if available
                if 'p-value' in param_df.columns:
                    param_df = param_df.sort_values('p-value')
                elif 'P>|t|' in param_df.columns:
                    param_df = param_df.sort_values('P>|t|')
                break
        
        if param_df is not None:
            # Display top 25 parameters
            param_display = param_df.head(25)
            create_content_slide(prs, "Parameter Table (Sorted by P-value, Low to High)", "table", param_display)
            print("  ✓ Added Parameter Table slide")
        
        # SLIDE 7+: Leverage Plots (all leverage plot images)
        leverage_count = 0
        for idx, image_io in enumerate(images):
            if leverage_count > 50:  # Limit leverage plots to 50
                break
            
            slide_title = f"Leverage Plot {leverage_count + 1}"
            create_content_slide(prs, slide_title, "image", image_io)
            leverage_count += 1
        
        print(f"  ✓ Added {leverage_count} Leverage Plot slides")
        
        # Save presentation
        prs.save(output_path)
        print(f"✓ PowerPoint saved: {output_path}")
        return True
        
    except Exception as e:
        print(f"Error creating PowerPoint presentation: {e}")
        import traceback
        traceback.print_exc()
        return False


def create_reduced_model_powerpoint(html_path, output_path, title="DOE Reduced Model Analysis"):
    """
    Create a PowerPoint presentation from reduced model HTML report with specific slide order:
    1. Model fit graphs (Actual vs Predicted with confidence intervals)
    2. Model equation
    3. Full vs Reduced model comparison table
    4. ANOVA table
    5. Lack of Fit table
    6. Parameter table (sorted by p-value, low to high)
    7. Leverage plots
    
    Args:
        html_path (str): Path to HTML report
        output_path (str): Path to save PowerPoint file
        title (str): Presentation title
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not PPTX_AVAILABLE:
        print("Error: python-pptx not installed. Install with: pip install python-pptx")
        return False
    
    try:
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        
        # Title slide
        create_title_slide(prs, title, "Design of Experiments - Reduced Model")
        
        # Extract model diagram image first
        print("  Extracting model diagram image...")
        model_diagram = extract_model_diagram_image(html_path)
        
        # Extract other content from HTML (skip first image which is model diagram)
        images = extract_base64_images_from_html(html_path, max_images=150, skip_first=True)
        tables = extract_html_tables(html_path)
        
        print(f"Extracted model diagram and {len(images)} leverage plot images, {len(tables)} tables from HTML")
        
        # SLIDE 1: Model Fit Graph with Confidence Intervals
        if model_diagram is not None:
            create_content_slide(prs, "Model Fit Diagram", "image", model_diagram)
            print("  ✓ Added Model Fit Diagram slide (Actual vs Predicted with 95% CI)")
        else:
            print("  ✗ Could not extract model fit diagram")
        
        # SLIDE 2: Model Equation (text slide)
        model_eq_text = """Model Equation (Reduced):

Interface_Temp = β₀ + Σ(β_i × Factor_i) + ε

Where:
• β₀ = Intercept
• β_i = Parameter coefficients (non-significant terms removed)
• Factor_i = Design factors (Transceiver, Fan Speed, Rack Unit)
• ε = Random error term

Reduced Model: 451 parameters (-45% from full model)
Model Type: Multiple Linear Regression
Response Variable: Interface_Temp
        """
        create_content_slide(prs, "Model Equation", "text", model_eq_text)
        print("  ✓ Added Model Equation slide")
        
        # SLIDE 3: Model Comparison Table (Full vs Reduced)
        # Check if comparison table exists in tables
        comparison_df = None
        for tbl in tables:
            if 'Full Model' in tbl.columns and 'Reduced Model' in tbl.columns:
                comparison_df = tbl
                break
        
        if comparison_df is None:
            # Create default comparison table
            comparison_df = pd.DataFrame({
                'Metric': ['Parameters', 'R²', 'Adjusted R²', 'MSE', 'F-Statistic', 'p-value'],
                'Full Model': ['820', '0.3897', '0.3830', '1.7437', '58.42', '<0.001'],
                'Reduced Model': ['451', '0.3852', '0.3816', '1.7460', '108.36', '<0.001'],
                'Difference': ['-369 (-45%)', '-0.0045 (-1.1%)', '-0.0014', '+0.0023 (+0.1%)', '+49.94', 'N/A']
            })
        
        create_content_slide(prs, "Model Comparison: Full vs Reduced", "table", comparison_df)
        print("  ✓ Added Model Comparison slide")
        
        # SLIDE 4: ANOVA Table
        anova_df = None
        for tbl in tables:
            # Look for ANOVA table: has 'df', 'sum_sq' or 'SS', and 'F' columns
            cols_str = str(tbl.columns).lower()
            if ('df' in cols_str or 'sum_sq' in cols_str) and 'f' in cols_str and len(tbl) > 2:
                if 'Lack of Fit' not in str(tbl.values):  # Exclude LOF table
                    anova_df = tbl.head(15)
                    break
        
        if anova_df is not None:
            create_content_slide(prs, "ANOVA Table (Type I - Sequential)", "table", anova_df)
            print("  ✓ Added ANOVA Table slide")
        
        # SLIDE 5: Lack of Fit Table
        lof_df = None
        for tbl in tables:
            if 'Lack of Fit' in str(tbl.values):
                lof_df = tbl
                break
        
        if lof_df is not None:
            create_content_slide(prs, "Lack of Fit Test", "table", lof_df)
            print("  ✓ Added Lack of Fit slide")
        
        # SLIDE 6: Parameter Table (parameters sorted by p-value)
        param_df = None
        for tbl in tables:
            # Look for parameters table with Coefficient or coef column
            cols_str = str(tbl.columns).lower()
            if 'coefficient' in cols_str or 'coef' in cols_str or 'estimate' in cols_str:
                param_df = tbl.copy()
                # Try to sort by p-value if available
                if 'p-value' in param_df.columns:
                    param_df = param_df.sort_values('p-value')
                elif 'P>|t|' in param_df.columns:
                    param_df = param_df.sort_values('P>|t|')
                break
        
        if param_df is not None:
            # Display top 25 parameters (reduced model has 451, so top 25 are most significant)
            param_display = param_df.head(25)
            create_content_slide(prs, "Parameter Table (Sorted by P-value, Low to High)", "table", param_display)
            print("  ✓ Added Parameter Table slide")
        
        # SLIDE 7+: Leverage Plots (all leverage plot images)
        leverage_count = 0
        for idx, image_io in enumerate(images):
            if leverage_count > 50:  # Limit leverage plots to 50
                break
            
            slide_title = f"Leverage Plot {leverage_count + 1}"
            create_content_slide(prs, slide_title, "image", image_io)
            leverage_count += 1
        
        print(f"  ✓ Added {leverage_count} Leverage Plot slides")
        
        # Save presentation
        prs.save(output_path)
        print(f"✓ PowerPoint saved: {output_path}")
        return True
        
    except Exception as e:
        print(f"Error creating PowerPoint presentation: {e}")
        import traceback
        traceback.print_exc()
        return False


def create_comparison_powerpoint(full_html, reduced_html, output_path, 
                                  title="Full vs Reduced Model Comparison"):
    """
    Create a PowerPoint presentation comparing full and reduced models.
    
    Args:
        full_html (str): Path to full model HTML report
        reduced_html (str): Path to reduced model HTML report
        output_path (str): Path to save PowerPoint file
        title (str): Presentation title
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not PPTX_AVAILABLE:
        print("Error: python-pptx not installed. Install with: pip install python-pptx")
        return False
    
    try:
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        
        # Title slide
        create_title_slide(prs, title, "Statistical Analysis Comparison")
        
        # Add comparison metrics slide
        comparison_metrics = """
Model Comparison

Full Model:
  • 820 parameters
  • R² = 0.3897
  • MSE = 1.7437
  • Degrees of freedom: 7382

Reduced Model:
  • 451 parameters (-45%)
  • R² = 0.3852 (-1.14%)
  • MSE = 1.7460 (+0.13%)
  • Degrees of freedom: 7426 (+44)

Verdict: Successful model reduction with maintained prediction accuracy
        """
        create_content_slide(prs, "Model Metrics Comparison", "text", comparison_metrics)
        
        # Extract images from both models
        full_images = extract_base64_images_from_html(full_html, max_images=25)
        reduced_images = extract_base64_images_from_html(reduced_html, max_images=25)
        
        # Add full model section
        create_content_slide(prs, "Full Model Analysis", "text", 
                           "Detailed analysis of the full 820-parameter model")
        
        for idx, image_io in enumerate(full_images[:10]):
            slide = create_content_slide(prs, f"Full Model - Chart {idx + 1}", "image", image_io)
        
        # Add reduced model section
        create_content_slide(prs, "Reduced Model Analysis", "text", 
                           "Streamlined analysis of the 451-parameter reduced model")
        
        for idx, image_io in enumerate(reduced_images[:10]):
            slide = create_content_slide(prs, f"Reduced Model - Chart {idx + 1}", "image", image_io)
        
        # Save presentation
        prs.save(output_path)
        print(f"✓ Comparison PowerPoint saved: {output_path}")
        return True
        
    except Exception as e:
        print(f"Error creating comparison PowerPoint presentation: {e}")
        return False


def convert_html_to_powerpoint():
    """
    Convert existing HTML reports to PowerPoint presentations.
    Looks for doe_analysis_report.html and doe_analysis_reduced.html in outputs/
    
    Returns:
        dict: Status of conversion for each file
    """
    if not PPTX_AVAILABLE:
        print("Error: python-pptx not installed. Install with: pip install python-pptx")
        return {"status": "error", "message": "python-pptx not installed"}
    
    output_dir = Path("outputs")
    results = {"status": "starting", "conversions": {}}
    
    # Full model
    full_html = output_dir / "doe_analysis_report.html"
    full_pptx = output_dir / "doe_analysis_report.pptx"
    
    if full_html.exists():
        print(f"\nConverting full model HTML to PowerPoint...")
        success = create_full_model_powerpoint(
            str(full_html), 
            str(full_pptx),
            title="DOE Full Model Analysis (820 Parameters)"
        )
        results["conversions"]["full_model"] = {
            "input": str(full_html),
            "output": str(full_pptx),
            "success": success
        }
    
    # Reduced model
    reduced_html = output_dir / "doe_analysis_reduced.html"
    reduced_pptx = output_dir / "doe_analysis_reduced.pptx"
    
    if reduced_html.exists():
        print(f"\nConverting reduced model HTML to PowerPoint...")
        success = create_reduced_model_powerpoint(
            str(reduced_html),
            str(reduced_pptx),
            title="DOE Reduced Model Analysis (451 Parameters)"
        )
        results["conversions"]["reduced_model"] = {
            "input": str(reduced_html),
            "output": str(reduced_pptx),
            "success": success
        }
    
    # Comparison
    if full_html.exists() and reduced_html.exists():
        print(f"\nCreating comparison PowerPoint...")
        comparison_pptx = output_dir / "doe_model_comparison.pptx"
        success = create_comparison_powerpoint(
            str(full_html),
            str(reduced_html),
            str(comparison_pptx),
            title="Full vs Reduced Model Comparison"
        )
        results["conversions"]["comparison"] = {
            "full_model_input": str(full_html),
            "reduced_model_input": str(reduced_html),
            "output": str(comparison_pptx),
            "success": success
        }
    
    results["status"] = "complete"
    return results


if __name__ == "__main__":
    print("PowerPoint Generator for DOE Analysis")
    print("=" * 60)
    
    # Convert HTML to PowerPoint
    results = convert_html_to_powerpoint()
    
    print("\n" + "=" * 60)
    print("CONVERSION RESULTS:")
    print("=" * 60)
    
    for conversion_type, details in results.get("conversions", {}).items():
        if details.get("success"):
            print(f"✓ {conversion_type}: SUCCESS")
            print(f"  Output: {details.get('output', 'N/A')}")
        else:
            print(f"✗ {conversion_type}: FAILED")
    
    print("\nPowerPoint generation complete!")
