# PowerPoint Generator - Code Examples & API Reference

## Quick Start

### 1. Automatic Generation (Recommended)
The easiest way - everything happens automatically when you run the pipeline:

```bash
python doep.py
```

This will generate 3 PowerPoint files automatically:
- `outputs/doe_analysis_report.pptx`
- `outputs/doe_analysis_reduced.pptx`
- `outputs/doe_model_comparison.pptx`

### 2. Manual Generation

Generate PowerPoints from existing HTML files:

```python
from powerpoint_generator import convert_html_to_powerpoint

# Convert all HTML reports in outputs/ to PowerPoint
results = convert_html_to_powerpoint()

# Check results
for conversion_type, details in results['conversions'].items():
    print(f"{conversion_type}: {details['success']}")
    if details['success']:
        print(f"  -> {details['output']}")
```

## API Reference

### Main Functions

#### `convert_html_to_powerpoint()`
Converts all HTML reports in `outputs/` directory to PowerPoint presentations.

**Returns:**
```python
{
    'status': 'complete',
    'conversions': {
        'full_model': {
            'input': 'outputs/doe_analysis_report.html',
            'output': 'outputs/doe_analysis_report.pptx',
            'success': True
        },
        'reduced_model': { ... },
        'comparison': { ... }
    }
}
```

**Example:**
```python
from powerpoint_generator import convert_html_to_powerpoint

results = convert_html_to_powerpoint()
if results['status'] == 'complete':
    print("✓ All conversions successful!")
```

---

#### `create_full_model_powerpoint(html_path, output_path, title="")`
Create PowerPoint from full model HTML report.

**Parameters:**
- `html_path` (str): Path to HTML file (e.g., `'outputs/doe_analysis_report.html'`)
- `output_path` (str): Output PowerPoint file path (e.g., `'my_presentation.pptx'`)
- `title` (str): Presentation title (default: `"DOE Full Model Analysis"`)

**Returns:** `bool` - True if successful, False otherwise

**Example:**
```python
from powerpoint_generator import create_full_model_powerpoint

success = create_full_model_powerpoint(
    'outputs/doe_analysis_report.html',
    'full_model_presentation.pptx',
    title='Full DOE Model Analysis Report'
)

if success:
    print("✓ PowerPoint created successfully!")
else:
    print("✗ PowerPoint creation failed!")
```

---

#### `create_reduced_model_powerpoint(html_path, output_path, title="")`
Create PowerPoint from reduced model HTML report.

**Parameters:**
- `html_path` (str): Path to HTML file (e.g., `'outputs/doe_analysis_reduced.html'`)
- `output_path` (str): Output PowerPoint file path (e.g., `'my_presentation.pptx'`)
- `title` (str): Presentation title (default: `"DOE Reduced Model Analysis"`)

**Returns:** `bool` - True if successful, False otherwise

**Example:**
```python
from powerpoint_generator import create_reduced_model_powerpoint

success = create_reduced_model_powerpoint(
    'outputs/doe_analysis_reduced.html',
    'reduced_model_presentation.pptx',
    title='Reduced DOE Model Analysis'
)
```

---

#### `create_comparison_powerpoint(full_html, reduced_html, output_path, title="")`
Create side-by-side comparison presentation.

**Parameters:**
- `full_html` (str): Path to full model HTML
- `reduced_html` (str): Path to reduced model HTML
- `output_path` (str): Output PowerPoint file path
- `title` (str): Presentation title (default: `"Full vs Reduced Model Comparison"`)

**Returns:** `bool` - True if successful, False otherwise

**Example:**
```python
from powerpoint_generator import create_comparison_powerpoint

success = create_comparison_powerpoint(
    'outputs/doe_analysis_report.html',
    'outputs/doe_analysis_reduced.html',
    'comparison.pptx',
    title='Model Comparison: Full vs Reduced'
)
```

---

### Utility Functions

#### `extract_base64_images_from_html(html_path, max_images=10)`
Extract base64-encoded images from HTML.

**Parameters:**
- `html_path` (str): Path to HTML file
- `max_images` (int): Maximum number of images to extract (default: 10)

**Returns:** `list` - List of BytesIO image objects

**Example:**
```python
from powerpoint_generator import extract_base64_images_from_html

images = extract_base64_images_from_html(
    'outputs/doe_analysis_report.html',
    max_images=50
)

print(f"Extracted {len(images)} images")

# Use images in custom presentation
for idx, image in enumerate(images):
    print(f"Image {idx + 1}: {len(image.getvalue())} bytes")
```

---

#### `extract_html_tables(html_path)`
Extract all tables from HTML as pandas DataFrames.

**Parameters:**
- `html_path` (str): Path to HTML file

**Returns:** `list` - List of pandas DataFrames

**Example:**
```python
from powerpoint_generator import extract_html_tables
import pandas as pd

tables = extract_html_tables('outputs/doe_analysis_report.html')

print(f"Found {len(tables)} tables")
for idx, df in enumerate(tables):
    print(f"\nTable {idx + 1}:")
    print(f"  Shape: {df.shape}")
    print(f"  Columns: {list(df.columns)}")
    print(df.head())
```

---

#### `create_title_slide(prs, title, subtitle="")`
Add a title slide to presentation.

**Parameters:**
- `prs` (Presentation): python-pptx Presentation object
- `title` (str): Main title text
- `subtitle` (str): Subtitle text (optional)

**Returns:** `Slide` - The created slide

**Example:**
```python
from pptx import Presentation
from powerpoint_generator import create_title_slide

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

slide = create_title_slide(
    prs,
    title='My DOE Analysis',
    subtitle='Statistical Design of Experiments'
)

prs.save('custom_presentation.pptx')
```

---

#### `create_content_slide(prs, title, content_type="text", content=None)`
Add a content slide with title and content.

**Parameters:**
- `prs` (Presentation): python-pptx Presentation object
- `title` (str): Slide title
- `content_type` (str): Type of content - `"text"`, `"table"`, or `"image"`
- `content`: Content to add (string for text, DataFrame for table, path/BytesIO for image)

**Returns:** `Slide` - The created slide

**Example - Text Slide:**
```python
from pptx import Presentation
from powerpoint_generator import create_content_slide

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

text_content = """
Model Summary:
• 820 parameters in full model
• R² = 0.3897
• MSE = 1.7437
• Significance: p < 0.001
"""

slide = create_content_slide(
    prs,
    title='Model Summary',
    content_type='text',
    content=text_content
)

prs.save('summary.pptx')
```

**Example - Table Slide:**
```python
import pandas as pd
from pptx import Presentation
from powerpoint_generator import create_content_slide

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Create sample data
data = {
    'Parameter': ['Intercept', 'Factor_A', 'Factor_B'],
    'Coefficient': [45.23, 12.56, -8.34],
    'Std Error': [2.1, 1.8, 2.3],
    'P-Value': [0.0001, 0.0015, 0.0234]
}
df = pd.DataFrame(data)

slide = create_content_slide(
    prs,
    title='Parameter Estimates',
    content_type='table',
    content=df
)

prs.save('table_slide.pptx')
```

**Example - Image Slide:**
```python
from pptx import Presentation
from powerpoint_generator import create_content_slide

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

slide = create_content_slide(
    prs,
    title='Model Fit Diagram',
    content_type='image',
    content='path/to/image.png'
)

prs.save('image_slide.pptx')
```

---

#### `add_image_to_slide(slide, image_source, left=Inches(1), top=Inches(1.4), width=Inches(8))`
Add an image to an existing slide.

**Parameters:**
- `slide` (Slide): Slide object to add image to
- `image_source`: Image file path or BytesIO object
- `left` (Length): Left position (default: 1 inch)
- `top` (Length): Top position (default: 1.4 inches)
- `width` (Length): Image width (default: 8 inches)

**Example:**
```python
from pptx import Presentation
from pptx.util import Inches
from powerpoint_generator import create_content_slide, add_image_to_slide

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

slide = create_content_slide(prs, 'My Slide', 'text', 'Some text content')

# Add an image
add_image_to_slide(
    slide,
    'outputs/chart.png',
    left=Inches(0.5),
    top=Inches(2),
    width=Inches(9)
)

prs.save('slide_with_image.pptx')
```

---

## Advanced Examples

### Custom Presentation from Scratch

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from powerpoint_generator import (
    create_title_slide, 
    create_content_slide,
    extract_base64_images_from_html,
    extract_html_tables
)

# Create presentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Add title slide
create_title_slide(
    prs,
    'My Custom DOE Analysis',
    'Advanced Statistical Report'
)

# Extract content from HTML
html_path = 'outputs/doe_analysis_report.html'
images = extract_base64_images_from_html(html_path, max_images=30)
tables = extract_html_tables(html_path)

# Add overview slide
overview_text = """
Analysis Overview:
• Full factorial design with 7600 observations
• 820 parameters in full model
• 451 parameters in reduced model
• Model reduction: 45% parameter elimination
• R² maintained at 99% of full model value
"""
create_content_slide(prs, 'Overview', 'text', overview_text)

# Add image slides
for idx, image in enumerate(images[:15]):
    create_content_slide(
        prs,
        f'Diagnostic Chart {idx + 1}',
        'image',
        image
    )

# Add table slides
for idx, df in enumerate(tables[:3]):
    create_content_slide(
        prs,
        f'Results Table {idx + 1}',
        'table',
        df.head(10)
    )

# Add conclusion slide
conclusion_text = """
Key Findings:
✓ Reduced model provides equivalent predictive performance
✓ Simplified model improves interpretability
✓ 45% fewer parameters with minimal accuracy loss
✓ Validated through LOF testing and cross-validation
"""
create_content_slide(prs, 'Conclusions', 'text', conclusion_text)

# Save presentation
prs.save('custom_doe_analysis.pptx')
print("✓ Custom presentation created!")
```

### Batch Processing Multiple Reports

```python
from pathlib import Path
from powerpoint_generator import create_full_model_powerpoint
import os

# Process all HTML files in directory
html_dir = Path('outputs')
for html_file in html_dir.glob('*.html'):
    if 'analysis' in html_file.name:
        # Create output filename
        output_file = html_file.with_suffix('.pptx')
        
        # Get title from filename
        title = html_file.stem.replace('_', ' ').title()
        
        print(f"Converting {html_file.name}...")
        success = create_full_model_powerpoint(
            str(html_file),
            str(output_file),
            title=title
        )
        
        if success:
            file_size = os.path.getsize(output_file)
            print(f"  ✓ Created {output_file.name} ({file_size / 1024 / 1024:.1f} MB)")
        else:
            print(f"  ✗ Failed to create {output_file.name}")
```

### Integration with Data Processing

```python
from pathlib import Path
from doe import fit_doe_model, fit_reduced_doe_model
from powerpoint_generator import convert_html_to_powerpoint

# Process data and fit models
print("Fitting DOE models...")
model, results, stats = fit_doe_model(doe_df)
reduced_model, reduced_results, reduced_stats = fit_reduced_doe_model(doe_df, results)

# Generate HTML reports (already done in pipeline)
print("HTML reports generated")

# Convert to PowerPoint
print("Converting to PowerPoint...")
pptx_results = convert_html_to_powerpoint()

# Verify results
if pptx_results['status'] == 'complete':
    for conversion_type, details in pptx_results['conversions'].items():
        if details['success']:
            output_path = Path(details['output'])
            size_mb = output_path.stat().st_size / 1024 / 1024
            print(f"✓ {conversion_type}: {output_path.name} ({size_mb:.1f} MB)")
        else:
            print(f"✗ {conversion_type}: Failed")

print("Complete!")
```

## Customization Guide

### Change Color Theme

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

def create_custom_title_slide(prs, title, subtitle="", bg_color=(31, 78, 121)):
    """Create title slide with custom color."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # Set background color
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*bg_color)
    
    # Add title
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(2), Inches(9), Inches(1.5)
    )
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    
    if subtitle:
        subtitle_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(3.8), Inches(9), Inches(1)
        )
        subtitle_frame = subtitle_box.text_frame
        p = subtitle_frame.paragraphs[0]
        p.text = subtitle
        p.font.size = Pt(28)
        p.font.color.rgb = RGBColor(200, 200, 200)
        p.alignment = PP_ALIGN.CENTER
    
    return slide

# Usage
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Use custom color (e.g., dark green)
create_custom_title_slide(prs, 'My Title', 'Subtitle', bg_color=(34, 102, 51))
```

### Modify Slide Dimensions

```python
from pptx import Presentation
from pptx.util import Inches

# Standard 16:9 widescreen (default)
prs = Presentation()  # 10" x 7.5"

# 4:3 standard
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Custom size (e.g., 12" x 9")
prs = Presentation()
prs.slide_width = Inches(12)
prs.slide_height = Inches(9)
```

## Troubleshooting

### Issue: "ImportError: No module named 'pptx'"

**Solution:**
```bash
pip install python-pptx
```

### Issue: No images appearing in PowerPoint

**Cause:** HTML file doesn't contain base64-encoded images or they're in unsupported format.

**Solution:** Check HTML file for `<img>` tags with `data:image/` src attributes.

```python
from bs4 import BeautifulSoup

with open('your_file.html', 'r') as f:
    soup = BeautifulSoup(f, 'html.parser')
    img_count = len(soup.find_all('img'))
    print(f"Found {img_count} images in HTML")
    
    for img in soup.find_all('img')[:5]:
        src = img.get('src', '')[:50]
        print(f"  - {src}...")
```

### Issue: Tables not formatting correctly

**Solution:** Check that tables have proper structure:

```python
import pandas as pd

# Verify tables can be extracted
tables = pd.read_html('your_file.html')
print(f"Extracted {len(tables)} tables")

for idx, df in enumerate(tables):
    print(f"\nTable {idx + 1}:")
    print(f"  Shape: {df.shape}")
    print(f"  Dtypes: {df.dtypes.to_dict()}")
```

---

**Last Updated:** December 1, 2025
**Status:** Production Ready ✅
