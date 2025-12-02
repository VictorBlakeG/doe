# PowerPoint Generator for DOE Analysis

## Overview

The `powerpoint_generator.py` module creates professional PowerPoint presentations (.pptx files) from your DOE analysis HTML reports. This provides an alternative format for sharing and presenting your statistical analysis results.

## Features

✅ **Professional Formatting**
- Custom slide layouts with branded colors (dark blue theme)
- Proper spacing and typography
- Consistent formatting throughout

✅ **Multiple Presentation Types**
- Full model analysis presentation
- Reduced model analysis presentation
- Side-by-side comparison presentation

✅ **Rich Content Integration**
- Embedded images from HTML reports (50+ charts per presentation)
- Formatted data tables with alternating row colors
- Text summary slides with key findings
- Leverage plots and diagnostic visualizations

✅ **Automatic Generation**
- Seamlessly integrated into the pipeline
- Automatic extraction from HTML sources
- Saves directly to outputs/ directory

## Generated Files

### 1. `doe_analysis_report.pptx` (1.5 MB)
- Full model analysis with 820 parameters
- Title slide and overview
- ~20-25 chart slides with leverage plots
- 2-3 data table slides
- Summary slide with key metrics

### 2. `doe_analysis_reduced.pptx` (1.5 MB)
- Reduced model analysis with 451 parameters
- Title slide and reduced model focus
- ~20-25 chart slides with leverage plots
- 2-3 data table slides
- Model reduction benefits summary

### 3. `doe_model_comparison.pptx` (1.6 MB)
- Side-by-side full vs reduced model comparison
- Metrics comparison slide
- Full model analysis section (10+ charts)
- Reduced model analysis section (10+ charts)
- Statistical validation

## Module Functions

### Core Functions

```python
create_title_slide(prs, title, subtitle="")
```
Creates a branded title slide with custom formatting.

```python
create_content_slide(prs, title, content_type="text", content=None)
```
Creates content slides supporting text, tables, and images.
- `content_type`: "text", "table", or "image"
- `content`: String, DataFrame, or image path/BytesIO

```python
extract_base64_images_from_html(html_path, max_images=10)
```
Extracts base64-encoded images from HTML files.
- Returns: List of BytesIO image objects

```python
extract_html_tables(html_path)
```
Extracts all tables from HTML using pandas.
- Returns: List of DataFrames

### Presentation Generation

```python
create_full_model_powerpoint(html_path, output_path, title="...")
```
Creates a PowerPoint from full model HTML report.

```python
create_reduced_model_powerpoint(html_path, output_path, title="...")
```
Creates a PowerPoint from reduced model HTML report.

```python
create_comparison_powerpoint(full_html, reduced_html, output_path, title="...")
```
Creates a comparison presentation of both models.

```python
convert_html_to_powerpoint()
```
Main function that automatically converts all available HTML reports in outputs/ directory.

## Usage

### Automatic (Recommended)
The pipeline now automatically generates PowerPoints:

```bash
python doep.py
```

### Manual Generation
```python
from powerpoint_generator import convert_html_to_powerpoint

# Convert all HTML reports to PowerPoint
results = convert_html_to_powerpoint()
```

### Custom Presentations
```python
from powerpoint_generator import create_full_model_powerpoint

create_full_model_powerpoint(
    'outputs/doe_analysis_report.html',
    'custom_output.pptx',
    title='My Custom DOE Analysis'
)
```

## Slide Anatomy

### Title Slide
- Dark blue background (RGB: 31, 78, 121)
- Large white main title (54pt bold)
- Smaller subtitle (28pt light gray)

### Content Slides
- Branded title bar with blue theme
- Horizontal separator line
- Content area with consistent margins
- Supports multiple content types

### Table Formatting
- Blue header row with white text
- Alternating gray/white data rows
- Auto-formatted numeric values (scientific notation for very small numbers)
- Column widths optimized for content
- Limited to 12 rows per slide for readability

### Image Slides
- Images centered and scaled appropriately
- Maintains aspect ratio
- Optimized for presentation viewing

## Technical Details

### Dependencies
- `python-pptx`: PowerPoint generation
- `pandas`: Table extraction and manipulation
- `beautifulsoup4`: HTML parsing
- `plotly` & `matplotlib`: Chart generation (for optional custom charts)

### File Format
- **Format**: Office Open XML (.pptx)
- **Compatibility**: Microsoft PowerPoint, Google Slides, LibreOffice
- **File Size**: ~1.5-1.6 MB per presentation
- **Slides**: 20-30 slides per presentation

### Image Extraction
- Extracts base64-encoded images from HTML
- Supports PNG, JPG, GIF formats
- Automatically handles encoding/decoding
- Limited to 50 images per report

### Table Extraction
- Uses pandas `read_html()` for robust extraction
- Handles complex HTML tables
- Converts to professional formatted slides
- Displays first 10 rows per table slide

## Customization

### Modify Slide Styling
Edit colors and fonts in the functions:
```python
fill.fore_color.rgb = RGBColor(31, 78, 121)  # Change blue theme
p.font.size = Pt(54)  # Adjust font sizes
```

### Control Content Limits
Modify max image/table limits in conversion functions:
```python
images = extract_base64_images_from_html(html_path, max_images=100)  # More images
```

### Add Custom Slides
```python
from powerpoint_generator import create_content_slide
from pptx import Presentation

prs = Presentation()
slide = create_content_slide(prs, "My Title", "text", "My content")
prs.save('custom.pptx')
```

## Integration with Pipeline

The PowerPoint generator is fully integrated into the analysis pipeline:

```
Data Import
    ↓
Data Cleaning
    ↓
DOE Design Setup
    ↓
Model Fitting
    ↓
HTML Report Generation
    ↓
PDF Conversion ← (existing)
    ↓
PowerPoint Conversion ← (NEW - Step 17)
    ↓
Complete
```

## Output Location

All generated PowerPoint files are saved to:
```
/Users/vblake/doe2/outputs/
├── doe_analysis_report.pptx          (Full model)
├── doe_analysis_reduced.pptx         (Reduced model)
└── doe_model_comparison.pptx         (Comparison)
```

## Example Output Structure

### doe_analysis_report.pptx (Full Model)
1. Title Slide: "DOE Full Model Analysis"
2. 20-25 Chart Slides: Leverage plots, fit diagrams, diagnostics
3. 2-3 Table Slides: Parameter estimates, ANOVA results
4. Summary Slide: Key findings and metrics

### doe_model_comparison.pptx
1. Title Slide: "Full vs Reduced Model Comparison"
2. Metrics Comparison: Side-by-side statistics
3. Full Model Section: 10 chart slides
4. Reduced Model Section: 10 chart slides

## Troubleshooting

### "python-pptx not installed"
```bash
pip install python-pptx beautifulsoup4
```

### Images not appearing in PowerPoint
- Ensure HTML files contain base64-encoded images
- Check that HTML files exist in outputs/ directory
- Verify image format compatibility

### Tables not formatting correctly
- Ensure HTML contains proper `<table>` elements
- Check DataFrame content before table creation
- Verify pandas `read_html()` successfully extracts tables

### File size issues
- Large files (1.5+ MB) are normal with embedded images
- Consider limiting number of images if needed
- Use compression tools if file size is critical

## Performance

- Typical generation time: 10-15 seconds per presentation
- Fast image extraction from HTML
- Efficient table processing
- Minimal memory overhead

## Future Enhancements

Potential improvements for future versions:
- Custom color themes and branding
- Animated transitions between slides
- Speaker notes with statistical interpretations
- Interactive charts (PowerPoint 365 support)
- Custom chart rendering for better quality
- Automatic slide organization by content type
- Export to PDF from PowerPoint for distribution

## Integration Example

```python
# doep.py workflow includes PowerPoint generation
from powerpoint_generator import convert_html_to_powerpoint

# After fitting models and generating HTML...
pptx_results = convert_html_to_powerpoint()

# Check results
for conversion_type, details in pptx_results['conversions'].items():
    if details['success']:
        print(f"✓ Generated: {details['output']}")
```

## Support

For issues or customization requests, refer to:
- `powerpoint_generator.py` source code
- python-pptx documentation: https://python-pptx.readthedocs.io/
- BeautifulSoup documentation: https://www.crummy.com/software/BeautifulSoup/

---

**Version**: 1.0
**Last Updated**: December 1, 2025
**Status**: Production Ready ✅
