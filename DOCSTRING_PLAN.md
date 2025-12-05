# Docstring Implementation & Documentation Plan

## Executive Summary

This plan outlines a comprehensive approach to:
1. **Add docstrings** (using `"""` notation) to **48 functions** across **13 Python modules**
2. **Create `doep_readme.md`** - A comprehensive pipeline overview document
3. **Maintain consistency** with PEP 257 docstring standards

---

## Part 1: Current State Assessment

### Docstring Coverage Summary

| File | Functions | With Docstrings | Coverage |
|------|-----------|-----------------|----------|
| balance.py | 1 | 0 | 0% |
| clean.py | 2 | 1 | 50% |
| doe.py | 12 | 6 | 50% |
| doep.py | 1 | 0 | 0% |
| export.py | 1 | 0 | 0% |
| importcsv.py | 3 | 1 | 33% |
| pdf_generator.py | 3 | 1 | 33% |
| pdf_generator_enhanced.py | 4 | 2 | 50% |
| pdf_generator_plotly.py | 7 | 3 | 43% |
| powerpoint_generator.py | 14 | 6 | 43% |
| prep.py | 2 | 1 | 50% |
| split.py | 1 | 0 | 0% |
| viz.py | 2 | 1 | 50% |
| **TOTALS** | **48** | **22** | **46%** |

### Files Requiring Most Work
1. **balance.py** - 1 function (0 docstrings)
2. **doep.py** - 1 function (0 docstrings)
3. **export.py** - 1 function (0 docstrings)
4. **split.py** - 1 function (0 docstrings)
5. **importcsv.py** - 3 functions (1 docstring)
6. **pdf_generator.py** - 3 functions (1 docstring)
7. **powerpoint_generator.py** - 14 functions (6 docstrings)
8. **pdf_generator_plotly.py** - 7 functions (3 docstrings)
9. **doe.py** - 12 functions (6 docstrings)

---

## Part 2: Docstring Standards

All docstrings will follow PEP 257 and use the following format:

### Format Template

```python
def function_name(param1, param2, param3=None):
    """
    Brief one-line description of what the function does.
    
    Longer description if needed. This can span multiple lines and provide
    context about the function's purpose, behavior, and use cases.
    
    Args:
        param1 (type): Description of param1
        param2 (type): Description of param2
        param3 (type, optional): Description of param3. Defaults to None.
        
    Returns:
        return_type: Description of return value
        
    Raises:
        ExceptionType: Description of when this exception is raised
        
    Example:
        >>> result = function_name(value1, value2)
        >>> print(result)
        expected_output
    """
```

### Docstring Components (in order)
1. **Brief summary** (1 line)
2. **Extended description** (if complex logic)
3. **Args** (parameters with types and descriptions)
4. **Returns** (return type and description)
5. **Raises** (exceptions that may be raised)
6. **Example** (usage example, if non-obvious)

---

## Part 3: Module-by-Module Docstring Implementation Plan

### Phase 1: Quick Fixes (Small Modules - No New Docstrings)

#### 1.1 **balance.py** (1 function, 0 docstrings)
- **File**: `/Users/vblake/doe2/balance.py`
- **Functions to document**:
  - `balance_dataframes()` - Balances two dataframes to equal size
- **Effort**: ~5 minutes

#### 1.2 **doep.py** (1 function, 0 docstrings)
- **File**: `/Users/vblake/doe2/doep.py`
- **Functions to document**:
  - `main()` - Already has docstring; verify it's complete ✓
- **Effort**: ~2 minutes (already has docstring)

#### 1.3 **export.py** (1 function, 0 docstrings)
- **File**: `/Users/vblake/doe2/export.py`
- **Functions to document**:
  - `export_fan_dfs_to_csv()` - Exports fan dataframes to CSV files
- **Effort**: ~5 minutes

#### 1.4 **split.py** (1 function, 0 docstrings)
- **File**: `/Users/vblake/doe2/split.py`
- **Functions to document**:
  - `split_fan()` - Splits data by fan speed threshold
- **Effort**: ~5 minutes

### Phase 2: Small Modules (2-3 Functions)

#### 2.1 **clean.py** (2 functions, 1 docstring)
- **File**: `/Users/vblake/doe2/clean.py`
- **Functions to document**:
  - `analyze_device_vendors()` - Analyze vendor distribution
  - `clean_tman()` - Clean transceiver manufacturer data (needs docstring)
- **Effort**: ~10 minutes

#### 2.2 **prep.py** (2 functions, 1 docstring)
- **File**: `/Users/vblake/doe2/prep.py`
- **Functions to document**:
  - `calculate_fan_speed_mean()` - Calculate mean fan speed
  - `create_fan_speed_histogram()` - needs docstring
- **Effort**: ~10 minutes

#### 2.3 **viz.py** (2 functions, 1 docstring)
- **File**: `/Users/vblake/doe2/viz.py`
- **Functions to document**:
  - `create_fan_hl_histogram()` - Create comparative fan speed histogram
  - `create_ttemp_hl_histogram()` - needs docstring
- **Effort**: ~10 minutes

#### 2.4 **importcsv.py** (3 functions, 1 docstring)
- **File**: `/Users/vblake/doe2/importcsv.py`
- **Functions to document**:
  - `load_csv_data()` - Load main CSV file
  - `show_data_summary()` - needs docstring
  - `remove_missing_data()` - needs docstring
- **Effort**: ~15 minutes

### Phase 3: PDF Generator Modules (3-7 Functions)

#### 3.1 **pdf_generator.py** (3 functions, 1 docstring)
- **File**: `/Users/vblake/doe2/pdf_generator.py`
- **Functions to document**:
  - `create_design_summary_pdf()` - Create DOE design PDF
  - `create_analysis_summary_pdf()` - needs docstring
  - `create_reduced_summary_pdf()` - needs docstring
- **Effort**: ~15 minutes

#### 3.2 **pdf_generator_enhanced.py** (4 functions, 2 docstrings)
- **File**: `/Users/vblake/doe2/pdf_generator_enhanced.py`
- **Functions to document**:
  - `create_design_summary_pdf()` - Already documented
  - `create_reduced_summary_pdf_fallback()` - needs docstring
  - `create_analysis_summary_pdf()` - Already documented
  - `create_full_model_summary_pdf()` - needs docstring
- **Effort**: ~15 minutes

#### 3.3 **pdf_generator_plotly.py** (7 functions, 3 docstrings)
- **File**: `/Users/vblake/doe2/pdf_generator_plotly.py`
- **Functions to document**:
  - `extract_plotly_chart_as_image()` - needs docstring
  - `create_model_formula_string()` - Already documented
  - `create_parameters_table()` - needs docstring
  - `create_reduced_model_pdf_enhanced()` - Already documented
  - `create_full_model_pdf_enhanced()` - needs docstring
  - `extract_coefficients_from_html()` - needs docstring
  - `extract_interaction_plots_from_html_for_pdf()` - needs docstring
- **Effort**: ~25 minutes

### Phase 4: Major Modules (12-14 Functions)

#### 4.1 **doe.py** (12 functions, 6 docstrings)
- **File**: `/Users/vblake/doe2/doe.py`
- **Functions to document**:
  - `setup_doe_design()` - Already documented
  - `create_full_factorial_design()` - Already documented
  - `fit_doe_model()` - Already documented
  - `_calculate_lack_of_fit()` - needs docstring
  - `_debug_model_comparison()` - needs docstring
  - `fit_reduced_doe_model()` - Already documented
  - `_clean_label()` - needs docstring
  - `_calculate_p_value()` - needs docstring
  - `_is_categorical()` - needs docstring
  - `create_interaction_plots()` - Already documented
  - `create_doe_report()` - Already documented
  - `create_reduced_doe_report()` - Already documented
- **Effort**: ~30 minutes

#### 4.2 **powerpoint_generator.py** (14 functions, 6 docstrings)
- **File**: `/Users/vblake/doe2/powerpoint_generator.py`
- **Functions to document**:
  - `create_title_slide()` - Already documented
  - `create_content_slide()` - Already documented
  - `add_image_to_slide()` - Already documented
  - `extract_base64_images_from_html()` - needs docstring
  - `extract_html_tables()` - Already documented
  - `create_equation_slide()` - Already documented
  - `add_side_by_side_leverage_comparisons()` - needs docstring
  - `create_full_model_powerpoint()` - Already documented
  - `create_reduced_model_powerpoint()` - Already documented
  - `create_comparison_powerpoint()` - Already documented
  - `convert_html_to_powerpoint()` - needs docstring
  - `add_image_background()` - needs docstring
  - `extract_model_fit_plot_from_pdf()` - needs docstring
  - `_extract_model_fit_from_pdf_pdfplumber()` - needs docstring
- **Effort**: ~35 minutes

---

## Part 4: `doep_readme.md` Structure

### File: `doep_readme.md`

#### Section 1: Pipeline Overview
- High-level flow diagram (ASCII)
- Quick start guide
- 16 numbered pipeline steps with descriptions

#### Section 2: Pipeline Steps & Function Mapping

For each of the 16 steps in `doep.py`, document:
- Step number and name
- Input/output data
- Primary function calls (nested under step)
- Brief description of what happens

**Format Example:**

```markdown
## Step 1: Load Fan Speed Dataframes from CSV
**Input:** `outputs/fan_low_df.csv`, `outputs/fan_high_df.csv`  
**Output:** `fan_low_df`, `fan_high_df` (pandas DataFrames)

### Functions Called:
- `pd.read_csv()` (pandas built-in)
  - Loads CSV files from outputs directory
  - Returns dataframe with 12,469 (low) and 3,848 (high) rows

**Description:** Loads pre-split fan speed data directly from CSV files, 
bypassing the original load/clean/split pipeline...
```

#### Section 3: Module Reference
- List all 13 modules
- For each module, show:
  - Module description
  - All functions with one-line summaries
  - Which pipeline steps use functions from this module

#### Section 4: Data Flow Diagram
- ASCII art showing data transformations through pipeline
- Input → Transformation → Output for each step

#### Section 5: Key Statistics
- Total functions: 48
- Total docstrings (before): 22
- Total docstrings (after): 70 (all)
- Lines of code per module
- Interdependencies

---

## Part 5: Quality Checklist

### For Each Docstring:
- [ ] Follows PEP 257 format
- [ ] One-liner summary is complete sentence
- [ ] All parameters documented with types
- [ ] Return type and value documented
- [ ] Exceptions documented (if applicable)
- [ ] Example provided (for complex functions)
- [ ] Uses """ """ notation (not #)
- [ ] Proper indentation (4 spaces)

### For `doep_readme.md`:
- [ ] Lists all 16 pipeline steps
- [ ] Each step has input/output documented
- [ ] Functions are nested under steps
- [ ] Descriptions are clear and concise
- [ ] ASCII diagrams render correctly
- [ ] Links to specific functions work
- [ ] Code examples are correct
- [ ] Module reference is complete
- [ ] TOC links work (if using markdown)

---

## Part 6: Built-in Tests

The repository includes embedded test code that validates core functionality. These are not formal unit tests but rather verification routines built into module entry points.

### Test 1: PDF Generator Enhanced Test
**File**: `pdf_generator_plotly.py`  
**Location**: Lines ~397-400 (if __name__ == '__main__')  
**Functions Executed**:
- `create_reduced_model_pdf_enhanced(pdf_path, html_file_path)` - Main test function

**Test Data**:
- **Input HTML**: `/Users/vblake/doe2/outputs/doe_analysis_reduced.html`
- **Output PDF**: `/tmp/test_enhanced.pdf` (temporary location)

**What It Tests**:
- Verifies HTML report can be parsed and converted to enhanced PDF format
- Tests Plotly chart extraction from HTML
- Validates PDF generation with proper styling and layouts
- Confirms reduced model analysis PDF renders correctly
- Checks image embedding and table formatting

**How to Run**:
```bash
cd /Users/vblake/doe2
python pdf_generator_plotly.py
```

**Expected Output**: PDF file created at `/tmp/test_enhanced.pdf` with all sections properly formatted

---

### Test 2: PowerPoint Generator Test
**File**: `powerpoint_generator.py`  
**Location**: Lines ~1185-1210 (if __name__ == "__main__")  
**Functions Executed**:
- `convert_html_to_powerpoint()` - Main test function

**Test Data**:
- **Input HTMLs**: 
  - `outputs/doe_analysis_report.html` (full model)
  - `outputs/doe_analysis_reduced.html` (reduced model)
- **Output PPTXs**:
  - `outputs/doe_analysis_report.pptx`
  - `outputs/doe_analysis_reduced.pptx`
  - `outputs/doe_model_comparison.pptx`

**What It Tests**:
- Verifies HTML reports can be converted to PowerPoint presentations
- Tests image extraction from HTML (leverage plots, interaction plots)
- Validates table extraction and formatting in slides
- Confirms slide creation with proper titles and content
- Tests model fit plot embedding from PDF
- Validates comparison presentation generation with metrics

**How to Run**:
```bash
cd /Users/vblake/doe2
python powerpoint_generator.py
```

**Expected Output**: Three PowerPoint files created with all slides properly formatted

---

### Test 3: Main Pipeline Test
**File**: `doep.py`  
**Location**: Lines ~125-130 (if __name__ == "__main__")  
**Functions Executed**:
- `main()` - Complete pipeline execution

**Test Data**:
- **Inputs**: 
  - `outputs/fan_low_df.csv` (12,469 rows)
  - `outputs/fan_high_df.csv` (3,848 rows)
- **Intermediate Processing**: Data cleaning, merging, analysis
- **Outputs**:
  - HTML reports (3 files)
  - PDF summaries (3 files)
  - PowerPoint presentations (3 files)

**What It Tests**:
- Validates complete end-to-end data pipeline
- Tests data loading and cleaning steps
- Verifies DOE model fitting (full and reduced)
- Confirms all output generation (HTML, PDF, PPTX)
- Validates data statistics and summaries
- Tests comparative analysis between full and reduced models

**How to Run**:
```bash
cd /Users/vblake/doe2
python doep.py
```

**Expected Output**: All 9 output files generated with complete analysis and visualizations

---

## Part 7: Repository Statistics

### Overall Project Metrics

| Metric | Value |
|--------|-------|
| Total Python Files (Code) | 13 |
| Total Lines of Code | 6,189 |
| Total Significant Lines* | 5,467 |
| Average Lines per Module | 476 |
| Total Functions/Methods | 48 |
| Functions with Docstrings | 22 (46%) |
| Functions Needing Docstrings | 26 (54%) |
| Estimated Repository Size | ~85 KB |

*Significant Lines = excluding blank lines and comments

---

### Detailed Module Breakdown

| Module | Purpose | Total Lines | Sig. Lines | Functions | Docstrings | Effort |
|--------|---------|------------|------------|-----------|------------|--------|
| **balance.py** | Data balancing | 52 | 45 | 1 | 0 | 5 min |
| **clean.py** | Data cleaning | 78 | 66 | 2 | 1 | 10 min |
| **doe.py** | DOE analysis | 1,628 | 1,450 | 12 | 6 | 30 min |
| **doep.py** | Main pipeline | 130 | 107 | 1 | 1* | 2 min |
| **export.py** | CSV export | 34 | 29 | 1 | 0 | 5 min |
| **importcsv.py** | CSV import | 75 | 68 | 3 | 1 | 15 min |
| **pdf_generator.py** | PDF creation | 350 | 326 | 3 | 1 | 15 min |
| **pdf_generator_enhanced.py** | Enhanced PDF | 577 | 534 | 4 | 2 | 15 min |
| **pdf_generator_plotly.py** | Plotly PDF | 829 | 744 | 7 | 3 | 25 min |
| **powerpoint_generator.py** | PowerPoint | 1,232 | 1,078 | 14 | 6 | 35 min |
| **prep.py** | Data prep | 113 | 100 | 2 | 1 | 10 min |
| **split.py** | Data split | 24 | 21 | 1 | 0 | 5 min |
| **viz.py** | Visualization | 287 | 259 | 2 | 1 | 10 min |
| **TOTALS** | | **6,189** | **5,467** | **48** | **22** | **~5 hrs** |

*doep.py already has docstring for main()

---

### Module Categories

**Data Processing Layer** (6 modules)
- balance.py, clean.py, export.py, importcsv.py, prep.py, split.py
- Combined: 387 total lines, 336 significant
- Focus: Data loading, cleaning, transformation

**Analysis Layer** (1 module)
- doe.py
- 1,628 total lines, 1,450 significant
- Focus: Statistical DOE modeling and analysis

**Output Generation Layer** (4 modules)
- pdf_generator.py, pdf_generator_enhanced.py, pdf_generator_plotly.py, powerpoint_generator.py
- Combined: 3,988 total lines, 3,682 significant
- Focus: Report generation in multiple formats

**Orchestration Layer** (1 module)
- doep.py
- 130 total lines, 107 significant
- Focus: Pipeline coordination

---

### Code Quality Metrics

**Largest Modules** (by significant lines):
1. powerpoint_generator.py - 1,078 sig lines (35% of codebase)
2. doe.py - 1,450 sig lines (26% of codebase)
3. pdf_generator_plotly.py - 744 sig lines (14% of codebase)

**Complexity Distribution**:
- High Complexity (>1000 lines): 2 modules (doe.py, powerpoint_generator.py)
- Medium Complexity (500-1000 lines): 3 modules (pdf_generator_enhanced.py, pdf_generator_plotly.py, prep.py)
- Low Complexity (<500 lines): 8 modules

**Documentation Coverage by Layer**:
- Data Processing: 3/9 functions documented (33%)
- Analysis: 6/12 functions documented (50%)
- Output Generation: 12/25 functions documented (48%)
- Orchestration: 1/1 functions documented (100%)

---

## Part 8: Setup and Installation Guide (`doep_setup.md`)

A comprehensive setup guide will be created with the following structure:

### 8.1 Python Installation

**Section**: System Python Prerequisites
- Verify Python 3.8+ is installed
- Check installation: `python3 --version`
- Verify pip is available: `pip3 --version`
- Platform-specific notes (macOS, Linux, Windows)

**Content**:
```bash
# macOS - using Homebrew
brew install python3

# Ubuntu/Debian
sudo apt-get install python3 python3-pip python3-venv

# Verify installation
python3 --version
pip3 --version
```

---

### 8.2 Virtual Environment Setup

**Section**: Creating and Activating venv
- Create isolated Python environment
- Location: `/Users/vblake/doe2/venv/`
- Activation/deactivation commands

**Content**:
```bash
# Navigate to project directory
cd /Users/vblake/doe2

# Create virtual environment
python3 -m venv venv

# Activate virtual environment
# On macOS/Linux:
source venv/bin/activate

# On Windows:
# venv\Scripts\activate

# Verify activation (should see (venv) prefix)
which python

# Deactivate when done
deactivate
```

**Why**: Isolates project dependencies, prevents conflicts with system packages

---

### 8.3 Code Libraries: DOE & Analysis (Math)

**Section**: Statistical Analysis and Data Science Libraries

**Packages**:
- `pandas==1.3.5` - Data manipulation and analysis
- `numpy==1.21.6` - Numerical computing
- `scipy==1.7.3` - Scientific computing (statistics, optimization)
- `statsmodels==0.13.5` - Statistical modeling (core for DOE)
- `scikit-learn==1.0.2` - Machine learning utilities

**Installation**:
```bash
# Activate venv first
source venv/bin/activate

# Install data science stack
pip install pandas numpy scipy statsmodels scikit-learn

# Verify installation
python3 -c "import pandas; import numpy; import scipy; import statsmodels; import sklearn; print('All math libraries installed!')"
```

**Why These**:
- **pandas**: Load and manipulate CSV data (12K+ rows)
- **numpy**: Numerical arrays and mathematical operations
- **scipy**: Statistical tests (ANOVA, lack-of-fit)
- **statsmodels**: OLS regression modeling, formula API (`ols()`, `C()` for categorical)
- **scikit-learn**: Additional statistical utilities

---

### 8.4 Code Libraries: Visualization (Non-Format Specific)

**Section**: Plotting and Interactive Visualization Libraries

**Packages**:
- `plotly==5.4.0` - Interactive plots (HTML/JS based)
- `matplotlib==3.5.2` - Static plotting (foundation for other tools)
- `kaleido==0.2.1` - Plotly image export to PNG/SVG

**Installation**:
```bash
# Activate venv first
source venv/bin/activate

# Install visualization libraries
pip install plotly matplotlib kaleido

# Verify installation
python3 -c "import plotly; import matplotlib; import kaleido; print('All visualization libraries installed!')"
```

**Why These**:
- **plotly**: Interactive HTML plots with hover/zoom (used for interaction plots, model fit)
- **matplotlib**: Foundation plotting library (used indirectly via other packages)
- **kaleido**: Plotly → PNG/SVG conversion (critical for PDF/PPTX embedding)

---

### 8.5 Code Libraries: Output Formats (HTML, PDF, PPTX)

**Section**: Report Generation and Export Libraries

**Packages**:

#### HTML Output
- `beautifulsoup4==4.10.0` - HTML parsing and manipulation

#### PDF Output
- `reportlab==3.6.6` - PDF generation from Python (core for custom PDF layouts)
- `pillow==9.1.1` - Image processing (PIL - for image embedding in PDFs)
- `pdfkit==1.0.0` - Optional wkhtmltopdf wrapper (advanced PDF features)

#### PowerPoint Output
- `python-pptx==0.6.21` - Create and modify PowerPoint presentations

#### Additional Utilities
- `lxml==4.9.1` - XML processing (HTML parsing enhancement)

**Installation**:
```bash
# Activate venv first
source venv/bin/activate

# Install output format libraries
pip install beautifulsoup4 reportlab pillow python-pptx lxml

# Optional: pdfkit (requires wkhtmltopdf system binary)
pip install pdfkit

# Verify installation
python3 -c "from bs4 import BeautifulSoup; from reportlab import pdfgen; from PIL import Image; from pptx import Presentation; print('All output libraries installed!')"
```

**Why These**:
- **beautifulsoup4**: Parse HTML reports for image/table extraction
- **reportlab**: Generate PDFs with precise layouts, tables, images
- **pillow**: Handle image resizing, conversion, embedding
- **python-pptx**: Create PowerPoint slides programmatically
- **lxml**: Enhanced HTML/XML parsing performance

---

### 8.6 Complete Installation Script

**Section**: One-command setup (for convenience)

```bash
#!/bin/bash
# Complete setup script

cd /Users/vblake/doe2

# Create venv
python3 -m venv venv
source venv/bin/activate

# Install all dependencies in groups
echo "Installing data science libraries..."
pip install pandas numpy scipy statsmodels scikit-learn

echo "Installing visualization libraries..."
pip install plotly matplotlib kaleido

echo "Installing output format libraries..."
pip install beautifulsoup4 reportlab pillow python-pptx lxml

echo "✓ Installation complete!"
pip list
```

**Usage**:
```bash
# Save as setup.sh
chmod +x setup.sh
./setup.sh
```

---

### 8.7 Requirements File

**Section**: `requirements.txt` for reproducibility

```text
# Data Science & Analysis
pandas==1.3.5
numpy==1.21.6
scipy==1.7.3
statsmodels==0.13.5
scikit-learn==1.0.2

# Visualization
plotly==5.4.0
matplotlib==3.5.2
kaleido==0.2.1

# HTML/XML Processing
beautifulsoup4==4.10.0
lxml==4.9.1

# PDF Generation
reportlab==3.6.6
pillow==9.1.1
pdfkit==1.0.0

# PowerPoint Generation
python-pptx==0.6.21
```

**Installation from file**:
```bash
source venv/bin/activate
pip install -r requirements.txt
```

---

### 8.8 Verification Checklist

**Section**: Validate complete installation

```bash
# Run each test
python3 -c "import pandas; print('✓ pandas')"
python3 -c "import numpy; print('✓ numpy')"
python3 -c "import scipy; print('✓ scipy')"
python3 -c "import statsmodels; print('✓ statsmodels')"
python3 -c "import sklearn; print('✓ scikit-learn')"
python3 -c "import plotly; print('✓ plotly')"
python3 -c "import matplotlib; print('✓ matplotlib')"
python3 -c "import kaleido; print('✓ kaleido')"
python3 -c "from bs4 import BeautifulSoup; print('✓ beautifulsoup4')"
python3 -c "from reportlab import pdfgen; print('✓ reportlab')"
python3 -c "from PIL import Image; print('✓ pillow')"
python3 -c "from pptx import Presentation; print('✓ python-pptx')"
python3 -c "import lxml; print('✓ lxml')"

# Or run the complete pipeline
cd /Users/vblake/doe2
python3 doep.py
```

**Expected Output**: All packages imported successfully, or if running pipeline: "Data processing pipeline completed!"

---

### 8.9 Troubleshooting Guide

**Section**: Common issues and solutions

| Issue | Solution |
|-------|----------|
| `venv not found` | Run `python3 -m venv venv` |
| `pip command not found` | Ensure venv is activated: `source venv/bin/activate` |
| `kaleido import fails` | Install additional system libraries (see platform notes) |
| `pdfkit fails` | Install wkhtmltopdf: `brew install wkhtmltopdf` (macOS) |
| `Permission denied on setup.sh` | Run `chmod +x setup.sh` |
| `ModuleNotFoundError` | Ensure venv is activated and run `pip install -r requirements.txt` |

---

## Next Steps

Would you like me to proceed with implementing this plan? I recommend starting with:
1. Approval of this plan
2. Quick wins first (balance.py, export.py, split.py)
3. Then moving to larger modules
4. Finally creating the comprehensive README

**Would you like to modify this plan before I proceed?**

