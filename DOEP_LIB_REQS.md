# DOE Pipeline Library Requirements

**Complete package list for reproducibility**

---

## Quick Install

```bash
# From requirements format (classic)
pip install -r requirements.txt

# Or install by category (see below)
```

---

## 📊 Core Data Science & Statistical Analysis

Core packages for data manipulation, numerical computing, and statistical modeling.

| Package | Version | Purpose |
|---------|---------|---------|
| **pandas** | 1.3.5 | Data loading, manipulation, CSV I/O |
| **numpy** | 1.21.6 | Numerical arrays, mathematical operations |
| **scipy** | 1.7.3 | Statistical tests, optimization, distributions |
| **statsmodels** | 0.13.5 | DOE modeling, `ols()` regression, `C()` categorical |
| **scikit-learn** | 1.0.2 | Machine learning utilities, preprocessing |

**Install:**
```bash
pip install pandas==1.3.5 numpy==1.21.6 scipy==1.7.3 statsmodels==0.13.5 scikit-learn==1.0.2
```

---

## 📈 Visualization & Plotting

Interactive and static plotting libraries for data visualization.

| Package | Version | Purpose |
|---------|---------|---------|
| **plotly** | 5.4.0 | Interactive HTML plots (hover, zoom, export) |
| **matplotlib** | 3.5.2 | Static plotting foundation |
| **kaleido** | 0.2.1 | PNG/SVG export from Plotly |

**Install:**
```bash
pip install plotly==5.4.0 matplotlib==3.5.2 kaleido==0.2.1
```

---

## 📄 HTML/XML Processing

Libraries for parsing and manipulating HTML and XML content.

| Package | Version | Purpose |
|---------|---------|---------|
| **beautifulsoup4** | 4.10.0 | HTML parsing and manipulation |
| **lxml** | 4.9.1 | XML/HTML processing with optimized performance |

**Install:**
```bash
pip install beautifulsoup4==4.10.0 lxml==4.9.1
```

---

## 📋 PDF Generation & Image Processing

Libraries for creating PDFs and handling image processing.

| Package | Version | Purpose |
|---------|---------|---------|
| **reportlab** | 3.6.6 | PDF generation (tables, images, layouts) |
| **pillow** | 9.1.1 | Image processing (PIL) |
| **pdfkit** | 1.0.0 | Advanced PDF features (optional, requires wkhtmltopdf) |

**Install:**
```bash
pip install reportlab==3.6.6 pillow==9.1.1 pdfkit==1.0.0
```

**Note:** `pdfkit` requires system binary `wkhtmltopdf`:
```bash
# macOS
brew install --cask wkhtmltopdf

# Ubuntu/Debian
sudo apt-get install wkhtmltopdf

# Windows
choco install wkhtmltopdf
```

---

## 🎯 PowerPoint Generation

Library for creating and modifying PowerPoint presentations.

| Package | Version | Purpose |
|---------|---------|---------|
| **python-pptx** | 0.6.21 | Create/modify PowerPoint presentations (.pptx) |

**Install:**
```bash
pip install python-pptx==0.6.21
```

---

## 📦 Complete Requirements List

For `pip install -r requirements.txt`:

```pip-requirements
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

---

## 📊 Summary

| Category | Packages | Count |
|----------|----------|-------|
| Data Science | pandas, numpy, scipy, statsmodels, scikit-learn | 5 |
| Visualization | plotly, matplotlib, kaleido | 3 |
| HTML/XML | beautifulsoup4, lxml | 2 |
| PDF | reportlab, pillow, pdfkit | 3 |
| PowerPoint | python-pptx | 1 |
| **Total** | | **13** |

| Metric | Value |
|--------|-------|
| Total Packages | 13 |
| Installation Time | 2-5 minutes |
| Disk Space | ~300-500 MB (including venv) |
| Python Version | 3.8+ |
| Platforms | macOS, Linux, Windows |

---

## ✅ Compatibility

- **Python:** 3.8+
- **Platforms:** macOS 10.15+, Ubuntu 18.04+, Windows 10+
- **Virtual Environment:** Recommended (venv, conda, poetry)

---

## 🔗 Related Documentation

- **[DOEP_SETUP.md](./DOEP_SETUP.md)** - Complete installation and setup guide
- **[DOEP_README.md](./DOEP_README.md)** - Pipeline documentation and reference
- **[DOCSTRING_PLAN.md](./DOCSTRING_PLAN.md)** - Architecture and planning documentation

---

## 📝 Version History

| Date | Version | Changes |
|------|---------|---------|
| 2024 | 1.0 | Initial requirements list for DOE Pipeline v2.0 |

---

**File Version:** 1.0  
**Last Updated:** 2024  
**Maintained By:** DOE Pipeline Development Team
