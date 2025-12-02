# PowerPoint Generator - Quick Reference Card

## 🚀 Quick Start (30 seconds)

```bash
cd /Users/vblake/doe2
python doep.py
```

Your 3 PowerPoint files will be in `outputs/`:
- `doe_analysis_report.pptx` (Full model)
- `doe_analysis_reduced.pptx` (Reduced model)
- `doe_model_comparison.pptx` (Comparison)

## 📌 Common Tasks

### Generate PowerPoints automatically
```bash
python doep.py
```

### Generate PowerPoints manually
```bash
python powerpoint_generator.py
```

### Create custom PowerPoint in Python
```python
from powerpoint_generator import create_full_model_powerpoint

create_full_model_powerpoint(
    'outputs/doe_analysis_report.html',
    'my_presentation.pptx',
    title='My Analysis'
)
```

### Extract images from HTML
```python
from powerpoint_generator import extract_base64_images_from_html

images = extract_base64_images_from_html('output.html')
print(f"Found {len(images)} images")
```

### Extract tables from HTML
```python
from powerpoint_generator import extract_html_tables

tables = extract_html_tables('output.html')
for idx, df in enumerate(tables):
    print(f"Table {idx}: {df.shape}")
```

## 📚 File Locations

```
/Users/vblake/doe2/
├── powerpoint_generator.py          ← Main module (470 lines)
├── POWERPOINT_README.md             ← Feature documentation
├── POWERPOINT_EXAMPLES.md           ← API & code examples
└── outputs/
    ├── doe_analysis_report.pptx     ← Full model (1.5 MB)
    ├── doe_analysis_reduced.pptx    ← Reduced model (1.5 MB)
    └── doe_model_comparison.pptx    ← Comparison (1.6 MB)
```

## 🔧 Functions Quick Reference

**Main Functions:**
- `create_full_model_powerpoint(html_path, output_path, title="")`
- `create_reduced_model_powerpoint(html_path, output_path, title="")`
- `create_comparison_powerpoint(full_html, reduced_html, output_path, title="")`
- `convert_html_to_powerpoint()` - Batch conversion

**Slide Functions:**
- `create_title_slide(prs, title, subtitle="")`
- `create_content_slide(prs, title, content_type, content)`
- `add_image_to_slide(slide, image_source, left, top, width)`

**Extraction Functions:**
- `extract_base64_images_from_html(html_path, max_images=10)`
- `extract_html_tables(html_path)`

## 💡 Tips & Tricks

### Limit images extracted
```python
images = extract_base64_images_from_html('file.html', max_images=20)
```

### Check if installation is working
```python
from powerpoint_generator import PPTX_AVAILABLE
print(f"python-pptx available: {PPTX_AVAILABLE}")
```

### View generated file sizes
```bash
ls -lh /Users/vblake/doe2/outputs/*.pptx
```

### Check git status
```bash
cd /Users/vblake/doe2
git status
git log --oneline -5
```

## ❓ Troubleshooting

**"ImportError: No module named 'pptx'"**
```bash
pip install python-pptx beautifulsoup4
```

**"No images in PowerPoint"**
- Check HTML has `<img>` tags with `data:image/` src
- Verify max_images parameter isn't too low

**"Tables not formatted correctly"**
- Verify HTML has proper `<table>` elements
- Check pandas can read the HTML with `pd.read_html()`

## 📖 Documentation Links

- **Feature Guide**: `POWERPOINT_README.md`
- **API Reference**: `POWERPOINT_EXAMPLES.md`
- **Source Code**: `powerpoint_generator.py`

## ✨ Features

✅ 3 presentation types (Full, Reduced, Comparison)
✅ 50+ embedded images per presentation
✅ Formatted data tables
✅ Professional dark blue theme
✅ Automatic integration with pipeline
✅ Wide format compatibility (PowerPoint, Google Slides, LibreOffice)

## 📊 Output Summary

| File | Size | Slides | Images |
|------|------|--------|--------|
| doe_analysis_report.pptx | 1.5 MB | 25+ | 50+ |
| doe_analysis_reduced.pptx | 1.5 MB | 25+ | 50+ |
| doe_model_comparison.pptx | 1.6 MB | 20+ | 20+ |

---

**Last Updated:** December 1, 2025
**Status:** ✅ Production Ready
