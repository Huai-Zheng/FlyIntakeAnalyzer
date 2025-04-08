# FlyIntakeAnalyzer - Drosophila Food Intake Analysis Tool
## Introduction

A PyQt5-based GUI application for analyzing Drosophila feeding experiment data through absorbance measurements. Automates standard curve calculation, sample concentration determination, and food intake quantification to generate standardized reports.

## Key Features

- 🖼️ User-friendly GUI interface
- 📊 Automatic standard curve fitting with R² calculation
- ⚡ Batch processing of sample data
- 📈 Error-bar embedded bar charts
- 📉 Standard curve plotting with trendline
- 📁 Excel report generation
- ✅ Auto-column-width adjustment and styling

## Quick Start

### Requirements
- Python 3.7+
- Dependencies: PyQt5, pandas, numpy, scipy, openpyxl

### Installation
```bash
# Clone repository
git clone https://github.com/Huai-Zheng/FlyIntakeAnalyzer.git

# Install dependencies
pip install -r requirements.txt

# Windows users: Pre-built executable available
```
### Usage
1. Prepare Excel file containing:
    "OD" sheet with properly formatted data
2. Run application:
```bash
python Main.py
```
3. Click 'Browse' to select input file
4. Click 'Analyze' to start processing
5. Results will be saved as *_RESULT.xlsx in input directory
### File Format Requirements
Input Excel file must contain:
```markdown
OD Sheet:
- Standard curve data: Cells C25-C32
- Reference concentration (C0): Cells D25-D32
- Sample data: Columns E-N (10 groups), rows 25-32
```
### Output Includes
- 📄 Results Sheet: Detailed calculations for all samples
- 📊 Summary Sheet: Group statistics (Mean ± SD)
- 📈 Standard_Curve Sheet: Standard curve plot and formula
- 📉 Auto-generated comparison charts
### Important Notes
⚠️ Before first use:
1. Install all required dependencies
2. Verify Excel file format compliance
3. Ensure write permissions for output directory
4. Do not modify source file during analysis
### Technical Details
- Multi-threading processing prevents UI freezing
- Supports Excel 2007+ formats (.xlsx)
- Responsive UI for different screen resolutions
- Built-in progress bar and logging system
## License
MIT License
## Contributing
Welcome to submit issues and pull requests! For feature requests or bug reports, please use the issue tracker.
## Version 1.0 | Updated 2025-3-1 | Developed by [Huai-Zheng Zheng]
```Markdown
Adaptation notes:
1. Maintained all functional details from Chinese version
2. Used technical terms common in life science research (e.g., "absorbance", "Drosophila")
3. Kept emoji visual markers for better document scanning
4. Added clear imperative verbs in installation/usage sections
5. Formatted Excel cell references as code blocks for clarity
6. Used standard English technical documentation structure
7. Added "Important Notes" section header for better accessibility

To complete the internationalization:
1. Provide English version of example Excel file
2. Add multilingual UI support in future versions
3. Include unit conversion notes for international researchers
4. Add documentation for statistical methods used (linear regression implementation)
```

