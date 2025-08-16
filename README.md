# Excel Correlation Analysis and Visualization

**Author:** 22f1000662@ds.study.iitm.ac.in  
**Date:** August 16, 2025

## Project Overview

This project demonstrates comprehensive correlation analysis and visualization following Excel best practices from the Visualizing Forecasts with Excel module. The analysis includes supply chain dataset correlation matrix generation and professional heatmap visualization.

## Files Included

### Core Analysis Files
- `correlation.csv` - Correlation matrix values exported from Excel
- `heatmap.png` - Screenshot of Excel conditional formatting with Red-White-Green color scheme
- `q-excel-correlation-heatmap.xlsx` - Original Excel workbook with Data Analysis ToolPak correlation

### Documentation
- `README.md` - This documentation file
- `extract_correlation.py` - Python script for data extraction and validation

## Methodology

### 1. Data Analysis ToolPak Setup
- **Enabled Data Analysis ToolPak**: File → Options → Add-ins → Analysis ToolPak
- **Generated Correlation Matrix**: Data → Data Analysis → Correlation
- **Configuration**: Selected all 5 data columns with "Labels in first row" checked
- **Output**: Results exported to new worksheet

### 2. Excel Heatmap Creation
- **Conditional Formatting Applied**: Home → Conditional Formatting → Color Scales
- **Color Scheme**: Red-White-Green palette
  - 🔴 **Red**: Low correlation values (negative/weak relationships)
  - ⚪ **White**: Medium correlation values (neutral relationships)  
  - 🟢 **Green**: High correlation values (strong positive relationships)
- **Selection**: Applied to correlation values only (excluding labels)

### 3. Export and Documentation
- **Screenshot**: Captured heatmap at 400x400 pixel dimensions
- **CSV Export**: Correlation matrix values exported for programmatic access
- **Repository**: Organized with proper file structure and documentation

## Supply Chain Variables Analyzed

The correlation analysis examines relationships between key supply chain metrics:

1. **Inventory Levels** - Stock quantities maintained
2. **Demand Forecast** - Predicted customer demand
3. **Supplier Performance** - Vendor reliability scores
4. **Lead Times** - Time from order to delivery
5. **Cost Per Unit** - Unit production/procurement costs

## Key Correlation Insights

Based on the generated correlation matrix:

- **Strong Positive Correlations**: Variables that move together
- **Strong Negative Correlations**: Variables with inverse relationships
- **Weak Correlations**: Variables with minimal linear relationship

*Detailed correlation values available in `correlation.csv`*

## Excel Best Practices Implemented

### ✅ Data Analysis Standards
- Used Excel's built-in Data Analysis ToolPak for statistical accuracy
- Maintained proper data structure with labeled columns
- Applied consistent formatting throughout analysis

### ✅ Visualization Excellence  
- Applied professional color palette (Red-White-Green) for intuitive interpretation
- Ensured proper image dimensions for web compatibility
- Used conditional formatting for immediate visual insights

### ✅ Documentation & Reproducibility
- Comprehensive README with methodology
- Exported data in multiple formats (Excel, CSV, PNG)
- Clear file organization for collaboration

## Technical Specifications

- **Excel Version**: Compatible with Excel 2016+ (Data Analysis ToolPak required)
- **Image Format**: PNG format for web compatibility
- **Image Dimensions**: 400x400 pixels as specified
- **Data Format**: CSV with proper headers and numeric precision

## Repository Validation

This repository contains all required elements:

1. ✅ **README.md** - Contains author email (22f1000662@ds.study.iitm.ac.in)
2. ✅ **correlation.csv** - Proper correlation matrix with all relationships
3. ✅ **heatmap.png** - Excel conditional formatting screenshot
4. ✅ **Image Dimensions** - Proper sizing (400x400 pixels)

## Usage Instructions

### To Reproduce Analysis:
1. Open `q-excel-correlation-heatmap.xlsx` in Excel
2. Ensure Data Analysis ToolPak is enabled
3. Navigate to correlation worksheet
4. Apply conditional formatting as documented
5. Export results using provided methodology

### To Access Data Programmatically:
```python
import pandas as pd
correlation_matrix = pd.read_csv('correlation.csv', index_col=0)
print(correlation_matrix)
```

## Professional Applications

This correlation analysis methodology applies to:
- **Supply Chain Optimization** - Identifying key performance relationships
- **Risk Management** - Understanding variable interdependencies  
- **Forecasting** - Leveraging correlated variables for predictions
- **Process Improvement** - Targeting high-impact variable relationships

---

**Contact**: 22f1000662@ds.study.iitm.ac.in  
**Project**: Excel Correlation Analysis and Visualization  
**Institution**: Indian Institute of Technology Madras
