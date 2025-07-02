# Chart Configuration Guide

This guide explains how to configure interactive charts in your Excel-based report generation system. The system now supports **pie charts**, **stacked column charts**, **area charts**, and **expanded pie charts** with one segment shown as a bar chart.

## Supported Chart Types

### 1. Stacked Column Chart
Perfect for showing multiple data series stacked on top of each other.

**Configuration Example:**
```json
{
  "chart_meta": {
    "chart_type": "stacked_column",
    "legend": true,
    "data_labels": true,
    "value_format": "00.0%",
    "source_sheet": "Revenue_By_Year",
    "category_range": "O20:O23",
    "series_range": ["W9", "AA9", "AF9"],
    "value_range": ["W20:W23", "AA20:AA23", "AF20:AF23"]
  },
  "series": {
    "colors": ["#b7b7b7", "#ffc000", "#0070c0"],
    "labels": ["Commission-Based", "Service/Refurbishment", "Ad-Based"]
  }
}
```

**Features:**
- ✅ Multiple series stacked vertically
- ✅ Custom colors for each series
- ✅ Data labels with percentage formatting
- ✅ Interactive legend
- ✅ Excel data extraction from specified ranges

### 2. Pie Chart
Ideal for showing proportions and percentages of a whole.

**Configuration Example:**
```json
{
  "chart_meta": {
    "chart_type": "pie",
    "legend": true,
    "data_labels": true,
    "value_format": "0.0%",
    "source_sheet": "Pie_Data",
    "category_range": "E23:E29",
    "value_range": "AA23:AA29"
  },
  "series": {
    "colors": ["#002060", "#ED7D31", "#548235", "#FFC000", "#4472C4", "#70AD47", "#9DC3E6"],
    "labels": ["Retail", "Home Improvement", "Travel/Entertainment", "Services", "Automotive", "Health Care and Wellness", "Others"]
  }
}
```

**Features:**
- ✅ Custom colors for each segment
- ✅ Data labels showing percentage and value
- ✅ Interactive legend
- ✅ Excel data extraction
- ✅ Hover tooltips with detailed information

### 3. Expanded Pie Chart
Advanced pie chart with one segment expanded into a detailed bar chart.

**Configuration Example:**
```json
{
  "chart_meta": {
    "chart_type": "pie",
    "legend": true,
    "data_labels": true,
    "value_format": "0.0%",
    "source_sheet": "Pie_Data",
    "category_range": "F86:F92",
    "value_range": "AA86:AA92",
    "expanded_segment": "Consumer Electronics"
  },
  "series": {
    "colors": ["#31859C", "#378AFF", "#93F03B", "#F54F52", "#FFA32F", "#E7D61E", "#a2bee4"],
    "labels": [
      "Apparel, Footwear & Accessories",
      "Consumer Electronics",
      "Toys, Kids, and Babies",
      "Jewellery",
      "Sporting Goods",
      "Entertainment & Gaming",
      "Other"
    ]
  }
}
```

**Features:**
- ✅ Side-by-side pie and bar chart layout
- ✅ Detailed view of selected segment
- ✅ Consistent color scheme
- ✅ Interactive elements
- ✅ Professional presentation

### 4. Area Chart
Great for showing trends over time with filled areas.

**Configuration Example:**
```json
{
  "chart_meta": {
    "chart_type": "area",
    "legend": true,
    "data_labels": true,
    "value_format": "0.0%",
    "source_sheet": "Area_Data",
    "category_range": "A1:A5",
    "value_range": ["B1:B5", "C1:C5", "D1:D5"]
  },
  "series": {
    "colors": ["#FF6B6B", "#4ECDC4", "#45B7D1"],
    "labels": ["Revenue", "Costs", "Profit"]
  }
}
```

**Features:**
- ✅ Filled areas for visual impact
- ✅ Multiple series support
- ✅ Custom colors and labels
- ✅ Interactive tooltips
- ✅ Excel data integration

## Configuration Parameters

### Chart Meta Parameters

| Parameter | Type | Description | Required |
|-----------|------|-------------|----------|
| `chart_type` | string | Chart type: "pie", "stacked_column", "area" | Yes |
| `legend` | boolean | Show/hide legend | No (default: true) |
| `data_labels` | boolean | Show/hide data labels | No (default: true) |
| `value_format` | string | Format for values (e.g., "0.0%", "00.0%") | No |
| `source_sheet` | string | Excel sheet name to extract data from | Yes |
| `category_range` | string | Excel range for categories (e.g., "A1:A5") | Yes |
| `value_range` | array/string | Excel range(s) for values | Yes |
| `expanded_segment` | string | For expanded pie charts, specify segment to expand | No |

### Series Parameters

| Parameter | Type | Description | Required |
|-----------|------|-------------|----------|
| `colors` | array | Array of hex color codes | Yes |
| `labels` | array | Array of series labels | Yes |

## Excel Data Structure

### For Stacked Column Charts
```
Sheet: Revenue_By_Year
- Categories: O20:O23 (Years: 2020, 2021, 2022, 2023)
- Series 1: W20:W23 (Commission-Based values)
- Series 2: AA20:AA23 (Service/Refurbishment values)
- Series 3: AF20:AF23 (Ad-Based values)
```

### For Pie Charts
```
Sheet: Pie_Data
- Categories: E23:E29 (Market segments)
- Values: AA23:AA29 (Percentage values)
```

### For Area Charts
```
Sheet: Area_Data
- Categories: A1:A5 (Time periods)
- Series 1: B1:B5 (Revenue values)
- Series 2: C1:C5 (Cost values)
- Series 3: D1:D5 (Profit values)
```

## Advanced Features

### Value Formatting
- `"0.0%"` - Shows as "25.5%"
- `"00.0%"` - Shows as "025.5%"
- `"0.0"` - Shows as "25.5"
- `"$0.0"` - Shows as "$25.5"

### Color Schemes
Use hex color codes for professional appearance:
- `"#002060"` - Dark blue
- `"#ED7D31"` - Orange
- `"#548235"` - Green
- `"#FFC000"` - Yellow
- `"#4472C4"` - Blue
- `"#70AD47"` - Light green
- `"#9DC3E6"` - Light blue

### Interactive Features
- **Hover Tooltips**: Detailed information on hover
- **Legend**: Click to show/hide series
- **Zoom**: Pan and zoom functionality
- **Download**: Save as PNG or SVG
- **Responsive**: Adapts to different screen sizes

## Implementation in Your Project

### 1. Excel Template Setup
Create your Excel file with the data structure as shown above.

### 2. Chart Configuration
Add your chart configuration to the Excel file in the appropriate cells, following the JSON format examples.

### 3. Word Template
In your Word template, use placeholders like:
```
${section1_chart}
${section2_chart}
${section3_chart}
```

### 4. Report Generation
The system will automatically:
1. Extract data from Excel ranges
2. Generate interactive Plotly charts
3. Create static Matplotlib charts for Word documents
4. Save both interactive HTML and static PNG files

## Testing

Run the test script to verify functionality:
```bash
cd backend
python test_charts.py
```

This will generate sample HTML files for each chart type.

## Troubleshooting

### Common Issues

1. **Data not loading**: Check Excel sheet names and ranges
2. **Colors not applying**: Ensure hex color codes are valid
3. **Charts not rendering**: Verify JSON configuration syntax
4. **Missing data labels**: Check `data_labels` parameter

### Debug Tips

1. Check the Flask application logs for detailed error messages
2. Verify Excel file format and data structure
3. Test with sample data first
4. Use the test script to validate configurations

## Best Practices

1. **Consistent Naming**: Use clear, descriptive names for sheets and ranges
2. **Color Harmony**: Choose complementary colors for professional appearance
3. **Data Validation**: Ensure Excel data is clean and properly formatted
4. **Performance**: Limit chart complexity for large datasets
5. **Accessibility**: Use high contrast colors and clear labels

## Support

For additional help or feature requests, please refer to the project documentation or contact the development team. 