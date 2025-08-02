# 📊 Automated Word Document Generator with Charts

A comprehensive web application that automatically generates professional Word documents with embedded charts and text using Word templates and Excel data.

## 🚀 Features Overview

- **Single Report Generation**: Generate individual reports from Excel files
- **Batch Processing**: Process multiple Excel files simultaneously
- **Advanced Chart Types**: Support for 15+ chart types including bubble, bar_of_pie, and more
- **User Authentication**: Secure login and registration system
- **Project Management**: Create, edit, and manage multiple projects
- **Error Handling**: Comprehensive error reporting and debugging
- **Professional Dashboard**: Modern, responsive user interface

## 🛠️ Technology Stack

### Backend
- **Python Flask**: Web framework
- **MongoDB**: Database for project and user management
- **Matplotlib**: Chart generation and image processing
- **Plotly**: Interactive chart creation
- **python-docx**: Word document manipulation
- **openpyxl**: Excel file processing
- **Pandas**: Data manipulation and analysis

### Frontend
- **React.js**: User interface framework
- **Material-UI (MUI)**: Component library
- **Axios**: HTTP client for API communication
- **React Router**: Navigation and routing

## 📋 Prerequisites

1. **Python 3.8+** installed
2. **Node.js 14+** installed
3. **MongoDB** running locally or cloud instance
4. **Word Template** (.docx file) with placeholders
5. **Excel Files** (.xlsx) with data and chart configurations

## 🏗️ Installation & Setup

### Backend Setup
```bash
cd backend
pip install -r requirements.txt
python app.py
```

### Frontend Setup
```bash
cd frontend-react
npm install
npm start
```

### Environment Variables
Create `.env` files in both backend and frontend directories:

**Backend (.env):**
```
MONGODB_URI=mongodb://localhost:27017/chart_generator
SECRET_KEY=your_secret_key_here
FLASK_ENV=development
```

**Frontend (.env):**
```
REACT_APP_API_URL=http://localhost:5000
```

## 📖 User Guide

### 1. Authentication

#### Registration
1. Navigate to `/register`
2. Fill in your details:
   - Full Name
   - Username
   - Email
   - Password
3. Click "Create Account"
4. You'll be redirected to login

#### Login
1. Navigate to `/login`
2. Enter your username and password
3. Click "Sign In"
4. Access your dashboard

### 2. Project Management

#### Creating a Project
1. Click "+ NEW PROJECT" button
2. Fill in project details:
   - **Project Name**: Descriptive name for your project
   - **Description**: Optional project description
   - **Word Template**: Upload your .docx template file (optional)
3. Click "Create Project"

#### Managing Projects
- **Edit**: Modify project details and template
- **Delete**: Remove projects (irreversible)
- **View Errors**: Check for chart generation issues
- **Generate Reports**: Create reports from Excel data

### 3. Word Template Setup

Your Word template (.docx) should contain placeholders where content will be inserted.

#### Placeholder Rules
- **Format**: `${placeholder_name}`
- **Location**: Must be inside text, not in shapes or headers/footers
- **Case-insensitive**: `${Section1_Text}` = `${section1_text}` = `${SECTION1_TEXT}`
- **Allowed characters**: Letters, numbers, underscore (_)

#### Valid Examples
```
${section1_text} → gets replaced with text content
${section1_chart} → gets replaced with chart image
${company_name} → gets replaced with company name
${quarterly_summary} → gets replaced with summary text
```

### 4. Excel Data Structure

Your Excel file must contain specific columns for data and chart configuration.

#### Required Columns

| Column | Purpose | Example |
|--------|---------|---------|
| **Text_Tag** | Matches placeholder in Word template | `section1_text` |
| **Text** | Actual text content to insert | `This is the first section content.` |
| **Chart_Tag** | Matches chart placeholder in Word template | `section1_chart` |
| **Chart_Type** | Type of chart to generate | `bar`, `line`, `pie`, `bubble`, `bar_of_pie` |
| **Chart_Attributes** | JSON configuration for chart | See chart configuration examples below |

#### Optional Columns
- **Chart_Data_Y2023**: Year-specific data
- **Chart_Data_Y2024**: Year-specific data
- **Chart_Data_CAGR**: Compound Annual Growth Rate
- **Report_Name**: For batch processing (extracted from Excel)
- **Report_Code**: For batch processing (extracted from Excel)

### 5. Chart Types & Configurations

#### 5.1 Bar Chart
```json
{
  "chart_meta": {
    "chart_title": "Monthly Sales",
    "x_label": "Month",
    "primary_y_label": "Sales ($)",
    "show_gridlines": true,
    "data_labels": true
  },
  "series": {
    "labels": ["Jan", "Feb", "Mar", "Apr"],
    "data": [
      {
        "name": "Sales",
        "type": "bar",
        "values": [100, 150, 200, 175],
        "marker": {
          "color": ["#FF5733", "#33C3FF", "#9B59B6", "#F39C12"]
        }
      }
    ]
  }
}
```

#### 5.2 Line Chart
```json
{
  "chart_meta": {
    "chart_title": "Revenue Trend",
    "x_label": "Quarter",
    "primary_y_label": "Revenue ($K)",
    "show_gridlines": true
  },
  "series": {
    "labels": ["Q1", "Q2", "Q3", "Q4"],
    "data": [
      {
        "name": "Revenue",
        "type": "line",
        "values": [500, 650, 800, 950],
        "marker": {
          "color": "#FF5733",
          "size": 8
        }
      }
    ]
  }
}
```

#### 5.3 Pie Chart
```json
{
  "chart_meta": {
    "chart_title": "Market Share",
    "data_labels": true,
    "value_format": ".1f"
  },
  "series": {
    "labels": ["Company A", "Company B", "Company C", "Others"],
    "values": [35, 25, 20, 20],
    "colors": ["#FF5733", "#33C3FF", "#9B59B6", "#F39C12"]
  }
}
```

#### 5.4 Bar of Pie Chart
```json
{
  "chart_meta": {
    "chart_title": "Product Sales Overview",
    "value_format": ".0f"
  },
  "series": {
    "labels": ["Product A", "Product B", "Other"],
    "values": [40, 20, 70],
    "colors": ["#FF5733", "#33C3FF", "#9B59B6"],
    "other_labels": ["Category X", "Category Y", "Category Z"],
    "other_values": [30, 25, 15],
    "other_colors": ["#E74C3C", "#3498DB", "#F39C12"]
  }
}
```

#### 5.5 Bubble Chart
```json
{
  "chart_meta": {
    "chart_title": "Product Performance Analysis",
    "x_label": "Units Sold",
    "primary_y_label": "Revenue ($K)",
    "disable_secondary_y": true,
    "showlegend": false
  },
  "series": {
    "labels": ["Products"],
    "x_axis": [100, 200, 150, 250],
    "data": [
      {
        "name": "Product Performance",
        "type": "bubble",
        "values": [20000, 40000, 30000, 50000],
        "size": [200, 500, 600, 800],
        "marker": {
          "color": ["#E74C3C", "#3498DB", "#F39C12", "#27AE60"],
          "opacity": 0.85,
          "line": {
            "color": "#FFFFFF",
            "width": 2
          }
        }
      }
    ]
  }
}
```

#### 5.6 Area Chart
```json
{
  "chart_meta": {
    "chart_title": "Cumulative Sales",
    "x_label": "Month",
    "primary_y_label": "Sales ($K)",
    "show_gridlines": true
  },
  "series": {
    "labels": ["Jan", "Feb", "Mar", "Apr", "May"],
    "data": [
      {
        "name": "Sales",
        "type": "area",
        "values": [100, 250, 450, 625, 800],
        "marker": {
          "color": "#FF5733",
          "opacity": 0.7
        }
      }
    ]
  }
}
```

### 6. Report Generation

#### 6.1 Single Report Generation
1. Select a project from your dashboard
2. Click "Generate Report" button
3. Choose "Single Report" tab
4. Upload your Excel file (.xlsx or .csv)
5. Click "Generate Report"
6. The system will:
   - Process your Excel data
   - Generate charts based on configurations
   - Create a Word document with embedded content
   - Automatically download the generated report

#### 6.2 Batch Report Generation
1. Select a project from your dashboard
2. Click "Generate Report" button
3. Choose "Batch Reports" tab
4. Upload a ZIP file containing multiple Excel files
5. Click "Generate Report"
6. The system will:
   - Extract Excel files from ZIP
   - Process each file individually
   - Generate reports for each Excel file
   - Create two folder structures:
     - `reports_by_name/`: Files named by Report_Name from Excel
     - `reports_by_code/`: Files named by Report_Code from Excel
   - Download a ZIP file containing all generated reports

#### Batch Processing Features
- **Excel Column Extraction**: Automatically reads `Report_Name` and `Report_Code` columns
- **Dual Folder Structure**: Organizes reports by name and code
- **Progress Tracking**: Real-time progress updates
- **Error Handling**: Continues processing even if some files fail
- **Automatic Cleanup**: Temporary files are cleaned up after download

### 7. Dashboard Features

#### 7.1 Project Overview
- **Total Projects**: Count of all projects
- **Recent Projects**: Projects created in last 7 days
- **Successful Reports**: Reports generated without errors
- **Projects with Issues**: Reports with generation errors

#### 7.2 Project Management
- **Search**: Find projects by name or description
- **Filter**: Filter by all projects or recent projects
- **Sort**: Sort by date created or project name
- **View Modes**: Toggle between table and grid views

#### 7.3 Project Actions
- **Generate Report**: Create single or batch reports
- **Edit Project**: Modify project details and template
- **View Errors**: Check for chart generation issues
- **Delete Project**: Remove projects permanently

### 8. Error Handling & Debugging

#### 8.1 Chart Generation Errors
- **Data Issues**: Missing or invalid data
- **Configuration Errors**: Invalid JSON in Chart_Attributes
- **Chart Type Issues**: Unsupported chart types
- **Size Issues**: Charts too large or small

#### 8.2 Report Generation Errors
- **Template Issues**: Missing or invalid Word template
- **Placeholder Issues**: Mismatched placeholders
- **File Access Issues**: Permission or path problems

#### 8.3 Error Reporting
- **Visual Indicators**: Color-coded error status
- **Detailed Messages**: Specific error descriptions
- **Error Logs**: Comprehensive logging for debugging
- **Error Clearing**: Option to clear error history

### 9. Advanced Features

#### 9.1 Chart Customization
- **Colors**: Custom color schemes for charts
- **Fonts**: Configurable font families and sizes
- **Gridlines**: Optional grid display
- **Data Labels**: Show/hide value labels
- **Axis Formatting**: Custom axis labels and formats

#### 9.2 Annotations
Add text annotations to charts:
```json
{
  "annotations": [
    {
      "text": "Peak Performance",
      "x_value": "Q3",
      "y_value": 800,
      "color": "red"
    }
  ]
}
```

#### 9.3 Margin Control
Adjust chart margins:
```json
{
  "margin": {
    "l": 50,
    "r": 50,
    "t": 30,
    "b": 50
  }
}
```

### 10. Best Practices

#### 10.1 Excel File Preparation
- **Data Validation**: Ensure data is clean and consistent
- **Column Headers**: Use exact column names as specified
- **JSON Formatting**: Validate JSON in Chart_Attributes column
- **File Size**: Keep Excel files under 10MB for optimal performance

#### 10.2 Word Template Design
- **Placeholder Placement**: Position placeholders logically
- **Formatting**: Apply consistent formatting to placeholders
- **Spacing**: Leave adequate space for charts
- **Testing**: Test with sample data before production use

#### 10.3 Chart Configuration
- **Color Schemes**: Use consistent colors across charts
- **Data Ranges**: Ensure data ranges are appropriate for chart types
- **Labels**: Use clear, descriptive labels
- **Sizing**: Consider chart size relative to document layout

### 11. Troubleshooting

#### 11.1 Common Issues
- **Charts Not Appearing**: Check Chart_Attributes JSON format
- **Text Not Replacing**: Verify placeholder names match Excel Text_Tag
- **Download Failures**: Check browser download settings
- **Performance Issues**: Reduce file sizes or number of charts

#### 11.2 Error Messages
- **"Invalid JSON"**: Check Chart_Attributes column formatting
- **"Missing Data"**: Ensure all required columns are present
- **"Template Error"**: Verify Word template file integrity
- **"Chart Generation Failed"**: Review chart configuration

### 12. API Reference

#### 12.1 Authentication Endpoints
- `POST /api/register` - User registration
- `POST /api/login` - User login
- `GET /api/logout` - User logout
- `GET /api/user` - Get current user info

#### 12.2 Project Endpoints
- `GET /api/projects` - List all projects
- `POST /api/projects` - Create new project
- `PUT /api/projects/{id}` - Update project
- `DELETE /api/projects/{id}` - Delete project
- `GET /api/projects/{id}` - Get project details

#### 12.3 Report Generation Endpoints
- `POST /api/projects/{id}/upload_report` - Generate single report
- `POST /api/projects/{id}/upload_zip` - Generate batch reports
- `GET /api/reports/{id}/download` - Download single report
- `GET /api/reports/batch_reports_{id}.zip` - Download batch reports

#### 12.4 Error Management Endpoints
- `GET /api/projects/{id}/chart_errors` - Get chart errors
- `POST /api/projects/{id}/clear_errors` - Clear project errors

## 📞 Support

For technical support or questions:
- Check the error logs in the dashboard
- Review the troubleshooting section
- Ensure all prerequisites are met
- Verify file formats and configurations

## 🔄 Version History

### v2.0.0 (Current)
- Enhanced batch processing with dual folder structure
- Improved chart types (bubble, bar_of_pie, etc.)
- Professional dashboard with search and filtering
- Enhanced error handling and reporting
- Modern UI with Material-UI components

### v1.0.0
- Basic single report generation
- Simple chart types (bar, line, pie)
- Basic project management
- Simple authentication system

---

**Note**: This documentation covers all current features. For the latest updates and new features, please refer to the project repository or contact the development team. 