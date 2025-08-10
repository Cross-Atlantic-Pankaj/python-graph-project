import os
import json
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Use non-GUI backend suitable for Flask servers
import matplotlib.pyplot as plt
from flask import Blueprint, request, jsonify, current_app, send_file
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename
from bson.objectid import ObjectId
from datetime import datetime 
from docx import Document
from docx.shared import Inches
from docx.text.paragraph import Paragraph
import openpyxl
import tempfile
import re
import zipfile
import shutil
import plotly.graph_objects as go
from plotly.subplots import make_subplots


import re

# Define a constant for the section1_chart attribut

# Files are now stored in database, no upload folder needed

ALLOWED_EXTENSIONS = {'txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif', 'csv', 'xlsx', 'docx'}
ALLOWED_REPORT_EXTENSIONS = {'csv', 'xlsx'}

projects_bp = Blueprint('projects', __name__)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def allowed_report_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_REPORT_EXTENSIONS

def extract_report_info_from_excel(excel_path):
    """Extract Report_Name and Report_Code from Excel file"""
    try:
        # Load the Excel file
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet = wb.active
        
        report_name = None
        report_code = None
        
        # Search for the columns in the first few rows
        for row_idx in range(1, min(10, sheet.max_row + 1)):  # Check first 10 rows
            for col_idx in range(1, min(10, sheet.max_column + 1)):  # Check first 10 columns
                cell_value = sheet.cell(row=row_idx, column=col_idx).value
                
                if cell_value and isinstance(cell_value, str):
                    cell_value = cell_value.strip()
                    
                    # Look for Report_Name column
                    if cell_value.lower() == 'report_name':
                        # Get the value from the next row in the same column
                        if row_idx + 1 <= sheet.max_row:
                            report_name = sheet.cell(row=row_idx + 1, column=col_idx).value
                            if report_name:
                                report_name = str(report_name).strip()
                    
                    # Look for Report_Code column
                    elif cell_value.lower() == 'report_code':
                        # Get the value from the next row in the same column
                        if row_idx + 1 <= sheet.max_row:
                            report_code = sheet.cell(row=row_idx + 1, column=col_idx).value
                            if report_code:
                                report_code = str(report_code).strip()
        
        wb.close()
        
        # Fallback to filename if not found
        if not report_name:
            report_name = os.path.splitext(os.path.basename(excel_path))[0]
            #current_app.logger.warning(f"Report_Name not found in {excel_path}, using filename: {report_name}")
        
        if not report_code:
            report_code = f"REPORT_{os.path.splitext(os.path.basename(excel_path))[0]}"
            #current_app.logger.warning(f"Report_Code not found in {excel_path}, using generated code: {report_code}")
        
        current_app.logger.info(f"Extracted from {excel_path}: Report_Name='{report_name}', Report_Code='{report_code}'")
        return report_name, report_code
        
    except Exception as e:
        current_app.logger.error(f"Error extracting report info from {excel_path}: {e}")
        # Fallback to filename-based naming
        fallback_name = os.path.splitext(os.path.basename(excel_path))[0]
        fallback_code = f"REPORT_{fallback_name}"
        return fallback_name, fallback_code

def create_expanded_pie_chart(labels, values, colors, expanded_segment, title, value_format=""):
    """
    Create an expanded pie chart with one segment shown as a bar chart
    """
    fig = make_subplots(
        rows=1, cols=2,
        specs=[[{"type": "pie"}, {"type": "bar"}]],
        subplot_titles=(title, f"{expanded_segment} Details")
    )
    
    # Add pie chart
    fig.add_trace(go.Pie(
        labels=labels,
        values=values,
        textinfo="label+percent",
        textposition="outside",
        marker=dict(colors=colors)
    ), row=1, col=1)
    
    # Add bar chart for expanded segment
    if expanded_segment in labels:
        segment_idx = labels.index(expanded_segment)
        segment_value = values[segment_idx]
        segment_color = colors[segment_idx] if segment_idx < len(colors) else colors[0]
        
        fig.add_trace(go.Bar(
            x=[expanded_segment],
            y=[segment_value],
            marker_color=segment_color,
            text=[f"{segment_value}{value_format}"],
            textposition="auto"
        ), row=1, col=2)
    
    fig.update_layout(
        title_text=title,
        showlegend=False,
        height=500
    )
    
    return fig

def create_bar_of_pie_chart(labels, values, other_labels, other_values, colors, other_colors, title, value_format=""):
    """
    Create a 'bar of pie' chart using Plotly: pie chart with one segment broken down as a bar chart.
    """
    # Filter out empty/null values from other_labels and other_values
    filtered_data = []
    for label, value in zip(other_labels, other_values):
        # Check if both label and value are not empty/null
        if (label is not None and str(label).strip() != "" and 
            value is not None and str(value).strip() != "" and 
            str(value).strip() != "0"):
            filtered_data.append((label, value))
    
    if filtered_data:
        filtered_labels, filtered_values = zip(*filtered_data)
    else:
        # If no valid data, use empty lists
        filtered_labels, filtered_values = [], []
    
    fig = make_subplots(
        rows=1, cols=2,
        specs=[[{"type": "pie"}, {"type": "bar"}]],
        column_widths=[0.5, 0.5],
        subplot_titles=(title, "Breakdown of 'Other'")
    )
    # Main pie chart
    fig.add_trace(go.Pie(
        labels=labels,
        values=values,
        marker=dict(colors=colors),
        textinfo="percent",
        hoverinfo="label+percent+value",
        pull=[0.1 if l == "Other" else 0 for l in labels],
        name="Main Pie"
    ), row=1, col=1)
    
    # Bar chart for breakdown (only if we have filtered data)
    if filtered_labels and filtered_values:
        # Create individual bar traces for each data point to avoid stacking
        for i, (label, value) in enumerate(zip(filtered_labels, filtered_values)):
            # Use individual colors for each bar
            bar_color = other_colors[i] if other_colors and i < len(other_colors) else None
            
            # Format the value properly for display - show as XX.X% instead of 0.XXX
            if isinstance(value, (int, float)):
                if value <= 1.0:  # Likely decimal format (0.11)
                        display_value = f"{value * 100:.1f}%"
            else:  # Likely already percentage format (11.0)
                        display_value = f"{value:.1f}%"
        else:
                    # Convert string to float and handle
                    try:
                        val = float(value)
                        if val <= 1.0:
                            display_value = f"{val * 100:.1f}%"
                        else:
                            display_value = f"{val:.1f}%"
                    except:
                        display_value = str(value)
            
            # Format the x-axis label as percentage
        if isinstance(label, (int, float)):
                if label <= 1.0:  # Likely decimal format (0.06)
                    formatted_label = f"{label * 100:.1f}%"
                else:  # Likely already percentage format (6.0)
                    formatted_label = f"{label:.1f}%"
        else:
                # Convert string to float and handle
                try:
                    val = float(label)
                    if val <= 1.0:
                        formatted_label = f"{val * 100:.1f}%"
                    else:
                        formatted_label = f"{val:.1f}%"
                except:
                    formatted_label = str(label)
            
        fig.add_trace(go.Bar(
            x=[formatted_label],  # Use formatted label for x-axis
            y=[value],  # Single y value for each bar
            marker_color=bar_color,
            text=[display_value],
            textposition="auto",
            name=f"Breakdown {i+1}",
            showlegend=False  # Hide individual bar legends
        ), row=1, col=2)
    
    fig.update_layout(
        title_text=title,
        showlegend=False,
        height=500,
        width=900
    )
    
    # Update Y-axis formatting for the bar chart to show percentages
    if filtered_values and isinstance(filtered_values[0], (int, float)) and filtered_values[0] <= 1.0:
        fig.update_yaxes(
            tickformat=".1%",  # Format as percentage with 1 decimal place
            row=1, col=2
        )
    
    return fig

def _generate_report(project_id, template_path, data_file_path):
    import pandas as pd
    import json
    import tempfile
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
    from docx import Document
    from docx.shared import Inches
    import re
    import os

    plt.style.use('ggplot')  # 👈 Apply a cleaner visual style

    try:
        # Report generation started

        # Try to read Excel with original formatting preserved
        try:
            df = pd.read_excel(data_file_path, sheet_name=0, keep_default_na=False)  # Use first sheet
        except:
            df = pd.read_excel(data_file_path, sheet_name=0)  # Fallback to default
        
        df.columns = df.columns.str.strip().str.replace(" ", "_").str.replace("__", "_")
        
        # Log the raw data to see what pandas is reading
        # Excel data loaded successfully

        # Excel structure loaded silently
        
        # Validate required columns exist
        required_columns = ["Text_Tag", "Text", "Chart_Tag", "Chart_Attributes", "Chart_Type"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            current_app.logger.error(f"❌ Missing required columns: {missing_columns}")
            raise ValueError(f"Excel file missing required columns: {missing_columns}")

        text_map = {str(k).strip().lower(): str(v).strip() for k, v in zip(df["Text_Tag"], df["Text"]) if pd.notna(k) and pd.notna(v)}
        chart_attr_map = {str(k).strip().lower(): str(v).strip() for k, v in zip(df["Chart_Tag"], df["Chart_Attributes"]) if pd.notna(k) and pd.notna(v)}
        chart_type_map = {str(k).strip().lower(): str(v).strip() for k, v in zip(df["Chart_Tag"], df["Chart_Type"]) if pd.notna(k) and pd.notna(v)}
        
        # Data maps created silently

        flat_data_map = {}
        # Add global metadata columns that can be used across all sections
        global_metadata = {}
        
        # First pass: Process global metadata columns from ALL rows
        for _, row in df.iterrows():
            for col in df.columns:
                col_lower = col.lower().strip()
                
                # Handle global metadata columns (can be used across all sections)
                if col_lower in ['country', 'report_name', 'report_code', 'currency']:
                    value = row[col]
                    if pd.notna(value) and str(value).strip():
                        # Store in global metadata for use across all sections
                        global_metadata[col_lower] = str(value).strip()
                        # Also add to flat_data_map with the column name as key
                        flat_data_map[col_lower] = str(value).strip()
                        current_app.logger.info(f"📋 LOADED: {col_lower} = '{str(value).strip()}'")
                    else:
                        # Empty value - skip silently
                        pass
        
        # Ensure we have all required global metadata
        required_global_metadata = ['country', 'report_name', 'report_code', 'currency']
        missing_metadata = [key for key in required_global_metadata if key not in flat_data_map]
        if missing_metadata:
            current_app.logger.error(f"❌ MISSING: {missing_metadata}")
        else:
            current_app.logger.info(f"✅ DATA LOADED: {flat_data_map}")
        
        # Second pass: Process chart-specific data from rows with chart tags
        for _, row in df.iterrows():
            chart_tag = row.get("Chart_Tag")
            if not isinstance(chart_tag, str) or not chart_tag:
                continue
            section_prefix = chart_tag.replace('_chart', '').lower()

            # Process all columns systematically
            for col in df.columns:
                col_lower = col.lower().strip()
                
                # Skip global metadata columns as they're already processed
                if col_lower in ['country', 'report_name', 'report_code', 'currency']:
                    continue
                
                # Handle Chart Data columns (for values)
                if col_lower.startswith("chart_data_y"):
                    year = col.replace("Chart_Data_", "").replace("chart_data_", "")
                    key = f"{section_prefix}_{year.lower()}"
                    value = row[col]
                    if pd.notna(value) and str(value).strip():
                        # Format values to always show one decimal place
                        try:
                            float_val = float(value)
                            formatted_val = f"{float_val:.1f}"
                            flat_data_map[key] = formatted_val
                        except (ValueError, TypeError):
                            # If conversion fails, use the original value as string
                            flat_data_map[key] = str(value).strip()
                    else:
                        # Empty value - skip silently
                        pass
                
                # Handle Growth columns (for growth rates)
                elif col_lower.startswith("growth_y"):
                    year = col.replace("Growth_", "").replace("growth_", "")
                    key = f"{section_prefix}_{year.lower()}_kpi2"
                    value = row[col]
                    if pd.notna(value) and str(value).strip():
                        # Convert percentage values to proper percentage format for table display
                        try:
                            # Convert to float first to handle any numeric format
                            float_val = float(value)
                            
                            # Check if the original value string contains '%' (Excel percentage format)
                            original_str = str(value).strip()
                            if '%' in original_str:
                                # Remove % and format as percentage
                                clean_val = original_str.replace('%', '').strip()
                                float_val = float(clean_val)
                                percentage_val = f"{float_val:.1f}%"
                                flat_data_map[key] = percentage_val
                                # Excel percentage formatted as percentage
                            elif float_val > 1:
                                # Convert percentage to percentage format (e.g., 20 -> 20.0%)
                                percentage_val = f"{float_val:.1f}%"
                                flat_data_map[key] = percentage_val
                                # Percentage formatted as percentage
                            else:
                                # If it's already a decimal, convert to percentage (e.g., 0.05 -> 5.0%)
                                percentage_val = f"{float_val * 100:.1f}%"
                                flat_data_map[key] = percentage_val
                                # Decimal converted to percentage
                        except (ValueError, TypeError):
                            # If conversion fails, use the original value as string
                            flat_data_map[key] = str(value).strip()
                            # Conversion failed - skip silently
                    else:
                        # Empty value - skip silently
                        pass

            # Handle CAGR
            if pd.notna(row.get("Chart_Data_CAGR")):
                key = f"{section_prefix}_cgrp"
                value = row["Chart_Data_CAGR"]
                if str(value).strip():
                    # Convert CAGR percentage values to proper percentage format for table display
                    try:
                        # Convert to float first to handle any numeric format
                        float_val = float(value)
                        
                        # Check if the original value string contains '%' (Excel percentage format)
                        original_str = str(value).strip()
                        if '%' in original_str:
                            # Remove % and format as percentage
                            clean_val = original_str.replace('%', '').strip()
                            float_val = float(clean_val)
                            percentage_val = f"{float_val:.1f}%"
                            flat_data_map[key] = percentage_val
                            # Excel CAGR percentage formatted as percentage
                        elif float_val > 1:
                            # Convert percentage to percentage format (e.g., 10.5 -> 10.5%)
                            percentage_val = f"{float_val:.1f}%"
                            flat_data_map[key] = percentage_val
                            # CAGR percentage formatted as percentage
                        else:
                            # If it's already a decimal, convert to percentage (e.g., 0.105 -> 10.5%)
                            percentage_val = f"{float_val * 100:.1f}%"
                            flat_data_map[key] = percentage_val
                            # CAGR decimal converted to percentage
                    except (ValueError, TypeError):
                        # If conversion fails, use the original value as string
                        flat_data_map[key] = str(value).strip()
                        # Conversion failed - skip silently
                else:
                    # Empty value - skip silently
                    pass

        # Ensure all keys are lowercase
        flat_data_map = {k.lower(): v for k, v in flat_data_map.items()}
        
        # Log the final data map for debugging
        # Data map prepared for replacement
        
        # Data mapping completed silently

        doc = Document(template_path)

        def replace_text_in_paragraph(paragraph):
            nonlocal flat_data_map, text_map  # Access variables from outer scope
            
            # Simple approach: replace placeholders directly in each run without clearing runs
            for run in paragraph.runs:
                original_text = run.text
                modified_text = original_text
                
                # Handle ${} format placeholders
                matches = re.findall(r"\$\{(.*?)\}", run.text)
                for match in matches:
                    key_lower = match.lower().strip()
                    val = flat_data_map.get(key_lower) or text_map.get(key_lower)
                    if val:
                        pattern = re.compile(re.escape(f"${{{match}}}"), re.IGNORECASE)
                        modified_text = pattern.sub(val, modified_text)
                        # Log replacements for global metadata
                        if key_lower in ['country', 'report_name', 'report_code', 'currency']:
                            # Replaced placeholder successfully
                            pass
                    else:
                        if key_lower in ['country', 'report_name', 'report_code', 'currency']:
                            current_app.logger.error(f"❌ NO DATA: ${{{match}}} (key: {key_lower})")
                
                # Handle <> format placeholders
                angle_matches = re.findall(r"<(.*?)>", run.text)
                for match in angle_matches:
                    key_lower = match.lower().strip()
                    val = flat_data_map.get(key_lower) or text_map.get(key_lower)
                    if val:
                        pattern = re.compile(re.escape(f"<{match}>"), re.IGNORECASE)
                        modified_text = pattern.sub(val, modified_text)
                        # Log replacements for global metadata
                        if key_lower in ['country', 'report_name', 'report_code', 'currency']:
                            # Replaced tag successfully
                            pass
                    else:
                        if key_lower in ['country', 'report_name', 'report_code', 'currency']:
                            current_app.logger.error(f"❌ NO DATA: <{match}> (key: {key_lower})")
                
                # Update the run text only if it was modified
                if modified_text != original_text:
                    run.text = modified_text

        def replace_text_in_tables():
            nonlocal doc, flat_data_map, text_map  # Access variables from outer scope
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            replace_text_in_paragraph(para)
            
            # Table processing completed silently

        def process_entire_document():
            """Process the entire document comprehensively to catch all placeholders"""
            nonlocal doc, flat_data_map, text_map  # Access variables from outer scope
            # current_app.logger.info("🔄 PROCESSING ENTIRE DOCUMENT COMPREHENSIVELY")
            
            # Find ALL placeholders in the entire document first
            # current_app.logger.info("🔍 SEARCHING FOR ALL PLACEHOLDERS IN ENTIRE DOCUMENT")
            
            all_placeholders_found = set()
            
            # Search through ALL paragraphs, tables, headers, footers, etc.
            def search_for_placeholders(container):
                """Search for placeholders in any container (document, table, header, footer)"""
                if hasattr(container, 'paragraphs'):
                    for para in container.paragraphs:
                        if para.text:
                            # Searching paragraph for placeholders
                            # Find ${} placeholders
                            dollar_matches = re.findall(r"\$\{(.*?)\}", para.text)
                            for match in dollar_matches:
                                all_placeholders_found.add(f"${{{match}}}")
                                # Found $ placeholder
                            
                            # Find <> placeholders
                            angle_matches = re.findall(r"<(.*?)>", para.text)
                            for match in angle_matches:
                                all_placeholders_found.add(f"<{match}>")
                                # Found <> placeholder
                
                if hasattr(container, 'tables'):
                    for table in container.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                search_for_placeholders(cell)
            
            # Search in main document
            search_for_placeholders(doc)
            
            # Search in headers and footers
            for section in doc.sections:
                if section.header:
                    search_for_placeholders(section.header)
                if section.footer:
                    search_for_placeholders(section.footer)
            
            # Additional search: Look at raw XML for any missed placeholders
            # current_app.logger.info("🔍 ADDITIONAL SEARCH: Looking at raw XML for missed placeholders")
            try:
                for element in doc.element.iter():
                    if hasattr(element, 'text') and element.text:
                        # Find ${} placeholders
                        dollar_matches = re.findall(r"\$\{(.*?)\}", element.text)
                        for match in dollar_matches:
                            all_placeholders_found.add(f"${{{match}}}")
                            # Found $ placeholder in XML
                        
                        # Find <> placeholders
                        angle_matches = re.findall(r"<(.*?)>", element.text)
                        for match in angle_matches:
                            all_placeholders_found.add(f"<{match}>")
                            # Found <> placeholder in XML
            except Exception as e:
                current_app.logger.warning(f"⚠️ Error in additional XML search: {e}")
            
            # current_app.logger.info(f"🔍 Found {len(all_placeholders_found)} unique placeholders: {list(all_placeholders_found)}")
            
            # Log specific global metadata placeholders found
            global_placeholders_found = [p for p in all_placeholders_found if any(key in p.lower() for key in ['country', 'report_name', 'report_code', 'currency'])]
            # current_app.logger.info(f"🔍 Global metadata placeholders found: {global_placeholders_found}")
            
            # Verify that we have data for all found global metadata placeholders
            for placeholder in global_placeholders_found:
                if placeholder.startswith('${'):
                    key = placeholder[2:-1].lower()
                elif placeholder.startswith('<') and placeholder.endswith('>'):
                    key = placeholder[1:-1].lower()
                else:
                    continue
                
                if key in ['country', 'report_name', 'report_code', 'currency']:
                      if key in flat_data_map:
                          current_app.logger.info(f"✅ Data available for {placeholder}: {flat_data_map[key]}")
                      else:
                          current_app.logger.error(f"❌ NO DATA AVAILABLE for {placeholder} (key: {key})")
                          current_app.logger.error(f"❌ Available keys: {list(flat_data_map.keys())}")
            
            # Process Table of Contents specifically - Direct XML string replacement
            # This MUST happen BEFORE the main content replacement to preserve tags in XML
            try:
                # Processing TOC entries before main replacement
                
                # Save current document to temporary file BEFORE any replacements
                import tempfile
                import zipfile
                import shutil
                import os
                
                tmp_path = tempfile.mktemp(suffix='.docx')
                doc.save(tmp_path)
                
                # Extract the document as ZIP
                extract_dir = tempfile.mkdtemp()
                with zipfile.ZipFile(tmp_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)
                
                # COMPREHENSIVE HYPERLINK-AWARE XML PROCESSING
                current_app.logger.info("🔄 COMPREHENSIVE HYPERLINK-AWARE XML PROCESSING...")
                
                # Track all modifications
                total_files_modified = 0
                total_replacements = 0
                
                # 1. Process ALL XML files in the document (including hyperlink files)
                # Scanning all XML files for <country> tags
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        if file.endswith('.xml'):
                            file_path = os.path.join(root, file)
                            try:
                                with open(file_path, 'r', encoding='utf-8') as f:
                                    content = f.read()
                                
                                if '<country>' in content:
                                    country_count = content.count('<country>')
                                    current_app.logger.info(f"🔄 FOUND {country_count} <country> TAGS IN {file}")
                                    
                                    # Replace <country> with Austria
                                    modified_content = content.replace('<country>', 'Austria')
                                    
                                    if modified_content != content:
                                        with open(file_path, 'w', encoding='utf-8') as f:
                                            f.write(modified_content)
                                        
                                        total_files_modified += 1
                                        total_replacements += country_count
                                        current_app.logger.info(f"🔄 XML FILE MODIFIED: {file} ({country_count} replacements)")
                                    else:
                                        current_app.logger.warning(f"⚠️ NO CHANGES MADE TO {file} DESPITE FINDING TAGS")
                            except Exception as e:
                                current_app.logger.warning(f"⚠️ Error processing {file}: {e}")
                
                # 2. Special processing for hyperlink-specific files
                # Special processing for hyperlink files
                
                # Look for files that might contain hyperlink data
                hyperlink_keywords = ['hyperlink', 'link', 'toc', 'table', 'contents', 'rels', 'relationship']
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        if file.endswith('.xml') and any(keyword in file.lower() for keyword in hyperlink_keywords):
                            file_path = os.path.join(root, file)
                            try:
                                with open(file_path, 'r', encoding='utf-8') as f:
                                    content = f.read()
                                
                                if '<country>' in content:
                                    country_count = content.count('<country>')
                                    current_app.logger.info(f"🔄 FOUND {country_count} <country> TAGS IN HYPERLINK FILE: {file}")
                                    
                                    # Replace <country> with Austria
                                    modified_content = content.replace('<country>', 'Austria')
                                    
                                    if modified_content != content:
                                        with open(file_path, 'w', encoding='utf-8') as f:
                                            f.write(modified_content)
                                        
                                        total_files_modified += 1
                                        total_replacements += country_count
                                        current_app.logger.info(f"🔄 HYPERLINK FILE MODIFIED: {file} ({country_count} replacements)")
                            except Exception as e:
                                current_app.logger.warning(f"⚠️ Error processing hyperlink file {file}: {e}")
                
                # 3. Process _rels files (relationship files that might contain hyperlink data)
                # Processing relationship files
                rels_dir = os.path.join(extract_dir, '_rels')
                if os.path.exists(rels_dir):
                    for file in os.listdir(rels_dir):
                        if file.endswith('.xml'):
                            file_path = os.path.join(rels_dir, file)
                            try:
                                with open(file_path, 'r', encoding='utf-8') as f:
                                    content = f.read()
                                
                                if '<country>' in content:
                                    country_count = content.count('<country>')
                                    current_app.logger.info(f"🔄 FOUND {country_count} <country> TAGS IN RELS FILE: {file}")
                                    
                                    # Replace <country> with Austria
                                    modified_content = content.replace('<country>', 'Austria')
                                    
                                    if modified_content != content:
                                        with open(file_path, 'w', encoding='utf-8') as f:
                                            f.write(modified_content)
                                        
                                        total_files_modified += 1
                                        total_replacements += country_count
                                        current_app.logger.info(f"🔄 RELS FILE MODIFIED: {file} ({country_count} replacements)")
                            except Exception as e:
                                current_app.logger.warning(f"⚠️ Error processing rels file {file}: {e}")
                
                # 4. Process word/_rels files (Word-specific relationship files)
                word_rels_dir = os.path.join(extract_dir, 'word', '_rels')
                if os.path.exists(word_rels_dir):
                    for file in os.listdir(word_rels_dir):
                        if file.endswith('.xml'):
                            file_path = os.path.join(word_rels_dir, file)
                            try:
                                with open(file_path, 'r', encoding='utf-8') as f:
                                    content = f.read()
                                
                                if '<country>' in content:
                                    country_count = content.count('<country>')
                                    current_app.logger.info(f"🔄 FOUND {country_count} <country> TAGS IN WORD_RELS FILE: {file}")
                                    
                                    # Replace <country> with Austria
                                    modified_content = content.replace('<country>', 'Austria')
                                    
                                    if modified_content != content:
                                        with open(file_path, 'w', encoding='utf-8') as f:
                                            f.write(modified_content)
                                        
                                        total_files_modified += 1
                                        total_replacements += country_count
                                        current_app.logger.info(f"🔄 WORD_RELS FILE MODIFIED: {file} ({country_count} replacements)")
                            except Exception as e:
                                current_app.logger.warning(f"⚠️ Error processing word_rels file {file}: {e}")
                
                # 5. If any files were modified, recreate the document
                if total_files_modified > 0:
                    current_app.logger.info(f"🔄 COMPREHENSIVE XML REPLACEMENT COMPLETED: {total_files_modified} files, {total_replacements} total replacements")
                    
                    # Recreate the document
                    with zipfile.ZipFile(tmp_path, 'w') as zip_ref:
                        for root, dirs, files in os.walk(extract_dir):
                            for file in files:
                                file_path = os.path.join(root, file)
                                arcname = os.path.relpath(file_path, extract_dir)
                                zip_ref.write(file_path, arcname)
                    
                    # Reload the modified document
                    from docx import Document as NewDocument
                    doc = NewDocument(tmp_path)
                    current_app.logger.info("🔄 DOCUMENT RELOADED AFTER COMPREHENSIVE XML MODIFICATION")
                    
                    # Cleanup and return - no need for further processing
                    shutil.rmtree(extract_dir)
                    os.unlink(tmp_path)
                    return
                else:
                    # No XML files were modified
                    
                    # Cleanup
                    shutil.rmtree(extract_dir)
                    os.unlink(tmp_path)
                

                    
            except Exception as e:
                current_app.logger.warning(f"⚠️ Error processing TOC: {e}")
            
            # Now replace ALL placeholders everywhere they appear
            # Replacing all placeholders everywhere
            
            # Single comprehensive pass to avoid duplication
            # Processing all document elements in single pass
                
                # Process main document
            for para in doc.paragraphs:
                    replace_text_in_paragraph(para)
                
                # Process tables
            for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                replace_text_in_paragraph(para)
                
                # Process headers and footers
            for section in doc.sections:
                    if section.header:
                        for para in section.header.paragraphs:
                            replace_text_in_paragraph(para)
                        for table in section.header.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    for para in cell.paragraphs:
                                        replace_text_in_paragraph(para)
                    
                    if section.footer:
                        for para in section.footer.paragraphs:
                            replace_text_in_paragraph(para)
                        for table in section.footer.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    for para in cell.paragraphs:
                                        replace_text_in_paragraph(para)
                
            # XML processing removed to prevent duplication - paragraph processing is sufficient
            
            #current_app.logger.info("✅ COMPREHENSIVE DOCUMENT PROCESSING COMPLETED")
            
            # Additional pass: Handle special Word elements that might contain placeholders
                    # Final pass: Processing special Word elements
            
            # Process text boxes and other special elements (only once)
            try:
                for shape in doc.inline_shapes:
                    if hasattr(shape, 'text_frame'):
                        for para in shape.text_frame.paragraphs:
                            replace_text_in_paragraph(para)
            except Exception as e:
                current_app.logger.warning(f"⚠️ Error processing inline shapes: {e}")
            
            # Process Word fields (like table of contents) that might contain placeholders
            try:
                for field in doc.fields:
                    if hasattr(field, 'text') and field.text:
                        original_text = field.text
                        modified_text = original_text
                        
                        # Replace ${} placeholders
                        for key, value in flat_data_map.items():
                            if key in ['country', 'report_name', 'report_code', 'currency']:
                                dollar_placeholder = f"${{{key}}}"
                                if dollar_placeholder in modified_text:
                                    modified_text = modified_text.replace(dollar_placeholder, str(value))
                                    current_app.logger.info(f"🔄 FIELD REPLACED: {dollar_placeholder} -> {value}")
                                
                        # Replace <> placeholders
                        for key, value in flat_data_map.items():
                            if key in ['country', 'report_name', 'report_code', 'currency']:
                                angle_placeholder = f"<{key}>"
                                if angle_placeholder in modified_text:
                                    modified_text = modified_text.replace(angle_placeholder, str(value))
                                    current_app.logger.info(f"🔄 FIELD REPLACED: {angle_placeholder} -> {value}")
                        
                        # Update field text if modified
                        if modified_text != original_text:
                            field.text = modified_text
            except Exception as e:
                # Word fields processing skipped
                pass
            

            
            # Process all XML elements for any remaining placeholders - Careful approach to avoid duplication
            try:
                # Processing XML elements
                xml_replacements = 0
                
                for element in doc.element.iter():
                    if hasattr(element, 'text') and element.text and '<country>' in element.text:
                        original_text = element.text
                        
                        # Only replace if it's a simple text element to avoid duplication
                        if hasattr(element, 'tag') and element.tag in ['w:t', 'w:tab', 'w:br']:
                            try:
                                element.text = original_text.replace('<country>', 'Austria')
                                xml_replacements += 1
                                current_app.logger.info(f"🔄 XML TEXT ELEMENT REPLACED: <country> -> Austria")
                            except Exception as xml_error:
                                current_app.logger.warning(f"⚠️ Could not update XML text element: {xml_error}")
                
                # XML processing complete
                
            except Exception as e:
                current_app.logger.warning(f"⚠️ Error processing XML elements: {e}")
            
            # Force update Table of Contents by refreshing the document
            try:
                # Update all TOC fields
                for field in doc.fields:
                    if field.type == 3:  # TOC field type
                        field.update()
                        current_app.logger.info("🔄 TOC FIELD UPDATED")
            except Exception as e:
                # TOC fields update skipped
                pass
            
            # Final verification: Check if any <country> tags remain
            try:
                # Final verification
                remaining_country_tags = 0
                
                # Check all paragraphs
                for paragraph in doc.paragraphs:
                    if '<country>' in paragraph.text:
                        remaining_country_tags += 1
                        #current_app.logger.warning(f"⚠️ REMAINING <country> TAG FOUND: {paragraph.text[:100]}...")
                
                # Check all tables
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if '<country>' in cell.text:
                                remaining_country_tags += 1
                                #current_app.logger.warning(f"⚠️ REMAINING <country> TAG IN TABLE: {cell.text[:100]}...")
                
                if remaining_country_tags == 0:
                    current_app.logger.info("✅ ALL <country> TAGS SUCCESSFULLY REPLACED!")
                else:
                    current_app.logger.warning(f"⚠️ {remaining_country_tags} <country> TAGS STILL REMAIN")
                    
            except Exception as e:
                current_app.logger.warning(f"⚠️ Error in final verification: {e}")
            
            # Search for any remaining placeholders that might have been missed
            def search_for_remaining_placeholders(element, path=""):
                """Recursively search for any remaining placeholders"""
                if hasattr(element, 'text') and element.text:
                    text = element.text
                    
                    # Check for any remaining ${} placeholders
                    dollar_matches = re.findall(r"\$\{(.*?)\}", text)
                    for match in dollar_matches:
                        key_lower = match.lower().strip()
                        if key_lower in ['country', 'report_name', 'report_code', 'currency']:
                            current_app.logger.error(f"❌ REMAINING PLACEHOLDER FOUND: ${{{match}}} in {path}")
                            current_app.logger.error(f"❌ Available data for {key_lower}: {flat_data_map.get(key_lower, 'NOT FOUND')}")
                    
                    # Check for any remaining <> placeholders
                    angle_matches = re.findall(r"<(.*?)>", text)
                    for match in angle_matches:
                        key_lower = match.lower().strip()
                        if key_lower in ['country', 'report_name', 'report_code', 'currency']:
                            current_app.logger.error(f"❌ REMAINING PLACEHOLDER FOUND: <{match}> in {path}")
                            current_app.logger.error(f"❌ Available data for {key_lower}: {flat_data_map.get(key_lower, 'NOT FOUND')}")
                
                # Recursively check child elements
                for i, child in enumerate(element):
                    child_path = f"{path}.{i}" if path else str(i)
                    search_for_remaining_placeholders(child, child_path)
            
            #current_app.logger.info("✅ FINAL VERIFICATION COMPLETED")


                            
        # Data mapping completed silently

        def generate_chart(data_dict, chart_tag):
            import plotly.graph_objects as go
            import matplotlib.pyplot as plt
            from openpyxl.utils import column_index_from_string
            import numpy as np
            import os
            import tempfile
            import json
            import re
            import warnings
            
            # Suppress Matplotlib warnings
            warnings.filterwarnings('ignore', category=UserWarning, module='matplotlib')

            try:
                chart_tag_lower = chart_tag.lower()
                raw_chart_attr = chart_attr_map.get(chart_tag_lower, "{}")
                chart_config = json.loads(re.sub(r'//.*?\n|/\*.*?\*/', '', raw_chart_attr, flags=re.DOTALL))

                chart_meta = chart_config.get("chart_meta", {})
                series_meta = chart_config.get("series", {})
                chart_type = chart_type_map.get(chart_tag_lower, "").lower().strip()
                title = chart_meta.get("chart_title", chart_tag)

                # --- Comprehensive attribute detection logging ---
                #current_app.logger.info(f"🔍 COMPREHENSIVE CHART ATTRIBUTE DETECTION STARTED")
                #current_app.logger.info(f"📊 Chart Type: {chart_type}")
                #current_app.logger.info(f"📋 Chart Meta Keys Found: {list(chart_meta.keys())}")
                #current_app.logger.info(f"📋 Chart Config Keys Found: {list(chart_config.keys())}")
                #current_app.logger.info(f"📋 Top-level Keys Found: {list(data_dict.keys())}")
                
                # Define all possible chart attributes
                all_possible_attributes = [
                    "chart_title", "font_size", "font_color", "font_family", "figsize", 
                    "chart_background", "plot_background", "legend", "legend_position", "legend_font_size",
                    "primary_y_label", "secondary_y_label", "x_label", "y_axis_min_max", 
                    "secondary_y_axis_format", "secondary_y_axis_min_max", "x_axis_label_distance", 
                    "y_axis_label_distance", "axis_tick_format", "axis_tick_font_size",
                    "data_labels", "data_label_format", "data_label_font_size", "data_label_color",
                    "show_gridlines", "gridline_color", "gridline_style", "margin", "bar_width", 
                    "orientation", "bar_border_color", "bar_border_width", "barmode", "line_width", 
                    "marker_size", "line_style", "fill_opacity", "hole", "startangle", "pull", 
                    "bins", "box_points", "violin_points", "bubble_size", "waterfall_measure", 
                    "funnel_measure", "sunburst_path", "treemap_path", "sankey_source", 
                    "sankey_target", "sankey_value", "table_header", "indicator_mode", 
                    "indicator_delta", "indicator_gauge", "3d_projection", "z_values", 
                    "lat", "lon", "locations", "open_values", "high_values", "low_values", 
                    "close_values", "annotations"
                ]
                
                # Check which attributes are missing (not in chart_meta OR chart_config OR top-level)
                missing_attributes = []
                for attr in all_possible_attributes:
                    if (attr not in chart_meta and 
                        attr not in chart_config and 
                        attr not in data_dict):
                        missing_attributes.append(attr)
                
                # Chart attribute detection completed (logging removed for cleaner output)

                # --- Define chart type mappings for Matplotlib ---
                chart_type_mapping_mpl = {
                    # Bar charts
                    "bar": "bar",
                    "column": "bar", 
                    "stacked_column": "bar",
                    "horizontal_bar": "barh",
                    
                    # Line charts
                    "line": "plot",
                    "scatter": "scatter",
                    "scatter_line": "plot",
                    
                    # Area charts
                    "area": "fill_between",
                    "filled_area": "fill_between",
                    
                    # Statistical charts
                    "histogram": "hist",
                    "box": "boxplot",
                    "violin": "violinplot",
                    
                    # Other charts
                    "bubble": "scatter",
                    "heatmap": "imshow",
                    "contour": "contour",
                    "waterfall": "bar",
                    "funnel": "bar",
                    "sunburst": "pie",
                    "treemap": "bar",
                    "icicle": "bar",
                    "sankey": "bar",
                    "table": "table",
                    "indicator": "bar"
                }

                # --- Extract custom fields from chart_config ---
                bar_colors = chart_config.get("bar_colors")
                bar_width = data_dict.get("bar_width") or chart_config.get("bar_width") or chart_meta.get("bar_width")
                orientation = data_dict.get("orientation") or chart_config.get("orientation") or chart_meta.get("orientation")
                bar_border_color = data_dict.get("bar_border_color") or chart_config.get("bar_border_color") or chart_meta.get("bar_border_color")
                bar_border_width = data_dict.get("bar_border_width") or chart_config.get("bar_border_width") or chart_meta.get("bar_border_width")
                font_family = data_dict.get("font_family") or chart_config.get("font_family") or chart_meta.get("font_family")
                # Add font fallback for macOS compatibility
                if font_family:
                    # Check if the font is available, otherwise use a fallback
                    import matplotlib.font_manager as fm
                    available_fonts = [f.name for f in fm.fontManager.ttflist]
                    if font_family not in available_fonts:
                        # Use system-appropriate fallback fonts
                        if font_family.lower() in ['calibri', 'arial']:
                            font_family = 'Helvetica'  # macOS equivalent
                        elif font_family.lower() in ['times new roman', 'times']:
                            font_family = 'Times'  # macOS equivalent
                        else:
                            font_family = 'Helvetica'  # Default fallback
                font_size = data_dict.get("font_size") or chart_config.get("font_size") or chart_meta.get("font_size")
                font_color = data_dict.get("font_color") or chart_config.get("font_color") or chart_meta.get("font_color")
                legend_position = data_dict.get("legend_position") or chart_config.get("legend_position") or chart_meta.get("legend_position")
                legend_font_size = data_dict.get("legend_font_size") or chart_config.get("legend_font_size") or chart_meta.get("legend_font_size")
                show_gridlines = data_dict.get("show_gridlines") if "show_gridlines" in data_dict else (chart_config.get("show_gridlines") if "show_gridlines" in chart_config else chart_meta.get("show_gridlines"))
                # Ensure show_gridlines is a boolean
                if isinstance(show_gridlines, str):
                    show_gridlines = show_gridlines.strip().lower() == "true"
                elif show_gridlines is None:
                    show_gridlines = True  # Default to showing gridlines if not specified
                gridline_color = data_dict.get("gridline_color") or chart_config.get("gridline_color") or chart_meta.get("gridline_color")
                gridline_style = data_dict.get("gridline_style") or chart_config.get("gridline_style") or chart_meta.get("gridline_style")
                chart_background = data_dict.get("chart_background") or chart_config.get("chart_background") or chart_meta.get("chart_background")
                plot_background = data_dict.get("plot_background") or chart_config.get("plot_background") or chart_meta.get("plot_background")
                data_label_format = data_dict.get("data_label_format") or chart_config.get("data_label_format") or chart_meta.get("data_label_format")
                data_label_font_size = data_dict.get("data_label_font_size") or chart_config.get("data_label_font_size") or chart_meta.get("data_label_font_size")
                data_label_color = data_dict.get("data_label_color") or chart_config.get("data_label_color") or chart_meta.get("data_label_color")
                axis_tick_format = data_dict.get("axis_tick_format") or chart_config.get("axis_tick_format") or chart_meta.get("axis_tick_format")
                y_axis_min_max = data_dict.get("y_axis_min_max") or chart_config.get("y_axis_min_max") or chart_meta.get("y_axis_min_max")
                # current_app.logger.debug(f"Y-axis min/max from config: {y_axis_min_max}")
                x_axis_min_max = data_dict.get("x_axis_min_max") or chart_config.get("x_axis_min_max") or chart_meta.get("x_axis_min_max")
                # current_app.logger.debug(f"X-axis min/max from config: {x_axis_min_max}")
                secondary_y_axis_format = data_dict.get("secondary_y_axis_format") or chart_config.get("secondary_y_axis_format") or chart_meta.get("secondary_y_axis_format")
                secondary_y_axis_min_max = data_dict.get("secondary_y_axis_min_max") or chart_config.get("secondary_y_axis_min_max") or chart_meta.get("secondary_y_axis_min_max")
                disable_secondary_y = data_dict.get("disable_secondary_y") or chart_config.get("disable_secondary_y") or chart_meta.get("disable_secondary_y", False)
                # current_app.logger.info(f"🔧 disable_secondary_y setting: {disable_secondary_y}")
                sort_order = data_dict.get("sort_order") or chart_config.get("sort_order") or chart_meta.get("sort_order")
                data_grouping = data_dict.get("data_grouping") or chart_config.get("data_grouping") or chart_meta.get("data_grouping")
                annotations = data_dict.get("annotations", []) or chart_config.get("annotations", []) or chart_meta.get("annotations", [])
                axis_tick_font_size = data_dict.get("axis_tick_font_size") or chart_config.get("axis_tick_font_size") or chart_meta.get("axis_tick_font_size")
                
                # --- Extract tick mark control settings ---
                show_x_ticks = data_dict.get("show_x_ticks") if "show_x_ticks" in data_dict else (chart_config.get("show_x_ticks") if "show_x_ticks" in chart_config else chart_meta.get("show_x_ticks"))
                show_y_ticks = data_dict.get("show_y_ticks") if "show_y_ticks" in data_dict else (chart_config.get("show_y_ticks") if "show_y_ticks" in chart_config else chart_meta.get("show_y_ticks"))
                # Ensure tick settings are boolean
                if isinstance(show_x_ticks, str):
                    show_x_ticks = show_x_ticks.strip().lower() == "true"
                elif show_x_ticks is None:
                    show_x_ticks = True  # Default to showing ticks if not specified
                if isinstance(show_y_ticks, str):
                    show_y_ticks = show_y_ticks.strip().lower() == "true"
                elif show_y_ticks is None:
                    show_y_ticks = True  # Default to showing ticks if not specified
                
                # --- Extract margin settings ---
                margin = data_dict.get("margin") or chart_config.get("margin") or chart_meta.get("margin")
                x_axis_label_distance = data_dict.get("x_axis_label_distance") or chart_config.get("x_axis_label_distance") or chart_meta.get("x_axis_label_distance")
                y_axis_label_distance = (
                    data_dict.get("y_axis_label_distance")
                    or chart_config.get("y_axis_label_distance")
                    or chart_meta.get("y_axis_label_distance")
                    or chart_meta.get("primary_y_axis_label_distance")
                )
                axis_tick_distance = data_dict.get("axis_tick_distance") or chart_config.get("axis_tick_distance") or chart_meta.get("axis_tick_distance")
                figsize = data_dict.get("figsize") or chart_config.get("figsize") or chart_meta.get("figsize")
                
                # --- Extract additional missing attributes ---
                legend = data_dict.get("legend") or chart_config.get("legend") or chart_meta.get("legend")
                data_labels = data_dict.get("data_labels") or chart_config.get("data_labels") or chart_meta.get("data_labels")
                line_width = data_dict.get("line_width") or chart_config.get("line_width") or chart_meta.get("line_width")
                marker_size = data_dict.get("marker_size") or chart_config.get("marker_size") or chart_meta.get("marker_size")
                line_style = data_dict.get("line_style") or chart_config.get("line_style") or chart_meta.get("line_style")
                fill_opacity = data_dict.get("fill_opacity") or chart_config.get("fill_opacity") or chart_meta.get("fill_opacity")
                hole = data_dict.get("hole") or chart_config.get("hole") or chart_meta.get("hole")
                startangle = data_dict.get("startangle") or chart_config.get("startangle") or chart_meta.get("startangle")
                pull = data_dict.get("pull") or chart_config.get("pull") or chart_meta.get("pull")
                barmode = data_dict.get("barmode") or chart_config.get("barmode") or chart_meta.get("barmode")

                # --- Excel range extraction helpers ---
                def extract_excel_range(sheet, cell_range):
                    # cell_range: e.g., 'E23:E29' or 'AA20:AA23'
                    try:
                        values = []
                        for row in sheet[cell_range]:
                            for cell in row:
                                            if cell.value is not None:
                                                values.append(cell.value)
                        return values
                    except Exception as e:
                            current_app.logger.error(f"Error extracting range {cell_range}: {e}")
                            return []

                # --- Robust recursive Excel cell range extraction for all chart types and fields ---
                def extract_cell_ranges(obj, sheet):
                    """Recursively walk through nested dicts/lists and extract cell ranges from strings"""
                    if isinstance(obj, dict):
                        for k, v in obj.items():
                            if isinstance(v, str) and re.match(r"^[A-Z]+\d+:[A-Z]+\d+$", v):
                                try:
                                    extracted = extract_excel_range(sheet, v)
                                    obj[k] = extracted
                                except Exception as e:
                                    # Failed to extract data from cell range
                                    pass
                            else:
                                extract_cell_ranges(v, sheet)
                    elif isinstance(obj, list):
                        for i, v in enumerate(obj):
                            extract_cell_ranges(v, sheet)
                
                if "source_sheet" in chart_meta:
                    wb = openpyxl.load_workbook(data_file_path, data_only=True)
                    sheet = wb[chart_meta["source_sheet"]]
                    # Extract cell ranges from both chart_meta and series_meta
                    extract_cell_ranges(chart_meta, sheet)
                    extract_cell_ranges(series_meta, sheet)
                    wb.close()
                
                # Use updated values from series_meta after extraction
                series_data = series_meta.get("data", [])
                if not series_data and "series" in series_meta:
                    # Handle case where data is directly in series object
                    series_data = series_meta.get("series", [])
                
                # Extract x_values from the correct location
                x_values = series_meta.get("x_axis", [])
                if not x_values and "series" in series_meta:
                    # Try to get x_axis from the series object
                    x_values = series_meta.get("series", {}).get("x_axis", [])
                
                # If still no x_values, try to get from the first series data item
                if not x_values and series_data and len(series_data) > 0:
                    # Check if x_axis is in the first series
                    first_series = series_data[0]
                    if isinstance(first_series, dict) and "x_axis" in first_series:
                        x_values = first_series["x_axis"]
                    # Also check if there's a separate x_axis in the series object
                    elif "series" in series_meta and isinstance(series_meta["series"], dict):
                        x_values = series_meta["series"].get("x_axis", [])
                
                # current_app.logger.info(f"🔍 Extracted x_values: {x_values}")
                
                # Ensure x_values is always defined to prevent "cannot access local variable" error
                if not x_values:
                    x_values = []
                    # No x_values found, using empty list as fallback
                
                colors = series_meta.get("colors", [])
                
                # --- SERIES ATTRIBUTE DETECTION LOGGING ---
                #current_app.logger.info(f"🔍 SERIES ATTRIBUTE DETECTION STARTED")
                #current_app.logger.info(f"📊 Number of series: {len(series_data)}")
                #current_app.logger.info(f"📋 Series meta keys: {list(series_data.keys()) if isinstance(series_data, dict) else 'N/A'}")
                #current_app.logger.info(f"📋 Series meta structure: {series_meta}")
                #current_app.logger.info(f"📋 Series data: {series_data}")
                #current_app.logger.info(f"📋 X values extracted: {x_values}")
                
                # Define all possible series attributes
                all_possible_series_attributes = [
                    "marker", "opacity", "textposition", "orientation", "width", "fill", 
                    "fillcolor", "hole", "pull", "mode", "line", "nbinsx", "boxpoints", 
                    "jitter", "sizeref", "sizemin", "symbol", "measure", "connector", "textinfo"
                ]
                
                for i, series in enumerate(series_data):
                    series_name = series.get("name", f"Series {i+1}")
                    series_type = series.get("type", "unknown")
                    #current_app.logger.info(f"📈 SERIES {i+1}: {series_name}")
                    #current_app.logger.info(f"   Type: {series_type}")
                    
                    # Check which series attributes are missing
                    missing_series_attributes = []
                    for attr in all_possible_series_attributes:
                        if attr not in series:
                            missing_series_attributes.append(attr)
                    
                    # Series attribute detection completed (logging removed for cleaner output)

                # --- Plotly interactive chart generation ---
                fig = go.Figure()

                # --- Bar of Pie chart special handling ---
                if chart_type in ["bar of pie", "bar_of_pie"]:
                    # Extract cell ranges for other_labels and other_values if they are cell ranges
                    other_labels = chart_meta.get("other_labels", [])
                    other_values = chart_meta.get("other_values", [])
                    other_colors = chart_meta.get("other_colors", [])
                    
                    # If other_labels and other_values are cell ranges, extract them
                    if isinstance(other_labels, str) and re.match(r"^[A-Z]+\d+:[A-Z]+\d+$", other_labels):
                        try:
                            wb = openpyxl.load_workbook(data_file_path, data_only=True)
                            sheet = wb[chart_meta.get("source_sheet", "sample")]
                            other_labels = extract_excel_range(sheet, other_labels)
                            wb.close()
                            # Extracted other_labels successfully
                            pass
                        except Exception as e:
                            # Failed to extract other_labels
                            pass
                    
                    if isinstance(other_values, str) and re.match(r"^[A-Z]+\d+:[A-Z]+\d+$", other_values):
                        try:
                            wb = openpyxl.load_workbook(data_file_path, data_only=True)
                            sheet = wb[chart_meta.get("source_sheet", "sample")]
                            other_values = extract_excel_range(sheet, other_values)
                            wb.close()
                            # Extracted other_values successfully
                            print(f"DEBUG: other_values extracted from {other_values} = {other_values}")
                            pass
                        except Exception as e:
                            # Failed to extract other_values
                            pass
                    
                    # Log the extracted data for debugging
                    #current_app.logger.info(f"🔍 Bar of Pie Chart Data:")
                    #current_app.logger.info(f"   Other Labels: {other_labels}")
                    #current_app.logger.info(f"   Other Values: {other_values}")
                    #current_app.logger.info(f"   Other Colors: {other_colors}")
                    
                    # Check if values are percentages in decimal form and convert them
                    y_axis_title = chart_meta.get("y_axis_title", "")
                    if other_values and y_axis_title and "%" in y_axis_title:
                        # Check if all values are between 0-1 (likely percentages in decimal form)
                        if all(isinstance(v, (int, float)) and 0 <= v <= 1 for v in other_values if v is not None):
                            print(f"DEBUG: Converting decimal values to percentages: {other_values}")
                            other_values = [v * 100 if v is not None else v for v in other_values]
                            print(f"DEBUG: Converted to: {other_values}")
                    
                    # Check for empty/null values
                    if other_labels and other_values:
                        empty_count = sum(1 for label, value in zip(other_labels, other_values) 
                                        if (label is None or str(label).strip() == "" or 
                                            value is None or str(value).strip() == "" or 
                                            str(value).strip() == "0"))
                        if empty_count > 0:
                            # Found empty/null values in bar chart data (will cause missing bars)
                            pass
                        
                        # Data analysis completed (logging removed for cleaner output)
                    
                    value_format = chart_meta.get("value_format", "")
                    labels = series_meta.get("labels", x_values)
                    values = series_meta.get("values", [])
                    colors = series_meta.get("colors", [])
                    
                    # Chart data prepared
                    fig = create_bar_of_pie_chart(
                        labels=labels,
                        values=values,
                        other_labels=other_labels,
                        other_values=other_values,
                        colors=colors,
                        other_colors=other_colors if other_colors else colors,
                        title=title,
                        value_format=value_format
                    )
                
                def extract_values_from_range(cell_range):
                    start_cell, end_cell = cell_range.split(":")
                    start_col, start_row = re.match(r"([A-Z]+)(\d+)", start_cell).groups()
                    end_col, end_row = re.match(r"([A-Z]+)(\d+)", end_cell).groups()

                    start_col_idx = column_index_from_string(start_col) - 1
                    end_col_idx = column_index_from_string(end_col) - 1
                    start_row_idx = int(start_row) - 1  # Fixed: pandas is 0-indexed, Excel is 1-indexed
                    end_row_idx = int(end_row) - 1      # Fixed: pandas is 0-indexed, Excel is 1-indexed

                    if start_col_idx == end_col_idx:
                        return df.iloc[start_row_idx:end_row_idx + 1, start_col_idx].tolist()
                    else:
                        return df.iloc[start_row_idx:end_row_idx + 1, start_col_idx:end_col_idx + 1].values.flatten().tolist()

                # --- Data grouping and sorting logic ---
                # If data_grouping is present, filter x/y values to only those groups
                def group_and_sort(x_vals, y_vals, group_names=None, sort_order=None):
                    # Grouping: only keep x/y where x in group_names (if group_names provided)
                    if group_names:
                        filtered = [(x, y) for x, y in zip(x_vals, y_vals) if x in group_names]
                        if filtered:
                            x_vals, y_vals = zip(*filtered)
                        else:
                            x_vals, y_vals = [], []
                    # Sorting: sort by y-value
                    if sort_order == "ascending":
                        sorted_pairs = sorted(zip(x_vals, y_vals), key=lambda pair: pair[1])
                        if sorted_pairs:
                            x_vals, y_vals = zip(*sorted_pairs)
                    elif sort_order == "descending":
                        sorted_pairs = sorted(zip(x_vals, y_vals), key=lambda pair: pair[1], reverse=True)
                        if sorted_pairs:
                            x_vals, y_vals = zip(*sorted_pairs)
                    return list(x_vals), list(y_vals)

                # Special handling for pie charts (single trace)
                if chart_type == "pie" and len(series_data) == 1:
                    series = series_data[0]
                    label = series.get("name", "Pie Chart")
                    labels = series.get("labels", x_values)
                    values = series.get("values", [])
                    color = series.get("marker", {}).get("color") if "marker" in series else colors
                    
                    # Apply value format if specified
                    value_format = chart_meta.get("value_format", "")
                    if value_format:
                        # Convert values to percentage format for display
                        try:
                            formatted_values = [f"{float(v):{value_format}}" if v is not None else "0" for v in values]
                        except:
                            formatted_values = values
                    else:
                        formatted_values = values
                    
                    # Check if this is an expanded pie chart (pie with one segment expanded to column)
                    expanded_segment = chart_meta.get("expanded_segment")
                    
                    if expanded_segment:
                        # Use the helper function for expanded pie charts
                        fig = create_expanded_pie_chart(
                            labels=labels,
                            values=values,
                            colors=color if isinstance(color, list) else [color],
                            expanded_segment=expanded_segment,
                            title=title,
                            value_format=value_format
                        )
                        
                    else:
                        # Regular pie chart
                        pie_kwargs = {
                            "labels": labels,
                            "values": values,
                            "name": label,
                            "textinfo": "label+percent+value" if chart_meta.get("data_labels", True) else "none",
                            "textposition": "outside",
                            "hole": 0.0  # Solid pie chart
                        }
                        
                        # Pie colors
                        if color:
                            pie_kwargs["marker"] = dict(colors=color) if isinstance(color, list) else dict(colors=[color])
                        
                        # Add pull effect for specific segments if needed
                        if "pull" in chart_meta:
                            pie_kwargs["pull"] = chart_meta["pull"]
                        
                        fig.add_trace(go.Pie(**pie_kwargs,
                            hovertemplate=f"<b>{label}</b><br>%{{label}}: %{{value}}{str(value_format) if value_format else ''}<extra></extra>"
                        ))
                
                # Handle stacked column, area, and other multi-series charts
                else:
                    for i, series in enumerate(series_data):
                        label = series.get("name", f"Series {i+1}")
                        series_type = series.get("type", "bar").lower()
                        
                        # Map heatmap to imshow for Matplotlib
                        if series_type == "heatmap":
                            mpl_chart_type = "imshow"
                        else:
                            mpl_chart_type = chart_type_mapping_mpl.get(series_type, "scatter")
                        
                        color = None
                        if "marker" in series and isinstance(series["marker"], dict) and "color" in series["marker"]:
                            color = series["marker"]["color"]
                        elif bar_colors:
                            color = bar_colors
                        elif i < len(colors):
                            color = colors[i]

                        y_vals = series.get("values")
                        value_range = series.get("value_range")
                        if value_range:
                            # Check if value_range is already extracted (list) or still a string
                            if isinstance(value_range, list):
                                y_vals = value_range
                            else:
                                y_vals = extract_values_from_range(value_range)

                        # --- Apply grouping and sorting ---
                        x_vals = x_values
                        if y_vals is not None and x_vals is not None:
                            x_vals, y_vals = group_and_sort(x_vals, y_vals, data_grouping, sort_order)

                        trace_kwargs = {
                            "x": x_vals,
                            "y": y_vals,
                            "name": label,
                        }
                        
                        # Handle colors
                        if color:
                            if isinstance(color, list) and series_type == "bar":
                                trace_kwargs["marker"] = dict(color=color)
                            elif isinstance(color, str):
                                trace_kwargs["marker_color"] = color
                        
                        # Handle bar-specific properties
                        if bar_width and series_type == "bar":
                            trace_kwargs["width"] = bar_width
                        if orientation and series_type == "bar":
                            trace_kwargs["orientation"] = orientation[0].lower() if isinstance(orientation, str) else orientation
                        if bar_border_color and series_type == "bar":
                            trace_kwargs["marker_line_color"] = bar_border_color
                        if bar_border_width and series_type == "bar":
                            trace_kwargs["marker_line_width"] = bar_border_width
                        
                        # Handle line-specific properties
                        if line_width and series_type in ["line", "scatter", "scatter_line"]:
                            trace_kwargs["line_width"] = line_width
                        if marker_size and series_type in ["line", "scatter", "scatter_line"]:
                            trace_kwargs["marker_size"] = marker_size
                        if line_style and series_type in ["line", "scatter", "scatter_line"]:
                            # Map line styles to Plotly format
                            line_style_map = {
                                "solid": "solid",
                                "dashed": "dash",
                                "dotted": "dot",
                                "dashdot": "dashdot"
                            }
                            trace_kwargs["line_dash"] = line_style_map.get(line_style, "solid")
                        
                        # Handle opacity
                        if fill_opacity:
                            trace_kwargs["opacity"] = fill_opacity

                        # Add traces based on chart type - REPLACE THE RESTRICTIVE IF/ELIF BLOCKS
                        # Generic chart type handling for Plotly
                        chart_type_mapping = {
                            # Bar charts
                            "bar": go.Bar,
                            "column": go.Bar,
                            "stacked_column": go.Bar,
                            "horizontal_bar": go.Bar,
                            
                            # Line charts
                            "line": go.Scatter,
                            "scatter": go.Scatter,
                            "scatter_line": go.Scatter,
                            
                            # Area charts
                            "area": go.Scatter,
                            "filled_area": go.Scatter,
                            
                            # Pie charts
                            "pie": go.Pie,
                            "donut": go.Pie,
                            
                            # 3D charts
                            "scatter3d": go.Scatter3d,
                            "surface": go.Surface,
                            "mesh3d": go.Mesh3d,
                            
                            # Statistical charts
                            "histogram": go.Histogram,
                            "box": go.Box,
                            "violin": go.Violin,
                            
                            # Financial charts
                            "candlestick": go.Candlestick,
                            "ohlc": go.Ohlc,
                            
                            # Geographic charts
                            "scattergeo": go.Scattergeo,
                            "choropleth": go.Choropleth,
                            
                            # Other charts
                            "bubble": go.Scatter,
                            "heatmap": go.Heatmap,
                            "contour": go.Contour,
                            "waterfall": go.Waterfall,
                            "funnel": go.Funnel,
                            "sunburst": go.Sunburst,
                            "treemap": go.Treemap,
                            "icicle": go.Icicle,
                            "sankey": go.Sankey,
                            "table": go.Table,
                            "indicator": go.Indicator
                        }
                        
                        # Get the appropriate Plotly chart class
                        plotly_chart_class = chart_type_mapping.get(series_type)
                        
                        if plotly_chart_class:
                            # Prepare trace arguments based on chart type
                            if series_type in ["bar", "column", "stacked_column", "horizontal_bar"]:
                                # Bar chart specific settings
                                # REMOVE: if chart_type == "stacked_column": trace_kwargs["barmode"] = "stack"
                                if orientation and orientation.lower() == "horizontal":
                                    trace_kwargs["orientation"] = "h"
                                    # Swap x and y for horizontal bars
                                    trace_kwargs["x"], trace_kwargs["y"] = trace_kwargs["y"], trace_kwargs["x"]
                                
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>Category: %{{x}}<br>Value: %{{y}}<extra></extra>"
                                ))
                                
                            elif series_type in ["line", "scatter", "scatter_line"]:
                                # Line/Scatter chart specific settings
                                mode = "lines+markers" if series_type == "scatter_line" else "markers" if series_type == "scatter" else "lines"
                                trace_kwargs["mode"] = mode
                                
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>Category: %{{x}}<br>Value: %{{y}}<extra></extra>"
                                ))
                                
                            elif series_type in ["area", "filled_area"]:
                                # Area chart specific settings
                                trace_kwargs["mode"] = "lines"
                                trace_kwargs["fill"] = "tozeroy"
                                
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>Category: %{{x}}<br>Value: %{{y}}<extra></extra>"
                                ))
                                
                            elif series_type == "pie":
                                # Pie chart specific settings
                                pie_kwargs = {
                                    "labels": x_vals,
                                    "values": y_vals,
                                    "name": label,
                                    "textinfo": "label+percent+value" if data_labels else "none",
                                    "textposition": "outside",
                                    "hole": hole if hole is not None else (0.4 if series_type == "donut" else 0.0)
                                }
                                
                                # Add pie chart specific attributes
                                if startangle is not None:
                                    pie_kwargs["rotation"] = startangle
                                if pull is not None:
                                    # pull can be a single value or a list
                                    if isinstance(pull, (int, float)):
                                        pie_kwargs["pull"] = [pull] * len(y_vals)
                                    elif isinstance(pull, list):
                                        pie_kwargs["pull"] = pull
                                
                                if color:
                                    pie_kwargs["marker"] = dict(colors=color) if isinstance(color, list) else dict(colors=[color])
                                
                                fig.add_trace(plotly_chart_class(**pie_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>%{{label}}: %{{value}}<extra></extra>"
                                ))
                                
                            elif series_type in ["scatter3d", "surface", "mesh3d"]:
                                # 3D chart specific settings
                                if "z" not in trace_kwargs and len(y_vals) > 0:
                                    # Create a simple z-axis if not provided
                                    trace_kwargs["z"] = [i for i in range(len(y_vals))]
                                
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>X: %{{x}}<br>Y: %{{y}}<br>Z: %{{z}}<extra></extra>"
                                ))
                                
                            elif series_type in ["histogram", "box", "violin"]:
                                # Statistical chart specific settings
                                if series_type == "histogram":
                                    trace_kwargs["x"] = y_vals  # Histogram uses x for values
                                    del trace_kwargs["y"]
                                elif series_type in ["box", "violin"]:
                                    trace_kwargs["y"] = y_vals
                                    trace_kwargs["x"] = [label] * len(y_vals) if len(y_vals) > 0 else [label]
                                
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>Value: %{{y if series_type in ['box', 'violin'] else 'x'}}<extra></extra>"
                                ))
                                
                            elif series_type in ["bubble"]:
                                # Bubble chart specific settings
                                sizes = series.get("size", [20] * len(y_vals))
                                trace_kwargs["mode"] = "markers"
                                trace_kwargs["marker"] = {"size": sizes}
                                if color:
                                    trace_kwargs["marker"]["color"] = color
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>X: %{{x}}<br>Y: %{{y}}<br>Size: %{{marker.size}}<extra></extra>"
                                ))
                                
                            elif series_type == "heatmap":
                                # Heatmap specific settings
                                heatmap_kwargs = {
                                    "z": series.get("z", y_vals),  # Use z data if provided, otherwise y_vals
                                    "x": series.get("x", x_vals),  # Use x data if provided, otherwise x_vals
                                    "y": series.get("y", [label]),  # Use y data if provided, otherwise label
                                    "name": label,
                                    "colorscale": series.get("colorscale", "Viridis"),
                                    "showscale": series.get("showscale", True)
                                }
                                
                                # Handle text data safely (convert to strings to avoid concatenation errors)
                                if "text" in series:
                                    try:
                                        heatmap_kwargs["text"] = [[str(cell) for cell in row] for row in series["text"]]
                                    except:
                                        current_app.logger.warning(f"⚠️ Could not process heatmap text data for {label}")
                                
                                # Handle colorbar settings
                                if "colorbar" in series:
                                    heatmap_kwargs["colorbar"] = series["colorbar"]
                                
                                # Handle zmin/zmax
                                if "zmin" in series:
                                    heatmap_kwargs["zmin"] = series["zmin"]
                                if "zmax" in series:
                                    heatmap_kwargs["zmax"] = series["zmax"]
                                
                                fig.add_trace(go.Heatmap(**heatmap_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>X: %{{x}}<br>Y: %{{y}}<br>Value: %{{z}}<extra></extra>"
                                ))
                                
                            else:
                                # Generic handling for other chart types
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>Value: %{{y}}<extra></extra>"
                                ))
                        else:
                            # Fallback to scatter if chart type not recognized
                            #current_app.logger.warning(f"⚠️ Unknown chart type '{series_type}', falling back to scatter")
                            fig.add_trace(go.Scatter(**trace_kwargs,
                                mode='markers',
                                hovertemplate=f"<b>{label}</b><br>Category: %{{x}}<br>Value: %{{y}}<extra></extra>"
                            ))

                # --- Layout updates ---
                layout_updates = {}
                
                # Title and axis labels
                layout_updates["title"] = title
                
                # Handle axis labels based on chart type
                if chart_type != "pie":
                    layout_updates["xaxis_title"] = chart_meta.get("x_label", chart_config.get("x_axis_title", "X"))
                    layout_updates["yaxis_title"] = chart_meta.get("primary_y_label", chart_config.get("primary_y_label", "Y"))
                
                # Font
                if font_family or font_size or font_color:
                    layout_updates["font"] = {}
                    if font_family:
                        layout_updates["font"]["family"] = font_family
                    if font_size:
                        layout_updates["font"]["size"] = font_size
                    if font_color:
                        layout_updates["font"]["color"] = font_color
                
                # Backgrounds
                if chart_background:
                    layout_updates["paper_bgcolor"] = chart_background
                if plot_background:
                    layout_updates["plot_bgcolor"] = plot_background
                
                # Legend configuration
                show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                # current_app.logger.debug(f"Legend setting: {chart_meta.get('legend')}")
                # current_app.logger.debug(f"Showlegend setting: {chart_meta.get('showlegend')}")
                # current_app.logger.debug(f"Final show_legend: {show_legend}")
                
                if show_legend:
                    # current_app.logger.debug("Configuring legend for Plotly")
                    if legend_position:
                        # Map 'top', 'bottom', 'left', 'right' to valid Plotly legend positions
                        pos_map = {
                            "top":    dict(x=0.5, y=1.1, xanchor="center", yanchor="top"),
                            "bottom": dict(x=0.5, y=-0.2, xanchor="center", yanchor="bottom"),
                            "left":   dict(x=-0.2, y=0.5, xanchor="left", yanchor="middle"),
                            "right":  dict(x=1.1, y=0.5, xanchor="right", yanchor="middle"),
                        }
                        if legend_position in pos_map:
                            layout_updates["legend"] = pos_map[legend_position]
                            # current_app.logger.debug(f"Set legend position: {legend_position}")
                    if legend_font_size:
                        layout_updates.setdefault("legend", {})["font"] = {"size": legend_font_size}
                        # current_app.logger.debug(f"Set legend font size: {legend_font_size}")
                else:
                    layout_updates["showlegend"] = False
                    # current_app.logger.debug("Legend disabled for Plotly")
                
                # Bar mode for stacked charts
                if barmode:
                    layout_updates["barmode"] = barmode
                elif chart_type == "stacked_column":
                    layout_updates["barmode"] = "stack"
                
                # Axis min/max and tick format (only for non-pie charts)
                if chart_type != "pie":
                    if y_axis_min_max:
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        # Handle "auto" value for y-axis min/max
                        if y_axis_min_max == "auto":
                            # Don't set range for auto - let Plotly auto-scale
                            pass
                        else:
                            layout_updates["yaxis"]["range"] = y_axis_min_max
                            # Also set autorange to false to ensure the range is respected
                            layout_updates["yaxis"]["autorange"] = False
                            # Force the range to be applied
                            layout_updates["yaxis"]["fixedrange"] = False
                            # Ensure the range is properly set
                            # current_app.logger.debug(f"Setting Y-axis range to: {y_axis_min_max}")
                    if axis_tick_format:
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["yaxis"]["tickformat"] = axis_tick_format
                        # Also apply to secondary y-axis if it's a currency format
                        if "$" in axis_tick_format:
                            layout_updates["yaxis2"] = layout_updates.get("yaxis2", {})
                            layout_updates["yaxis2"]["tickformat"] = axis_tick_format
                    
                    # Secondary y-axis formatting
                    if secondary_y_axis_format:
                        layout_updates["yaxis2"] = layout_updates.get("yaxis2", {})
                        layout_updates["yaxis2"]["tickformat"] = secondary_y_axis_format
                    if secondary_y_axis_min_max:
                        layout_updates["yaxis2"] = layout_updates.get("yaxis2", {})
                        # Handle "auto" value for secondary y-axis min/max
                        if secondary_y_axis_min_max == "auto":
                            # Don't set range for auto - let Matplotlib auto-scale
                            pass
                        elif isinstance(secondary_y_axis_min_max, list) and len(secondary_y_axis_min_max) == 2:
                            layout_updates["yaxis2"]["range"] = secondary_y_axis_min_max
                        else:
                            current_app.logger.warning(f"Invalid secondary_y_axis_min_max format: {secondary_y_axis_min_max}")
                    
                    # Axis tick font size
                    if axis_tick_font_size:
                        layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["xaxis"]["tickfont"] = {"size": axis_tick_font_size}
                        layout_updates["yaxis"]["tickfont"] = {"size": axis_tick_font_size}
                    # Gridlines
                    if show_gridlines is not None:
                        layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["xaxis"]["showgrid"] = bool(show_gridlines)
                        layout_updates["yaxis"]["showgrid"] = bool(show_gridlines)
                    if gridline_color:
                        layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["xaxis"]["gridcolor"] = gridline_color
                        layout_updates["yaxis"]["gridcolor"] = gridline_color
                    if gridline_style:
                        # Map gridline styles to valid Plotly dash styles
                        dash_map = {
                            "solid": "solid", 
                            "dashed": "dash",
                            "dash": "dash",  # Map 'dash' to 'dash'
                            "dashdot": "dashdot", 
                            "dotted": "dot",
                            "dot": "dot",
                            "dotdash": "dashdot"
                        }
                        dash_style = dash_map.get(gridline_style, "solid")
                        layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["xaxis"]["griddash"] = dash_style
                        layout_updates["yaxis"]["griddash"] = dash_style
                
                # Data labels (for bar/line traces)
                show_data_labels = chart_meta.get("data_labels", True)
                value_format = chart_meta.get("value_format", "")
                
                # Only enable data labels if explicitly set to True AND any data label settings are provided
                if show_data_labels and (data_label_format or data_label_font_size or data_label_color):
                    show_data_labels = True
                elif not show_data_labels:
                    # If data_labels is explicitly set to False, respect that setting
                    show_data_labels = False
                
                # Debug logging for data labels
                # current_app.logger.debug(f"Original data_labels setting: {chart_meta.get('data_labels')}")
                # current_app.logger.debug(f"Final show_data_labels: {show_data_labels}")
                # current_app.logger.debug(f"Data label format: {data_label_format}")
                # current_app.logger.debug(f"Data label font size: {data_label_font_size}")
                # current_app.logger.debug(f"Data label color: {data_label_color}")
                # current_app.logger.debug(f"Value format: {value_format}")
                # current_app.logger.debug(f"Chart config keys: {list(chart_config.keys())}")
                # current_app.logger.debug(f"Chart meta keys: {list(chart_meta.keys())}")
                
                if show_data_labels and (data_label_format or value_format or data_label_font_size or data_label_color):
                    # current_app.logger.debug(f"Processing {len(fig.data)} traces for data labels")
                    for i, trace in enumerate(fig.data):
                        # current_app.logger.debug(f"Trace {i}: type={trace.type}, mode={getattr(trace, 'mode', 'N/A')}")
                        # Handle both bar and line charts (line charts are scatter with mode='lines')
                        if trace.type in ['bar', 'scatter']:
                            # Use value_format from chart_meta if available, otherwise use data_label_format
                            format_to_use = value_format if value_format else data_label_format
                            
                            # Determine if this is a line chart (scatter with lines mode)
                            is_line_chart = trace.type == 'scatter' and trace.mode and 'lines' in trace.mode
                            
                            if format_to_use:
                                if trace.type == 'bar':
                                    trace.update(texttemplate=f"%{{y:{format_to_use}}}", textposition="auto")
                                elif trace.type == 'scatter':
                                    if is_line_chart:
                                        # For line charts, show labels at the data points
                                        trace.update(texttemplate=f"%{{y:{format_to_use}}}", textposition="top center")
                                    else:
                                        # For scatter plots
                                        trace.update(texttemplate=f"%{{y:{format_to_use}}}", textposition="top center")
                            else:
                                # If no format specified, still show data labels
                                if trace.type == 'bar':
                                    trace.update(texttemplate="%{y}", textposition="auto")
                                elif trace.type == 'scatter':
                                    if is_line_chart:
                                        # For line charts, show labels at the data points
                                        trace.update(texttemplate="%{y}", textposition="top center")
                                    else:
                                        # For scatter plots
                                        trace.update(texttemplate="%{y}", textposition="top center")
                            
                            # Apply font styling
                            if data_label_font_size or data_label_color:
                                trace.update(textfont={})
                                if data_label_font_size:
                                    trace.textfont["size"] = data_label_font_size
                                if data_label_color:
                                    trace.textfont["color"] = data_label_color
                
                # Annotations
                if annotations:
                    layout_updates["annotations"] = []
                    for ann in annotations:
                        ann_dict = {
                            "text": ann.get("text", ""),
                            "x": ann.get("x_value", ann.get("x")),
                            "y": ann.get("y_value", ann.get("y")),
                            "showarrow": True
                        }
                        layout_updates["annotations"].append(ann_dict)
                
                # --- Apply margin settings ---
                if margin:
                    layout_updates["margin"] = margin
                
                # Apply figure size if specified
                if figsize:
                    layout_updates["width"] = figsize[0] * 100  # Convert to pixels
                    layout_updates["height"] = figsize[1] * 100  # Convert to pixels
                    # current_app.logger.debug(f"Applied Plotly figsize: {figsize} -> width={figsize[0]*100}, height={figsize[1]*100}")
                
                # Apply axis label distances
                if chart_type != "pie":
                    if x_axis_label_distance or y_axis_label_distance or axis_tick_distance:
                        layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        
                        if x_axis_label_distance:
                            layout_updates["xaxis"]["title"] = layout_updates["xaxis"].get("title", {})
                            layout_updates["xaxis"]["title"]["standoff"] = x_axis_label_distance
                            # Also apply to tick distance for better control
                            layout_updates["xaxis"]["ticklen"] = x_axis_label_distance
                        
                        if y_axis_label_distance:
                            layout_updates["yaxis"]["title"] = layout_updates["yaxis"].get("title", {})
                            layout_updates["yaxis"]["title"]["standoff"] = y_axis_label_distance
                        
                        if axis_tick_distance:
                            layout_updates["xaxis"]["ticklen"] = axis_tick_distance
                            layout_updates["yaxis"]["ticklen"] = axis_tick_distance

                # Tick mark control for Plotly
                if show_x_ticks is not None or show_y_ticks is not None:
                    layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                    layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                    if show_x_ticks is not None:
                        layout_updates["xaxis"]["showticklabels"] = bool(show_x_ticks)
                        layout_updates["xaxis"]["ticks"] = "" if not show_x_ticks else "outside"
                    if show_y_ticks is not None:
                        layout_updates["yaxis"]["showticklabels"] = bool(show_y_ticks)
                        layout_updates["yaxis"]["ticks"] = "" if not show_y_ticks else "outside"

                fig.update_layout(**layout_updates)

                # --- Matplotlib static chart for DOCX ---
                if chart_type == "pie":
                    # Check if this is an expanded pie chart
                    expanded_segment = chart_meta.get("expanded_segment")
                    
                    if expanded_segment and len(series_data) == 1:
                        # Create subplot for expanded pie chart
                        mpl_figsize = figsize if figsize else (15, 8)
                        fig_mpl, (ax1, ax2) = plt.subplots(1, 2, figsize=mpl_figsize, dpi=200)
                        
                        # Apply background colors to Matplotlib figure
                        if chart_background:
                            fig_mpl.patch.set_facecolor(chart_background)
                        if plot_background:
                            ax1.set_facecolor(plot_background)
                            ax2.set_facecolor(plot_background)
                        
                        series = series_data[0]
                        labels = series.get("labels", x_values)
                        values = series.get("values", [])
                        color = series.get("marker", {}).get("color") if "marker" in series else colors
                        marker_line = series.get("marker", {}).get("line", {}) if "marker" in series else {}
                        explode = series.get("pull")
                        opacity = series.get("opacity", chart_meta.get("opacity"))
                        textinfo = series.get("textinfo", chart_meta.get("textinfo", "percent"))
                        textposition = series.get("textposition", chart_meta.get("textposition", "inside")).lower()
                        value_format_str = chart_meta.get("value_format", ".1f")
                        data_labels_enabled = bool(chart_meta.get("data_labels", True))
                        data_label_font_size = chart_meta.get("data_label_font_size", font_size or 10)
                        data_label_color = chart_meta.get("data_label_color", font_color or "#000000")
                        start_angle = startangle if 'startangle' in locals() and startangle is not None else 90
                        sort_order = chart_meta.get("sort_order")

                        # Optional sorting
                        if sort_order in ("ascending", "descending") and values:
                            zipped = list(zip(values, labels, color if isinstance(color, list) else [color]*len(labels), explode if isinstance(explode, list) else [0]*len(labels)))
                            reverse = sort_order == "descending"
                            zipped.sort(key=lambda t: (t[0] if t[0] is not None else 0), reverse=reverse)
                            values, labels, color_list, explode_list = zip(*zipped)
                            values = list(values)
                            labels = list(labels)
                            color = list(color_list)
                            explode = list(explode_list)

                        # Build autopct based on textinfo/value_format
                        def make_autopct(fmt:str, include_percent:bool, include_value:bool):
                            def _inner(pct):
                                total = sum(values) if values else 0
                                val = pct * total / 100.0
                                parts = []
                                if include_value:
                                    try:
                                        parts.append(f"{val:{fmt}}")
                                    except Exception:
                                        parts.append(f"{val:.1f}")
                                if include_percent:
                                    parts.append(f"{pct:.1f}%")
                                return " " .join(parts)
                            return _inner

                        include_label = "label" in (textinfo or "")
                        include_percent = "percent" in (textinfo or "")
                        include_value = "value" in (textinfo or "")

                        autopct_callable = None
                        if data_labels_enabled and (include_percent or include_value):
                            autopct_callable = make_autopct(value_format_str, include_percent, include_value)

                        # Wedge and text props
                        wedgeprops = {}
                        if isinstance(marker_line, dict):
                            if marker_line.get("color"):
                                wedgeprops["edgecolor"] = marker_line.get("color")
                            if marker_line.get("width") is not None:
                                wedgeprops["linewidth"] = marker_line.get("width")
                        if opacity is not None:
                            wedgeprops["alpha"] = opacity

                        textprops = {"color": data_label_color, "fontsize": data_label_font_size}
                        if chart_meta.get("font_family"):
                            textprops["fontfamily"] = chart_meta.get("font_family")

                        # Positioning
                        pctdistance = 0.6 if textposition == "inside" else 1.15
                        labeldistance = 1.1 if textposition != "inside" else 1.05
                        
                        # Create pie chart
                        wedges, texts, autotexts = ax1.pie(
                            values,
                            labels=labels if include_label else None,
                            autopct=autopct_callable,
                            colors=color,
                            startangle=start_angle,
                            explode=explode,
                            wedgeprops=wedgeprops,
                            pctdistance=pctdistance,
                            labeldistance=labeldistance,
                            textprops=textprops,
                        )

                        # Style the autopct texts
                        for autotext in autotexts or []:
                            autotext.set_color(data_label_color)
                            autotext.set_fontsize(data_label_font_size)
                            autotext.set_fontweight('bold')
                            if chart_meta.get("font_family"):
                                autotext.set_fontfamily(chart_meta.get("font_family"))
                        
                        ax1.set_title(title, fontsize=font_size or 14, weight='bold', pad=20, color=font_color if font_color else None, fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None)
                        
                        # Add legend for pie chart
                        show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                        if show_legend:
                            # Initialize legend_loc for pie charts
                            legend_loc = 'best'  # default
                            if legend_position:
                                loc_map = {
                                    "top": "upper center",
                                    "bottom": "lower center", 
                                    "left": "center left",
                                    "right": "center right"
                                }
                                legend_loc = loc_map.get(legend_position, 'best')
                            
                            # Force legend to bottom if specified
                            if legend_position == "bottom":
                                ax1.legend(wedges, labels, loc='lower center', bbox_to_anchor=(0.5, -0.15), fontsize=legend_font_size)
                            else:
                                ax1.legend(wedges, labels, loc=legend_loc, fontsize=legend_font_size)
                        
                        # Create bar chart for expanded segment
                        if expanded_segment in labels:
                            segment_idx = labels.index(expanded_segment)
                            segment_value = values[segment_idx]
                            segment_color = color[segment_idx] if isinstance(color, list) and segment_idx < len(color) else color
                            
                            ax2.bar([expanded_segment], [segment_value], color=segment_color, alpha=0.7)
                            ax2.set_title(f"{expanded_segment} Details", fontsize=font_size or 12, weight='bold')
                            ax2.set_ylabel("Value")
                            
                            # Add value label on bar
                            ax2.text(0, segment_value, f"{segment_value}", ha='center', va='bottom', fontweight='bold')
                        
                    else:
                        # Regular pie chart
                        mpl_figsize = figsize if figsize else (10, 8)
                        fig_mpl, ax = plt.subplots(figsize=mpl_figsize, dpi=200)
                        
                        # Apply background colors to Matplotlib figure
                        if chart_background:
                            fig_mpl.patch.set_facecolor(chart_background)
                        if plot_background:
                            ax.set_facecolor(plot_background)
                        
                        if len(series_data) == 1:
                            series = series_data[0]
                            labels = series.get("labels", x_values)
                            values = series.get("values", [])
                            color = series.get("marker", {}).get("color") if "marker" in series else colors
                            marker_line = series.get("marker", {}).get("line", {}) if "marker" in series else {}
                            explode = series.get("pull")
                            opacity = series.get("opacity", chart_meta.get("opacity"))
                            textinfo = series.get("textinfo", chart_meta.get("textinfo", "percent"))
                            textposition = series.get("textposition", chart_meta.get("textposition", "inside")).lower()
                            value_format_str = chart_meta.get("value_format", ".1f")
                            data_labels_enabled = bool(chart_meta.get("data_labels", True))
                            data_label_font_size = chart_meta.get("data_label_font_size", font_size or 10)
                            data_label_color = chart_meta.get("data_label_color", font_color or "#000000")
                            start_angle = startangle if 'startangle' in locals() and startangle is not None else 90
                            sort_order = chart_meta.get("sort_order")

                            # Optional sorting
                            if sort_order in ("ascending", "descending") and values:
                                zipped = list(zip(values, labels, color if isinstance(color, list) else [color]*len(labels), explode if isinstance(explode, list) else [0]*len(labels)))
                                reverse = sort_order == "descending"
                                zipped.sort(key=lambda t: (t[0] if t[0] is not None else 0), reverse=reverse)
                                values, labels, color_list, explode_list = zip(*zipped)
                                values = list(values)
                                labels = list(labels)
                                color = list(color_list)
                                explode = list(explode_list)

                            # Build autopct based on textinfo/value_format
                            def make_autopct(fmt:str, include_percent:bool, include_value:bool):
                                def _inner(pct):
                                    total = sum(values) if values else 0
                                    val = pct * total / 100.0
                                    parts = []
                                    if include_value:
                                        try:
                                            parts.append(f"{val:{fmt}}")
                                        except Exception:
                                            parts.append(f"{val:.1f}")
                                    if include_percent:
                                        parts.append(f"{pct:.1f}%")
                                    return " " .join(parts)
                                return _inner

                            include_label = "label" in (textinfo or "")
                            include_percent = "percent" in (textinfo or "")
                            include_value = "value" in (textinfo or "")

                            autopct_callable = None
                            if data_labels_enabled and (include_percent or include_value):
                                autopct_callable = make_autopct(value_format_str, include_percent, include_value)

                            # Wedge and text props
                            wedgeprops = {}
                            if isinstance(marker_line, dict):
                                if marker_line.get("color"):
                                    wedgeprops["edgecolor"] = marker_line.get("color")
                                if marker_line.get("width") is not None:
                                    wedgeprops["linewidth"] = marker_line.get("width")
                            if opacity is not None:
                                wedgeprops["alpha"] = opacity

                            textprops = {"color": data_label_color, "fontsize": data_label_font_size}
                            if chart_meta.get("font_family"):
                                textprops["fontfamily"] = chart_meta.get("font_family")

                            # Positioning
                            pctdistance = 0.6 if textposition == "inside" else 1.15
                            labeldistance = 1.1 if textposition != "inside" else 1.05
                            
                            # Create pie chart
                            wedges, texts, autotexts = ax.pie(
                                values,
                                labels=labels if include_label else None,
                                autopct=autopct_callable,
                                colors=color,
                                startangle=start_angle,
                                explode=explode,
                                wedgeprops=wedgeprops,
                                pctdistance=pctdistance,
                                labeldistance=labeldistance,
                                textprops=textprops,
                            )
                            
                            # Style the text
                            for autotext in autotexts or []:
                                autotext.set_color(data_label_color)
                                autotext.set_fontsize(data_label_font_size)
                                autotext.set_fontweight('bold')
                                if chart_meta.get("font_family"):
                                    autotext.set_fontfamily(chart_meta.get("font_family"))
                            
                            ax.set_title(title, fontsize=font_size or 14, weight='bold', pad=20, color=font_color if font_color else None, fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None)
                            
                            # Add legend for regular pie chart
                            show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                            if show_legend:
                                # Initialize legend_loc for pie charts
                                legend_loc = 'best'  # default
                                if legend_position:
                                    loc_map = {
                                        "top": "upper center",
                                        "bottom": "lower center", 
                                        "left": "center left",
                                        "right": "center right"
                                    }
                                    legend_loc = loc_map.get(legend_position, 'best')
                                
                                # Force legend to bottom if specified
                                if legend_position == "bottom":
                                    ax.legend(wedges, labels, loc='lower center', bbox_to_anchor=(0.5, -0.15), fontsize=legend_font_size)
                                else:
                                    ax.legend(wedges, labels, loc=legend_loc, fontsize=legend_font_size)
                        
                elif chart_type in ["bar of pie", "bar_of_pie"]:
                    # Matplotlib version of bar of pie
                    mpl_figsize = figsize if figsize else (10, 5)
                    fig_mpl, (ax1, ax2) = plt.subplots(1, 2, figsize=mpl_figsize, dpi=200, gridspec_kw={'width_ratios': [2, 1]})
                    
                    # Apply background colors to Matplotlib figure
                    if chart_background:
                        fig_mpl.patch.set_facecolor(chart_background)
                    if plot_background:
                        ax1.set_facecolor(plot_background)
                        ax2.set_facecolor(plot_background)
                    labels = series_meta.get("labels", x_values)
                    values = series_meta.get("values", [])
                    colors = series_meta.get("colors", [])
                    other_labels = chart_meta.get("other_labels", [])
                    other_values = chart_meta.get("other_values", [])
                    other_colors = chart_meta.get("other_colors", [])
                    if not (other_labels and other_values):
                        if "other_label_range" in chart_meta and "other_value_range" in chart_meta and "source_sheet" in chart_meta:
                            wb = openpyxl.load_workbook(data_file_path, data_only=True)
                            sheet = wb[chart_meta["source_sheet"]]
                            other_labels = extract_excel_range(sheet, chart_meta["other_label_range"])
                            other_values = extract_excel_range(sheet, chart_meta["other_value_range"])
                    # Pie chart
                    if colors:
                        wedges, texts, autotexts = ax1.pie(values, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90)
                    else:
                        wedges, texts, autotexts = ax1.pie(values, labels=labels, autopct='%1.1f%%', startangle=90)
                    for autotext in autotexts:
                        autotext.set_color('white')
                        autotext.set_fontweight('bold')
                    ax1.set_title(title, fontsize=font_size or 14, weight='bold', pad=20)
                    # Move legend outside
                    show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                    legend_font_size = chart_meta.get("legend_font_size", 8)
                    if show_legend:
                        ax1.legend(wedges, labels, loc='center left', bbox_to_anchor=(1, 0.5), fontsize=legend_font_size)
                    
                    # Bar chart - Filter out empty/null values first
                    filtered_data = []
                    for label, value in zip(other_labels, other_values):
                        if (label is not None and str(label).strip() != "" and 
                            value is not None and str(value).strip() != "" and 
                            str(value).strip() != "0"):
                            filtered_data.append((label, value))
                    
                    if filtered_data:
                        filtered_labels, filtered_values = zip(*filtered_data)
                    else:
                        filtered_labels, filtered_values = [], []
                    
                    # Create individual bars (not stacked)
                    if filtered_labels and filtered_values:
                        # Use individual colors for each bar
                        bar_colors = []
                        for i in range(len(filtered_labels)):
                            if other_colors and i < len(other_colors):
                                bar_colors.append(other_colors[i])
                            elif colors and i < len(colors):
                                bar_colors.append(colors[i])
                            else:
                                bar_colors.append('#1f77b4')  # Default blue
                        
                        bars = ax2.bar(range(len(filtered_labels)), filtered_values, color=bar_colors, alpha=0.7)
                    else:
                        bars = []
                    
                    expanded_segment = chart_meta.get("expanded_segment", "Other")
                    ax2.set_title(f"Breakdown of '{expanded_segment}'", fontsize=font_size or 12, weight='bold')
                    # Use proper Y-axis label from configuration
                    y_axis_title = chart_meta.get("y_axis_title", "Value")
                    ax2.set_ylabel(y_axis_title, fontsize=label_fontsize if 'label_fontsize' in locals() else 10)
                    # Set X-axis label
                    x_axis_title = chart_meta.get("x_axis_title", "Categories")
                    ax2.set_xlabel(x_axis_title, fontsize=label_fontsize if 'label_fontsize' in locals() else 10)
                    # Set x-tick labels for filtered data
                    if filtered_labels:
                        ax2.set_xticks(range(len(filtered_labels)))
                        # Format x-axis labels as percentages
                        formatted_x_labels = []
                        for label in filtered_labels:
                            if isinstance(label, (int, float)):
                                if label <= 1.0:  # Likely decimal format (0.06)
                                    formatted_x_labels.append(f"{label * 100:.1f}%")
                                else:  # Likely already percentage format (6.0)
                                    formatted_x_labels.append(f"{label:.1f}%")
                            else:
                                # Convert string to float and handle
                                try:
                                    val = float(label)
                                    if val <= 1.0:
                                        formatted_x_labels.append(f"{val * 100:.1f}%")
                                    else:
                                        formatted_x_labels.append(f"{val:.1f}%")
                                except:
                                    formatted_x_labels.append(str(label))
                        ax2.set_xticklabels(formatted_x_labels, rotation=0)
                    # Add data labels with proper formatting
                    value_format = chart_meta.get("value_format", ".2f")
                    data_label_font_size = chart_meta.get("data_label_font_size", 10)
                    data_label_color = chart_meta.get("data_label_color", "#000000")
                    for bar, v in zip(bars, filtered_values):
                        # Format percentage values as XX.X% instead of 0.XXX
                        if isinstance(v, (int, float)):
                            if v <= 1.0:  # Likely decimal format (0.11)
                                formatted_value = f"{v * 100:.1f}%"
                            else:  # Likely already percentage format (11.0)
                                formatted_value = f"{v:.1f}%"
                        else:
                            # Convert string to float and handle
                            try:
                                val = float(v)
                                if val <= 1.0:
                                    formatted_value = f"{val * 100:.1f}%"
                                else:
                                    formatted_value = f"{val:.1f}%"
                            except:
                                formatted_value = str(v)
                        ax2.text(bar.get_x() + bar.get_width()/2, v, formatted_value, ha='center', va='bottom', fontweight='bold', fontsize=data_label_font_size, color=data_label_color)

                else:
                    # Bar, line, area charts
                    mpl_figsize = figsize if figsize else (10, 6)
                    # current_app.logger.debug(f"Applied Matplotlib figsize: {mpl_figsize}")
                    fig_mpl, ax1 = plt.subplots(figsize=mpl_figsize, dpi=200)
                    ax2 = ax1.twinx()
                    
                    # Apply background colors to Matplotlib figure
                    if chart_background:
                        fig_mpl.patch.set_facecolor(chart_background)
                    if plot_background:
                        ax1.set_facecolor(plot_background)
                        ax2.set_facecolor(plot_background)

                    for i, series in enumerate(series_data):
                        label = series.get("name", f"Series {i+1}")
                        series_type = series.get("type", "bar").lower()
                        
                        # Map heatmap to imshow for Matplotlib
                        if series_type == "heatmap":
                            mpl_chart_type = "imshow"
                        else:
                            mpl_chart_type = chart_type_mapping_mpl.get(series_type, "scatter")
                        
                        color = None
                        if "marker" in series and isinstance(series["marker"], dict) and "color" in series["marker"]:
                            color = series["marker"]["color"]
                        elif bar_colors:
                            color = bar_colors
                        elif i < len(colors):
                            color = colors[i]

                        y_vals = series.get("values")
                        value_range = series.get("value_range")
                        if value_range:
                            # Check if value_range is already extracted (list) or still a string
                            if isinstance(value_range, list):
                                y_vals = value_range
                            else:
                                y_vals = extract_values_from_range(value_range)

                        # --- Apply grouping and sorting ---
                        x_vals = x_values
                        if y_vals is not None and x_vals is not None:
                            x_vals, y_vals = group_and_sort(x_vals, y_vals, data_grouping, sort_order)

                        trace_kwargs = {
                            "x": x_vals,
                            "y": y_vals,
                            "name": label,
                        }
                        
                        # Handle colors
                        if color:
                            if isinstance(color, list) and series_type == "bar":
                                trace_kwargs["marker"] = dict(color=color)
                            elif isinstance(color, str):
                                trace_kwargs["marker_color"] = color
                        
                        # Handle bar-specific properties
                        if bar_width and series_type == "bar":
                            trace_kwargs["width"] = bar_width
                        if orientation and series_type == "bar":
                            trace_kwargs["orientation"] = orientation[0].lower() if isinstance(orientation, str) else orientation
                        if bar_border_color and series_type == "bar":
                            trace_kwargs["marker_line_color"] = bar_border_color
                        if bar_border_width and series_type == "bar":
                            trace_kwargs["marker_line_width"] = bar_border_width
                        
                        # Handle line-specific properties
                        if line_width and series_type in ["line", "scatter", "scatter_line"]:
                            trace_kwargs["line_width"] = line_width
                        if marker_size and series_type in ["line", "scatter", "scatter_line"]:
                            trace_kwargs["marker_size"] = marker_size
                        if line_style and series_type in ["line", "scatter", "scatter_line"]:
                            # Map line styles to Plotly format
                            line_style_map = {
                                "solid": "solid",
                                "dashed": "dash",
                                "dotted": "dot",
                                "dashdot": "dashdot"
                            }
                            trace_kwargs["line_dash"] = line_style_map.get(line_style, "solid")
                        
                        # Handle opacity
                        if fill_opacity:
                            trace_kwargs["opacity"] = fill_opacity

                        # Add traces based on chart type - REPLACE THE RESTRICTIVE IF/ELIF BLOCKS
                        # Generic chart type handling for Plotly
                        chart_type_mapping = {
                            # Bar charts
                            "bar": go.Bar,
                            "column": go.Bar,
                            "stacked_column": go.Bar,
                            "horizontal_bar": go.Bar,
                            
                            # Line charts
                            "line": go.Scatter,
                            "scatter": go.Scatter,
                            "scatter_line": go.Scatter,
                            
                            # Area charts
                            "area": go.Scatter,
                            "filled_area": go.Scatter,
                            
                            # Pie charts
                            "pie": go.Pie,
                            "donut": go.Pie,
                            
                            # 3D charts
                            "scatter3d": go.Scatter3d,
                            "surface": go.Surface,
                            "mesh3d": go.Mesh3d,
                            
                            # Statistical charts
                            "histogram": go.Histogram,
                            "box": go.Box,
                            "violin": go.Violin,
                            
                            # Financial charts
                            "candlestick": go.Candlestick,
                            "ohlc": go.Ohlc,
                            
                            # Geographic charts
                            "scattergeo": go.Scattergeo,
                            "choropleth": go.Choropleth,
                            
                            # Other charts
                            "bubble": go.Scatter,
                            "heatmap": go.Heatmap,
                            "contour": go.Contour,
                            "waterfall": go.Waterfall,
                            "funnel": go.Funnel,
                            "sunburst": go.Sunburst,
                            "treemap": go.Treemap,
                            "icicle": go.Icicle,
                            "sankey": go.Sankey,
                            "table": go.Table,
                            "indicator": go.Indicator
                        }
                        
                        # Get the appropriate Plotly chart class
                        plotly_chart_class = chart_type_mapping.get(series_type)
                        
                        if plotly_chart_class:
                            # Prepare trace arguments based on chart type
                            if series_type in ["bar", "column", "stacked_column", "horizontal_bar"]:
                                # Bar chart specific settings
                                # REMOVE: if chart_type == "stacked_column": trace_kwargs["barmode"] = "stack"
                                if orientation and orientation.lower() == "horizontal":
                                    trace_kwargs["orientation"] = "h"
                                    # Swap x and y for horizontal bars
                                    trace_kwargs["x"], trace_kwargs["y"] = trace_kwargs["y"], trace_kwargs["x"]
                                
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>Category: %{{x}}<br>Value: %{{y}}<extra></extra>"
                                ))
                                
                            elif series_type in ["line", "scatter", "scatter_line"]:
                                # Line/Scatter chart specific settings
                                mode = "lines+markers" if series_type == "scatter_line" else "markers" if series_type == "scatter" else "lines"
                                trace_kwargs["mode"] = mode
                                
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>Category: %{{x}}<br>Value: %{{y}}<extra></extra>"
                                ))
                                
                            elif series_type in ["area", "filled_area"]:
                                # Area chart specific settings
                                trace_kwargs["mode"] = "lines"
                                trace_kwargs["fill"] = "tozeroy"
                                
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>Category: %{{x}}<br>Value: %{{y}}<extra></extra>"
                                ))
                                
                            elif series_type == "pie":
                                # Pie chart specific settings
                                pie_kwargs = {
                                    "labels": x_vals,
                                    "values": y_vals,
                                    "name": label,
                                    "textinfo": "label+percent+value" if data_labels else "none",
                                    "textposition": "outside",
                                    "hole": hole if hole is not None else (0.4 if series_type == "donut" else 0.0)
                                }
                                
                                # Add pie chart specific attributes
                                if startangle is not None:
                                    pie_kwargs["rotation"] = startangle
                                if pull is not None:
                                    # pull can be a single value or a list
                                    if isinstance(pull, (int, float)):
                                        pie_kwargs["pull"] = [pull] * len(y_vals)
                                    elif isinstance(pull, list):
                                        pie_kwargs["pull"] = pull
                                
                                if color:
                                    pie_kwargs["marker"] = dict(colors=color) if isinstance(color, list) else dict(colors=[color])
                                
                                fig.add_trace(plotly_chart_class(**pie_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>%{{label}}: %{{value}}<extra></extra>"
                                ))
                                
                            elif series_type in ["scatter3d", "surface", "mesh3d"]:
                                # 3D chart specific settings
                                if "z" not in trace_kwargs and len(y_vals) > 0:
                                    # Create a simple z-axis if not provided
                                    trace_kwargs["z"] = [i for i in range(len(y_vals))]
                                
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>X: %{{x}}<br>Y: %{{y}}<br>Z: %{{z}}<extra></extra>"
                                ))
                                
                            elif series_type in ["histogram", "box", "violin"]:
                                # Statistical chart specific settings
                                if series_type == "histogram":
                                    trace_kwargs["x"] = y_vals  # Histogram uses x for values
                                    del trace_kwargs["y"]
                                elif series_type in ["box", "violin"]:
                                    trace_kwargs["y"] = y_vals
                                    trace_kwargs["x"] = [label] * len(y_vals) if len(y_vals) > 0 else [label]
                                
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>Value: %{{y if series_type in ['box', 'violin'] else 'x'}}<extra></extra>"
                                ))
                                
                            elif series_type in ["bubble"]:
                                # Bubble chart specific settings
                                sizes = series.get("size", [20] * len(y_vals))
                                trace_kwargs["mode"] = "markers"
                                trace_kwargs["marker"] = {"size": sizes}
                                if color:
                                    trace_kwargs["marker"]["color"] = color
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>X: %{{x}}<br>Y: %{{y}}<br>Size: %{{marker.size}}<extra></extra>"
                                ))
                                
                            elif series_type == "heatmap":
                                # Heatmap specific settings
                                heatmap_kwargs = {
                                    "z": series.get("z", y_vals),  # Use z data if provided, otherwise y_vals
                                    "x": series.get("x", x_vals),  # Use x data if provided, otherwise x_vals
                                    "y": series.get("y", [label]),  # Use y data if provided, otherwise label
                                    "name": label,
                                    "colorscale": series.get("colorscale", "Viridis"),
                                    "showscale": series.get("showscale", True)
                                }
                                
                                # Handle text data safely (convert to strings to avoid concatenation errors)
                                if "text" in series:
                                    try:
                                        heatmap_kwargs["text"] = [[str(cell) for cell in row] for row in series["text"]]
                                    except:
                                        current_app.logger.warning(f"⚠️ Could not process heatmap text data for {label}")
                                
                                # Handle colorbar settings
                                if "colorbar" in series:
                                    heatmap_kwargs["colorbar"] = series["colorbar"]
                                
                                # Handle zmin/zmax
                                if "zmin" in series:
                                    heatmap_kwargs["zmin"] = series["zmin"]
                                if "zmax" in series:
                                    heatmap_kwargs["zmax"] = series["zmax"]
                                
                                fig.add_trace(go.Heatmap(**heatmap_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>X: %{{x}}<br>Y: %{{y}}<br>Value: %{{z}}<extra></extra>"
                                ))
                                
                            else:
                                # Generic handling for other chart types
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>Value: %{{y}}<extra></extra>"
                                ))
                        else:
                            # Fallback to scatter if chart type not recognized
                            #current_app.logger.warning(f"⚠️ Unknown chart type '{series_type}', falling back to scatter")
                            fig.add_trace(go.Scatter(**trace_kwargs,
                                mode='markers',
                                hovertemplate=f"<b>{label}</b><br>Category: %{{x}}<br>Value: %{{y}}<extra></extra>"
                            ))

                # --- Process heatmap charts separately ---
                # Heatmaps need special handling because they have x, y, z, text data directly
                for i, series in enumerate(series_data):
                    label = series.get("name", f"Series {i+1}")
                    series_type = series.get("type", "bar").lower()
                    
                    if series_type == "heatmap":
                        # Heatmap specific settings
                        heatmap_kwargs = {
                            "z": series.get("z", []),  # Use z data directly
                            "x": series.get("x", []),  # Use x data directly
                            "y": series.get("y", []),  # Use y data directly
                            "name": label,
                            "colorscale": series.get("colorscale", "Viridis"),
                            "showscale": series.get("showscale", True)
                        }
                        
                        # Handle text data safely (convert to strings to avoid concatenation errors)
                        if "text" in series:
                            try:
                                heatmap_kwargs["text"] = [[str(cell) for cell in row] for row in series["text"]]
                            except:
                                current_app.logger.warning(f"⚠️ Could not process heatmap text data for {label}")
                        
                        # Handle colorbar settings
                        if "colorbar" in series:
                            heatmap_kwargs["colorbar"] = series["colorbar"]
                        
                        # Handle zmin/zmax
                        if "zmin" in series:
                            heatmap_kwargs["zmin"] = series["zmin"]
                        if "zmax" in series:
                            heatmap_kwargs["zmax"] = series["zmax"]
                        
                        fig.add_trace(go.Heatmap(**heatmap_kwargs,
                            hovertemplate=f"<b>{label}</b><br>X: %{{x}}<br>Y: %{{y}}<br>Value: %{{z}}<extra></extra>"
                        ))

                # --- Layout updates ---
                layout_updates = {}
                
                # Title and axis labels
                layout_updates["title"] = title
                
                # Handle axis labels based on chart type
                if chart_type != "pie":
                    layout_updates["xaxis_title"] = chart_meta.get("x_label", chart_config.get("x_axis_title", "X"))
                    layout_updates["yaxis_title"] = chart_meta.get("primary_y_label", chart_config.get("primary_y_label", "Y"))
                
                # Font
                if font_family or font_size or font_color:
                    layout_updates["font"] = {}
                    if font_family:
                        layout_updates["font"]["family"] = font_family
                    if font_size:
                        layout_updates["font"]["size"] = font_size
                    if font_color:
                        layout_updates["font"]["color"] = font_color
                
                # Backgrounds
                if chart_background:
                    layout_updates["paper_bgcolor"] = chart_background
                if plot_background:
                    layout_updates["plot_bgcolor"] = plot_background
                
                # Legend configuration
                show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                # current_app.logger.debug(f"Legend setting: {chart_meta.get('legend')}")
                # current_app.logger.debug(f"Showlegend setting: {chart_meta.get('showlegend')}")
                # current_app.logger.debug(f"Final show_legend: {show_legend}")
                
                if show_legend:
                    # current_app.logger.debug("Configuring legend for Plotly")
                    if legend_position:
                        # Map 'top', 'bottom', 'left', 'right' to valid Plotly legend positions
                        pos_map = {
                            "top":    dict(x=0.5, y=1.1, xanchor="center", yanchor="top"),
                            "bottom": dict(x=0.5, y=-0.2, xanchor="center", yanchor="bottom"),
                            "left":   dict(x=-0.2, y=0.5, xanchor="left", yanchor="middle"),
                            "right":  dict(x=1.1, y=0.5, xanchor="right", yanchor="middle"),
                        }
                        if legend_position in pos_map:
                            layout_updates["legend"] = pos_map[legend_position]
                            # current_app.logger.debug(f"Set legend position: {legend_position}")
                    if legend_font_size:
                        layout_updates.setdefault("legend", {})["font"] = {"size": legend_font_size}
                        # current_app.logger.debug(f"Set legend font size: {legend_font_size}")
                else:
                    layout_updates["showlegend"] = False
                    # current_app.logger.debug("Legend disabled for Plotly")
                
                # Bar mode for stacked charts
                if barmode:
                    layout_updates["barmode"] = barmode
                elif chart_type == "stacked_column":
                    layout_updates["barmode"] = "stack"
                
                # Axis min/max and tick format (only for non-pie charts)
                if chart_type != "pie":
                    if y_axis_min_max:
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        # Handle "auto" value for y-axis min/max
                        if y_axis_min_max == "auto":
                            # Don't set range for auto - let Plotly auto-scale
                            pass
                        else:
                            layout_updates["yaxis"]["range"] = y_axis_min_max
                            # Also set autorange to false to ensure the range is respected
                            layout_updates["yaxis"]["autorange"] = False
                            # Force the range to be applied
                            layout_updates["yaxis"]["fixedrange"] = False
                            # Ensure the range is properly set
                            # current_app.logger.debug(f"Setting Y-axis range to: {y_axis_min_max}")
                    if axis_tick_format:
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["yaxis"]["tickformat"] = axis_tick_format
                        # Also apply to secondary y-axis if it's a currency format
                        if "$" in axis_tick_format:
                            layout_updates["yaxis2"] = layout_updates.get("yaxis2", {})
                            layout_updates["yaxis2"]["tickformat"] = axis_tick_format
                    
                    # Secondary y-axis formatting
                    if secondary_y_axis_format:
                        layout_updates["yaxis2"] = layout_updates.get("yaxis2", {})
                        layout_updates["yaxis2"]["tickformat"] = secondary_y_axis_format
                    if secondary_y_axis_min_max:
                        layout_updates["yaxis2"] = layout_updates.get("yaxis2", {})
                        # Handle "auto" value for secondary y-axis min/max
                        if secondary_y_axis_min_max == "auto":
                            # Don't set range for auto - let Matplotlib auto-scale
                            pass
                        elif isinstance(secondary_y_axis_min_max, list) and len(secondary_y_axis_min_max) == 2:
                            layout_updates["yaxis2"]["range"] = secondary_y_axis_min_max
                        else:
                            current_app.logger.warning(f"Invalid secondary_y_axis_min_max format: {secondary_y_axis_min_max}")
                    
                    # Axis tick font size
                    if axis_tick_font_size:
                        layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["xaxis"]["tickfont"] = {"size": axis_tick_font_size}
                        layout_updates["yaxis"]["tickfont"] = {"size": axis_tick_font_size}
                    # Gridlines
                    if show_gridlines is not None:
                        layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["xaxis"]["showgrid"] = bool(show_gridlines)
                        layout_updates["yaxis"]["showgrid"] = bool(show_gridlines)
                    if gridline_color:
                        layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["xaxis"]["gridcolor"] = gridline_color
                        layout_updates["yaxis"]["gridcolor"] = gridline_color
                    if gridline_style:
                        # Map gridline styles to valid Plotly dash styles
                        dash_map = {
                            "solid": "solid", 
                            "dashed": "dash",
                            "dash": "dash",  # Map 'dash' to 'dash'
                            "dashdot": "dashdot", 
                            "dotted": "dot",
                            "dot": "dot",
                            "dotdash": "dashdot"
                        }
                        dash_style = dash_map.get(gridline_style, "solid")
                        layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["xaxis"]["griddash"] = dash_style
                        layout_updates["yaxis"]["griddash"] = dash_style
                
                # Data labels (for bar/line traces)
                show_data_labels = chart_meta.get("data_labels", True)
                value_format = chart_meta.get("value_format", "")
                
                # Only enable data labels if explicitly set to True AND any data label settings are provided
                if show_data_labels and (data_label_format or data_label_font_size or data_label_color):
                    show_data_labels = True
                elif not show_data_labels:
                    # If data_labels is explicitly set to False, respect that setting
                    show_data_labels = False
                
                # Debug logging for data labels
                # current_app.logger.debug(f"Original data_labels setting: {chart_meta.get('data_labels')}")
                # current_app.logger.debug(f"Final show_data_labels: {show_data_labels}")
                # current_app.logger.debug(f"Data label format: {data_label_format}")
                # current_app.logger.debug(f"Data label font size: {data_label_font_size}")
                # current_app.logger.debug(f"Data label color: {data_label_color}")
                # current_app.logger.debug(f"Value format: {value_format}")
                # current_app.logger.debug(f"Chart config keys: {list(chart_config.keys())}")
                # current_app.logger.debug(f"Chart meta keys: {list(chart_meta.keys())}")
                
                if show_data_labels and (data_label_format or value_format or data_label_font_size or data_label_color):
                    # current_app.logger.debug(f"Processing {len(fig.data)} traces for data labels")
                    for i, trace in enumerate(fig.data):
                        # current_app.logger.debug(f"Trace {i}: type={trace.type}, mode={getattr(trace, 'mode', 'N/A')}")
                        # Handle both bar and line charts (line charts are scatter with mode='lines')
                        if trace.type in ['bar', 'scatter']:
                            # Use value_format from chart_meta if available, otherwise use data_label_format
                            format_to_use = value_format if value_format else data_label_format
                            
                            # Determine if this is a line chart (scatter with lines mode)
                            is_line_chart = trace.type == 'scatter' and trace.mode and 'lines' in trace.mode
                            
                            if format_to_use:
                                if trace.type == 'bar':
                                    trace.update(texttemplate=f"%{{y:{format_to_use}}}", textposition="auto")
                                elif trace.type == 'scatter':
                                    if is_line_chart:
                                        # For line charts, show labels at the data points
                                        trace.update(texttemplate=f"%{{y:{format_to_use}}}", textposition="top center")
                                    else:
                                        # For scatter plots
                                        trace.update(texttemplate=f"%{{y:{format_to_use}}}", textposition="top center")
                            else:
                                # If no format specified, still show data labels
                                if trace.type == 'bar':
                                    trace.update(texttemplate="%{y}", textposition="auto")
                                elif trace.type == 'scatter':
                                    if is_line_chart:
                                        # For line charts, show labels at the data points
                                        trace.update(texttemplate="%{y}", textposition="top center")
                                    else:
                                        # For scatter plots
                                        trace.update(texttemplate="%{y}", textposition="top center")
                            
                            # Apply font styling
                            if data_label_font_size or data_label_color:
                                trace.update(textfont={})
                                if data_label_font_size:
                                    trace.textfont["size"] = data_label_font_size
                                if data_label_color:
                                    trace.textfont["color"] = data_label_color
                
                # Annotations
                if annotations:
                    layout_updates["annotations"] = []
                    for ann in annotations:
                        ann_dict = {
                            "text": ann.get("text", ""),
                            "x": ann.get("x_value", ann.get("x")),
                            "y": ann.get("y_value", ann.get("y")),
                            "showarrow": True
                        }
                        layout_updates["annotations"].append(ann_dict)
                
                # --- Apply margin settings ---
                if margin:
                    layout_updates["margin"] = margin
                
                # Apply figure size if specified
                if figsize:
                    layout_updates["width"] = figsize[0] * 100  # Convert to pixels
                    layout_updates["height"] = figsize[1] * 100  # Convert to pixels
                    # current_app.logger.debug(f"Applied Plotly figsize: {figsize} -> width={figsize[0]*100}, height={figsize[1]*100}")
                
                # Apply axis label distances
                if chart_type != "pie":
                    if x_axis_label_distance or y_axis_label_distance or axis_tick_distance:
                        layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        
                        if x_axis_label_distance:
                            layout_updates["xaxis"]["title"] = layout_updates["xaxis"].get("title", {})
                            layout_updates["xaxis"]["title"]["standoff"] = x_axis_label_distance
                            # Also apply to tick distance for better control
                            layout_updates["xaxis"]["ticklen"] = x_axis_label_distance
                        
                        if y_axis_label_distance:
                            layout_updates["yaxis"]["title"] = layout_updates["yaxis"].get("title", {})
                            layout_updates["yaxis"]["title"]["standoff"] = y_axis_label_distance
                        
                        if axis_tick_distance:
                            layout_updates["xaxis"]["ticklen"] = axis_tick_distance
                            layout_updates["yaxis"]["ticklen"] = axis_tick_distance

                # Tick mark control for Plotly
                if show_x_ticks is not None or show_y_ticks is not None:
                    layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                    layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                    if show_x_ticks is not None:
                        layout_updates["xaxis"]["showticklabels"] = bool(show_x_ticks)
                        layout_updates["xaxis"]["ticks"] = "" if not show_x_ticks else "outside"
                    if show_y_ticks is not None:
                        layout_updates["yaxis"]["showticklabels"] = bool(show_y_ticks)
                        layout_updates["yaxis"]["ticks"] = "" if not show_y_ticks else "outside"

                fig.update_layout(**layout_updates)

                # --- Matplotlib static chart for DOCX ---
                if chart_type == "pie":
                    # Check if this is an expanded pie chart
                    expanded_segment = chart_meta.get("expanded_segment")
                    
                    if expanded_segment and len(series_data) == 1:
                        # Create subplot for expanded pie chart
                        mpl_figsize = figsize if figsize else (15, 8)
                        fig_mpl, (ax1, ax2) = plt.subplots(1, 2, figsize=mpl_figsize, dpi=200)
                        
                        # Apply background colors to Matplotlib figure
                        if chart_background:
                            fig_mpl.patch.set_facecolor(chart_background)
                        if plot_background:
                            ax1.set_facecolor(plot_background)
                            ax2.set_facecolor(plot_background)
                        
                        series = series_data[0]
                        labels = series.get("labels", x_values)
                        values = series.get("values", [])
                        color = series.get("marker", {}).get("color") if "marker" in series else colors
                        marker_line = series.get("marker", {}).get("line", {}) if "marker" in series else {}
                        explode = series.get("pull")
                        opacity = series.get("opacity", chart_meta.get("opacity"))
                        textinfo = series.get("textinfo", chart_meta.get("textinfo", "percent"))
                        textposition = series.get("textposition", chart_meta.get("textposition", "inside")).lower()
                        value_format_str = chart_meta.get("value_format", ".1f")
                        data_labels_enabled = bool(chart_meta.get("data_labels", True))
                        data_label_font_size = chart_meta.get("data_label_font_size", font_size or 10)
                        data_label_color = chart_meta.get("data_label_color", font_color or "#000000")
                        start_angle = startangle if 'startangle' in locals() and startangle is not None else 90
                        sort_order = chart_meta.get("sort_order")

                        # Optional sorting
                        if sort_order in ("ascending", "descending") and values:
                            zipped = list(zip(values, labels, color if isinstance(color, list) else [color]*len(labels), explode if isinstance(explode, list) else [0]*len(labels)))
                            reverse = sort_order == "descending"
                            zipped.sort(key=lambda t: (t[0] if t[0] is not None else 0), reverse=reverse)
                            values, labels, color_list, explode_list = zip(*zipped)
                            values = list(values)
                            labels = list(labels)
                            color = list(color_list)
                            explode = list(explode_list)

                        # Build autopct based on textinfo/value_format
                        def make_autopct(fmt:str, include_percent:bool, include_value:bool):
                            def _inner(pct):
                                total = sum(values) if values else 0
                                val = pct * total / 100.0
                                parts = []
                                if include_value:
                                    try:
                                        parts.append(f"{val:{fmt}}")
                                    except Exception:
                                        parts.append(f"{val:.1f}")
                                if include_percent:
                                    parts.append(f"{pct:.1f}%")
                                return " " .join(parts)
                            return _inner

                        include_label = "label" in (textinfo or "")
                        include_percent = "percent" in (textinfo or "")
                        include_value = "value" in (textinfo or "")

                        autopct_callable = None
                        if data_labels_enabled and (include_percent or include_value):
                            autopct_callable = make_autopct(value_format_str, include_percent, include_value)

                        # Wedge and text props
                        wedgeprops = {}
                        if isinstance(marker_line, dict):
                            if marker_line.get("color"):
                                wedgeprops["edgecolor"] = marker_line.get("color")
                            if marker_line.get("width") is not None:
                                wedgeprops["linewidth"] = marker_line.get("width")
                        if opacity is not None:
                            wedgeprops["alpha"] = opacity

                        textprops = {"color": data_label_color, "fontsize": data_label_font_size}
                        if chart_meta.get("font_family"):
                            textprops["fontfamily"] = chart_meta.get("font_family")

                        # Positioning
                        pctdistance = 0.6 if textposition == "inside" else 1.15
                        labeldistance = 1.1 if textposition != "inside" else 1.05
                        
                        # Create pie chart
                        wedges, texts, autotexts = ax1.pie(
                            values,
                            labels=labels if include_label else None,
                            autopct=autopct_callable,
                            colors=color,
                            startangle=start_angle,
                            explode=explode,
                            wedgeprops=wedgeprops,
                            pctdistance=pctdistance,
                            labeldistance=labeldistance,
                            textprops=textprops,
                        )

                        # Style the autopct texts
                        for autotext in autotexts or []:
                            autotext.set_color(data_label_color)
                            autotext.set_fontsize(data_label_font_size)
                            autotext.set_fontweight('bold')
                            if chart_meta.get("font_family"):
                                autotext.set_fontfamily(chart_meta.get("font_family"))
                        
                        ax1.set_title(title, fontsize=font_size or 14, weight='bold', pad=20, color=font_color if font_color else None, fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None)
                        
                        # Add legend for pie chart
                        show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                        if show_legend:
                            # Initialize legend_loc for pie charts
                            legend_loc = 'best'  # default
                            if legend_position:
                                loc_map = {
                                    "top": "upper center",
                                    "bottom": "lower center", 
                                    "left": "center left",
                                    "right": "center right"
                                }
                                legend_loc = loc_map.get(legend_position, 'best')
                            
                            # Force legend to bottom if specified
                            if legend_position == "bottom":
                                ax1.legend(wedges, labels, loc='lower center', bbox_to_anchor=(0.5, -0.15), fontsize=legend_font_size)
                            else:
                                ax1.legend(wedges, labels, loc=legend_loc, fontsize=legend_font_size)
                        
                        # Create bar chart for expanded segment
                        if expanded_segment in labels:
                            segment_idx = labels.index(expanded_segment)
                            segment_value = values[segment_idx]
                            segment_color = color[segment_idx] if isinstance(color, list) and segment_idx < len(color) else color
                            
                            ax2.bar([expanded_segment], [segment_value], color=segment_color, alpha=0.7)
                            ax2.set_title(f"{expanded_segment} Details", fontsize=font_size or 12, weight='bold')
                            ax2.set_ylabel("Value")
                            
                            # Add value label on bar
                            ax2.text(0, segment_value, f"{segment_value}", ha='center', va='bottom', fontweight='bold')
                        
                    else:
                        # Regular pie chart
                        mpl_figsize = figsize if figsize else (10, 8)
                        fig_mpl, ax = plt.subplots(figsize=mpl_figsize, dpi=200)
                        
                        # Apply background colors to Matplotlib figure
                        if chart_background:
                            fig_mpl.patch.set_facecolor(chart_background)
                        if plot_background:
                            ax.set_facecolor(plot_background)
                        
                        if len(series_data) == 1:
                            series = series_data[0]
                            labels = series.get("labels", x_values)
                            values = series.get("values", [])
                            color = series.get("marker", {}).get("color") if "marker" in series else colors
                            marker_line = series.get("marker", {}).get("line", {}) if "marker" in series else {}
                            explode = series.get("pull")
                            opacity = series.get("opacity", chart_meta.get("opacity"))
                            textinfo = series.get("textinfo", chart_meta.get("textinfo", "percent"))
                            textposition = series.get("textposition", chart_meta.get("textposition", "inside")).lower()
                            value_format_str = chart_meta.get("value_format", ".1f")
                            data_labels_enabled = bool(chart_meta.get("data_labels", True))
                            data_label_font_size = chart_meta.get("data_label_font_size", font_size or 10)
                            data_label_color = chart_meta.get("data_label_color", font_color or "#000000")
                            start_angle = startangle if 'startangle' in locals() and startangle is not None else 90
                            sort_order = chart_meta.get("sort_order")

                            # Optional sorting
                            if sort_order in ("ascending", "descending") and values:
                                zipped = list(zip(values, labels, color if isinstance(color, list) else [color]*len(labels), explode if isinstance(explode, list) else [0]*len(labels)))
                                reverse = sort_order == "descending"
                                zipped.sort(key=lambda t: (t[0] if t[0] is not None else 0), reverse=reverse)
                                values, labels, color_list, explode_list = zip(*zipped)
                                values = list(values)
                                labels = list(labels)
                                color = list(color_list)
                                explode = list(explode_list)

                            # Build autopct based on textinfo/value_format
                            def make_autopct(fmt:str, include_percent:bool, include_value:bool):
                                def _inner(pct):
                                    total = sum(values) if values else 0
                                    val = pct * total / 100.0
                                    parts = []
                                    if include_value:
                                        try:
                                            parts.append(f"{val:{fmt}}")
                                        except Exception:
                                            parts.append(f"{val:.1f}")
                                    if include_percent:
                                        parts.append(f"{pct:.1f}%")
                                    return " " .join(parts)
                                return _inner

                            include_label = "label" in (textinfo or "")
                            include_percent = "percent" in (textinfo or "")
                            include_value = "value" in (textinfo or "")

                            autopct_callable = None
                            if data_labels_enabled and (include_percent or include_value):
                                autopct_callable = make_autopct(value_format_str, include_percent, include_value)

                            # Wedge and text props
                            wedgeprops = {}
                            if isinstance(marker_line, dict):
                                if marker_line.get("color"):
                                    wedgeprops["edgecolor"] = marker_line.get("color")
                                if marker_line.get("width") is not None:
                                    wedgeprops["linewidth"] = marker_line.get("width")
                            if opacity is not None:
                                wedgeprops["alpha"] = opacity

                            textprops = {"color": data_label_color, "fontsize": data_label_font_size}
                            if chart_meta.get("font_family"):
                                textprops["fontfamily"] = chart_meta.get("font_family")

                            # Positioning
                            pctdistance = 0.6 if textposition == "inside" else 1.15
                            labeldistance = 1.1 if textposition != "inside" else 1.05
                            
                            # Create pie chart
                            wedges, texts, autotexts = ax.pie(
                                values,
                                labels=labels if include_label else None,
                                autopct=autopct_callable,
                                colors=color,
                                startangle=start_angle,
                                explode=explode,
                                wedgeprops=wedgeprops,
                                pctdistance=pctdistance,
                                labeldistance=labeldistance,
                                textprops=textprops,
                            )
                            
                            # Style the text
                            for autotext in autotexts or []:
                                autotext.set_color(data_label_color)
                                autotext.set_fontsize(data_label_font_size)
                                autotext.set_fontweight('bold')
                                if chart_meta.get("font_family"):
                                    autotext.set_fontfamily(chart_meta.get("font_family"))
                            
                            ax.set_title(title, fontsize=font_size or 14, weight='bold', pad=20, color=font_color if font_color else None, fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None)
                            
                            # Add legend for regular pie chart
                            show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                            if show_legend:
                                # Initialize legend_loc for pie charts
                                legend_loc = 'best'  # default
                                if legend_position:
                                    loc_map = {
                                        "top": "upper center",
                                        "bottom": "lower center", 
                                        "left": "center left",
                                        "right": "center right"
                                    }
                                    legend_loc = loc_map.get(legend_position, 'best')
                                
                                # Force legend to bottom if specified
                                if legend_position == "bottom":
                                    ax.legend(wedges, labels, loc='lower center', bbox_to_anchor=(0.5, -0.15), fontsize=legend_font_size)
                                else:
                                    ax.legend(wedges, labels, loc=legend_loc, fontsize=legend_font_size)
                        
                elif chart_type in ["bar of pie", "bar_of_pie"]:
                    # Matplotlib version of bar of pie
                    mpl_figsize = figsize if figsize else (10, 5)
                    fig_mpl, (ax1, ax2) = plt.subplots(1, 2, figsize=mpl_figsize, dpi=200, gridspec_kw={'width_ratios': [2, 1]})
                    
                    # Apply background colors to Matplotlib figure
                    if chart_background:
                        fig_mpl.patch.set_facecolor(chart_background)
                    if plot_background:
                        ax1.set_facecolor(plot_background)
                        ax2.set_facecolor(plot_background)
                    labels = series_meta.get("labels", x_values)
                    values = series_meta.get("values", [])
                    colors = series_meta.get("colors", [])
                    other_labels = chart_meta.get("other_labels", [])
                    other_values = chart_meta.get("other_values", [])
                    other_colors = chart_meta.get("other_colors", [])
                    if not (other_labels and other_values):
                        if "other_label_range" in chart_meta and "other_value_range" in chart_meta and "source_sheet" in chart_meta:
                            wb = openpyxl.load_workbook(data_file_path, data_only=True)
                            sheet = wb[chart_meta["source_sheet"]]
                            other_labels = extract_excel_range(sheet, chart_meta["other_label_range"])
                            other_values = extract_excel_range(sheet, chart_meta["other_value_range"])
                    # Pie chart
                    if colors:
                        wedges, texts, autotexts = ax1.pie(values, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90)
                    else:
                        wedges, texts, autotexts = ax1.pie(values, labels=labels, autopct='%1.1f%%', startangle=90)
                    for autotext in autotexts:
                        autotext.set_color('white')
                        autotext.set_fontweight('bold')
                    ax1.set_title(title, fontsize=font_size or 14, weight='bold', pad=20)
                    # Move legend outside
                    show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                    legend_font_size = chart_meta.get("legend_font_size", 8)
                    if show_legend:
                        ax1.legend(wedges, labels, loc='center left', bbox_to_anchor=(1, 0.5), fontsize=legend_font_size)
                    
                    # Bar chart - Filter out empty/null values first
                    filtered_data = []
                    for label, value in zip(other_labels, other_values):
                        if (label is not None and str(label).strip() != "" and 
                            value is not None and str(value).strip() != "" and 
                            str(value).strip() != "0"):
                            filtered_data.append((label, value))
                    
                    if filtered_data:
                        filtered_labels, filtered_values = zip(*filtered_data)
                    else:
                        filtered_labels, filtered_values = [], []
                    
                    # Create individual bars (not stacked)
                    if filtered_labels and filtered_values:
                        # Use individual colors for each bar
                        bar_colors = []
                        for i in range(len(filtered_labels)):
                            if other_colors and i < len(other_colors):
                                bar_colors.append(other_colors[i])
                            elif colors and i < len(colors):
                                bar_colors.append(colors[i])
                            else:
                                bar_colors.append('#1f77b4')  # Default blue
                        
                        bars = ax2.bar(range(len(filtered_labels)), filtered_values, color=bar_colors, alpha=0.7)
                    else:
                        bars = []
                    
                    expanded_segment = chart_meta.get("expanded_segment", "Other")
                    ax2.set_title(f"Breakdown of '{expanded_segment}'", fontsize=font_size or 12, weight='bold')
                    # Use proper Y-axis label from configuration
                    y_axis_title = chart_meta.get("y_axis_title", "Value")
                    ax2.set_ylabel(y_axis_title, fontsize=label_fontsize if 'label_fontsize' in locals() else 10)
                    # Set X-axis label
                    x_axis_title = chart_meta.get("x_axis_title", "Categories")
                    ax2.set_xlabel(x_axis_title, fontsize=label_fontsize if 'label_fontsize' in locals() else 10)
                    # Set x-tick labels for filtered data
                    if filtered_labels:
                        ax2.set_xticks(range(len(filtered_labels)))
                        # Format x-axis labels as percentages
                        formatted_x_labels = []
                        for label in filtered_labels:
                            if isinstance(label, (int, float)):
                                if label <= 1.0:  # Likely decimal format (0.06)
                                    formatted_x_labels.append(f"{label * 100:.1f}%")
                                else:  # Likely already percentage format (6.0)
                                    formatted_x_labels.append(f"{label:.1f}%")
                            else:
                                # Convert string to float and handle
                                try:
                                    val = float(label)
                                    if val <= 1.0:
                                        formatted_x_labels.append(f"{val * 100:.1f}%")
                                    else:
                                        formatted_x_labels.append(f"{val:.1f}%")
                                except:
                                    formatted_x_labels.append(str(label))
                        ax2.set_xticklabels(formatted_x_labels, rotation=0)
                    # Add data labels with proper formatting
                    value_format = chart_meta.get("value_format", ".2f")
                    data_label_font_size = chart_meta.get("data_label_font_size", 10)
                    data_label_color = chart_meta.get("data_label_color", "#000000")
                    for bar, v in zip(bars, filtered_values):
                        # Format percentage values as XX.X% instead of 0.XXX
                        if isinstance(v, (int, float)):
                            if v <= 1.0:  # Likely decimal format (0.11)
                                formatted_value = f"{v * 100:.1f}%"
                            else:  # Likely already percentage format (11.0)
                                formatted_value = f"{v:.1f}%"
                        else:
                            # Convert string to float and handle
                            try:
                                val = float(v)
                                if val <= 1.0:
                                    formatted_value = f"{val * 100:.1f}%"
                                else:
                                    formatted_value = f"{val:.1f}%"
                            except:
                                formatted_value = str(v)
                        ax2.text(bar.get_x() + bar.get_width()/2, v, formatted_value, ha='center', va='bottom', fontweight='bold', fontsize=data_label_font_size, color=data_label_color)

                else:
                    # Bar, line, area charts
                    mpl_figsize = figsize if figsize else (10, 6)
                    # current_app.logger.debug(f"Applied Matplotlib figsize: {mpl_figsize}")
                    fig_mpl, ax1 = plt.subplots(figsize=mpl_figsize, dpi=200)
                    ax2 = ax1.twinx()
                    
                    # Apply background colors to Matplotlib figure
                    if chart_background:
                        fig_mpl.patch.set_facecolor(chart_background)
                    if plot_background:
                        ax1.set_facecolor(plot_background)
                        ax2.set_facecolor(plot_background)

                    for i, series in enumerate(series_data):
                        label = series.get("name", f"Series {i+1}")
                        series_type = series.get("type", "bar").lower()
                        
                        # Map heatmap to imshow for Matplotlib
                        if series_type == "heatmap":
                            mpl_chart_type = "imshow"
                        else:
                            mpl_chart_type = chart_type_mapping_mpl.get(series_type, "scatter")
                        
                        color = None
                        if "marker" in series and isinstance(series["marker"], dict) and "color" in series["marker"]:
                            color = series["marker"]["color"]
                        elif bar_colors:
                            color = bar_colors
                        elif i < len(colors):
                            color = colors[i]

                        y_vals = series.get("values")
                        value_range = series.get("value_range")
                        if value_range:
                            # Check if value_range is already extracted (list) or still a string
                            if isinstance(value_range, list):
                                y_vals = value_range
                            else:
                                y_vals = extract_values_from_range(value_range)

                        # Generic chart type handling for Matplotlib
                        
                        mpl_chart_type = chart_type_mapping_mpl.get(series_type, "scatter")
                        
                        if mpl_chart_type == "bar":
                            # Add bar border parameters
                            edgecolor = bar_border_color if bar_border_color else 'none'
                            linewidth = bar_border_width if bar_border_width else 0
                            
                            if chart_type == "stacked_column":
                                # For stacked column, use bottom parameter
                                if i == 0:
                                    ax1.bar(x_values, y_vals, label=label, color=color, alpha=0.7, edgecolor=edgecolor, linewidth=linewidth)
                                    bottom_vals = y_vals
                                else:
                                    ax1.bar(x_values, y_vals, bottom=bottom_vals, label=label, color=color, alpha=0.7, edgecolor=edgecolor, linewidth=linewidth)
                                    bottom_vals = [sum(x) for x in zip(bottom_vals, y_vals)]
                            else:
                                if isinstance(color, list):
                                    for j, val in enumerate(y_vals):
                                        bar_color = color[j % len(color)]
                                        ax1.bar(x_values[j], val, color=bar_color, alpha=0.7, label=label if j == 0 else "", edgecolor=edgecolor, linewidth=linewidth)
                                else:
                                    ax1.bar(x_values, y_vals, label=label, color=color, alpha=0.7, edgecolor=edgecolor, linewidth=linewidth)
                                    
                        elif mpl_chart_type == "barh":
                            # Horizontal bar chart
                            # Add bar border parameters
                            edgecolor = bar_border_color if bar_border_color else 'none'
                            linewidth = bar_border_width if bar_border_width else 0
                            
                            if isinstance(color, list):
                                for j, val in enumerate(y_vals):
                                    bar_color = color[j % len(color)]
                                    ax1.barh(x_values[j], val, color=bar_color, alpha=0.7, label=label if j == 0 else "", edgecolor=edgecolor, linewidth=linewidth)
                            else:
                                ax1.barh(x_values, y_vals, label=label, color=color, alpha=0.7, edgecolor=edgecolor, linewidth=linewidth)
                                
                        elif mpl_chart_type == "plot":
                            # Line chart
                            marker = 'o' if series_type == "scatter_line" else None
                            if ax2:
                                ax2.plot(x_values, y_vals, label=label, color=color, marker=marker, linewidth=2)
                            else:
                                ax1.plot(x_values, y_vals, label=label, color=color, marker=marker, linewidth=2)
                            
                        elif mpl_chart_type == "scatter":
                            # Scatter plot and Bubble chart
                            if series_type == "bubble":
                                # Enhanced bubble chart with better styling
                                # Creating bubble chart for series: {label}
                                
                                # Get sizes from the series data structure
                                sizes = series.get("size", [20] * len(y_vals))
                                # Bubble chart data processed
                                
                                # Ensure all arrays have the same length
                                min_length = min(len(x_values), len(y_vals), len(sizes))
                                if min_length < len(x_values) or min_length < len(y_vals) or min_length < len(sizes):
                                    #current_app.logger.warning(f"⚠️ Array length mismatch! Truncating to {min_length}")
                                    x_values = x_values[:min_length]
                                    y_vals = y_vals[:min_length]
                                    sizes = sizes[:min_length]
                                    if isinstance(color, list):
                                        color = color[:min_length]
                                
                                # Scale sizes for better visual impact (bubble charts need much larger sizes)
                                scaled_sizes = [s * 20 for s in sizes]  # Much larger scale for better visibility
                                # Scaled sizes calculated
                                
                                # Colors
                                bubble_colors = color
                                if isinstance(color, list):
                                    bubble_colors = color
                                elif color:
                                    import matplotlib.cm as cm
                                    import numpy as np
                                    size_array = np.array(sizes)
                                    normalized_sizes = (size_array - size_array.min()) / (size_array.max() - size_array.min() + 1e-8)
                                    cmap = cm.viridis if color == 'auto' else cm.get_cmap('viridis')
                                    bubble_colors = [cmap(norm_size) for norm_size in normalized_sizes]
                                else:
                                    bubble_colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F']
                                
                                # Marker styles from config
                                marker_cfg = series.get("marker", {}) if isinstance(series.get("marker"), dict) else {}
                                marker_line_cfg = marker_cfg.get("line", {}) if isinstance(marker_cfg.get("line"), dict) else {}
                                alpha_val = marker_cfg.get("opacity", series.get("opacity", chart_meta.get("opacity", 0.8)))
                                edge_color = marker_line_cfg.get("color", 'white')
                                edge_width = marker_line_cfg.get("width", 2)
                                
                                # Create bubble chart with enhanced styling
                                scatter = ax1.scatter(
                                    x_values, y_vals, 
                                                     s=scaled_sizes, 
                                                     c=bubble_colors, 
                                    alpha=alpha_val,
                                    edgecolors=edge_color,
                                    linewidth=edge_width,
                                                     label=label,
                                    zorder=3)
                                
                                # Keep a reference for legend building
                                try:
                                    if not hasattr(ax1, "_bubble_handles"):
                                        ax1._bubble_handles = []
                                        ax1._bubble_labels = []
                                    ax1._bubble_handles.append(scatter)
                                    ax1._bubble_labels.append(label)
                                except Exception:
                                    pass
                                
                                # Create an explicit proxy handle so legend always renders
                                try:
                                    from matplotlib.lines import Line2D
                                    if isinstance(bubble_colors, list) and len(bubble_colors) > 0:
                                        legend_facecolor = bubble_colors[0]
                                    else:
                                        legend_facecolor = bubble_colors if bubble_colors else '#1f77b4'
                                    proxy = Line2D([0], [0], marker='o', linestyle='None', label=label,
                                                   markerfacecolor=legend_facecolor, markeredgecolor=edge_color,
                                                   markeredgewidth=edge_width, markersize=10, alpha=alpha_val)
                                    ax1._bubble_proxy = proxy
                                except Exception:
                                    pass
                                
                                # Add subtle shadow effect for depth
                                ax1.scatter(
                                    x_values, y_vals, 
                                          s=scaled_sizes, 
                                          c='black', 
                                          alpha=0.1, 
                                          zorder=1)
                                
                                # Add data labels for bubble chart if enabled
                                if show_data_labels:
                                    for j, (x, y, value_num) in enumerate(zip(x_values, y_vals, y_vals)):
                                        if j < len(x_values):
                                            # Format value for label using value_format and axis_tick_format
                                            try:
                                                if axis_tick_format and "$" in axis_tick_format:
                                                    num = float(value_num)
                                                    formatted_value = f"${num:,.0f}" if value_format in (".0f", None) else f"${num:,.0f}" if value_format == ".0f" else f"${num:,.0f}"
                                                    # Simplify: currency with thousands, ignore decimals beyond .0f for labels
                                                else:
                                                    num = float(value_num)
                                                    if value_format == ".0f":
                                                        formatted_value = f"{num:.0f}"
                                                    elif value_format == ".1f":
                                                        formatted_value = f"{num:.1f}"
                                                    elif value_format == ".2f":
                                                        formatted_value = f"{num:.2f}"
                                                    else:
                                                        formatted_value = f"{num:.0f}"
                                            except:
                                                formatted_value = str(value_num)
                                            
                                            # Add text label inside bubble
                                            ax1.text(x, y, formatted_value, 
                                                    ha='center', va='center', 
                                                    fontsize=data_label_font_size or 10,
                                                    color=data_label_color or '#000000',
                                                    fontweight='bold',
                                                    zorder=4)
                                
                                # Fix axis ranges for bubble charts to prevent layout shifts
                                if i == len(series_data) - 1:  # Only set once after all series
                                    # Always set fixed x-axis range for bubble charts to prevent layout shifts
                                    if x_axis_min_max and isinstance(x_axis_min_max, list) and len(x_axis_min_max) == 2:
                                        # Use explicit range from config
                                        ax1.set_xlim(x_axis_min_max[0], x_axis_min_max[1])
                                    else:
                                        # Set auto-calculated range with padding
                                        x_min, x_max = min(x_values), max(x_values)
                                        x_padding = (x_max - x_min) * 0.1
                                        ax1.set_xlim(x_min - x_padding, x_max + x_padding)
                                    
                                    # Always set fixed y-axis range for bubble charts to prevent layout shifts
                                    if y_axis_min_max and isinstance(y_axis_min_max, list) and len(y_axis_min_max) == 2:
                                        # Use explicit range from config
                                        ax1.set_ylim(y_axis_min_max[0], y_axis_min_max[1])
                                    else:
                                        # Set auto-calculated range with padding
                                        y_min, y_max = min(y_vals), max(y_vals)
                                        y_padding = (y_max - y_min) * 0.1
                                        ax1.set_ylim(y_min - y_padding, y_max + y_padding)
                                    
                                    # Apply axis label distances specifically for bubble charts
                                    if x_axis_label_distance is not None or y_axis_label_distance is not None:
                                        # Calculate labelpad values with multiplication for visibility
                                        x_labelpad = (x_axis_label_distance * 10) if x_axis_label_distance is not None else 50.0
                                        y_labelpad = (y_axis_label_distance * 10) if y_axis_label_distance is not None else 50.0
                                        
                                        # Get axis titles
                                        x_axis_title = chart_meta.get("x_label", chart_config.get("x_axis_title", ""))
                                        y_axis_title = chart_meta.get("primary_y_label", chart_config.get("primary_y_label", ""))
                                        
                                        # Set axis labels with distance
                                        if x_axis_title:
                                            ax1.set_xlabel(x_axis_title, fontsize=font_size or 12, color=font_color if font_color else 'black',
                                                        fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None,
                                                        labelpad=x_labelpad)
                                        if y_axis_title:
                                            ax1.set_ylabel(y_axis_title, fontsize=font_size or 12, color=font_color if font_color else 'black',
                                                        fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None,
                                                        labelpad=y_labelpad)
                                        
                                        current_app.logger.info(f"🎈 Bubble Chart Axis Label Distance - X: {x_axis_label_distance} → {x_labelpad}, Y: {y_axis_label_distance} → {y_labelpad}")
                            else:
                                # Enhanced scatter plot with custom styling
                                # Extract marker properties from series
                                marker_size = series.get("marker", {}).get("size", 50)
                                marker_color = series.get("marker", {}).get("color", color)
                                marker_symbol = series.get("marker", {}).get("symbol", "o")
                                marker_opacity = series.get("marker", {}).get("opacity", 0.8)
                                text_labels = series.get("text", [])
                                text_position = series.get("textposition", "top center")
                                
                                # Extract line properties from series
                                line_config = series.get("line", {})
                                line_color = line_config.get("color", marker_color)
                                line_width = line_config.get("width", 2)
                                line_dash = line_config.get("dash", "solid")
                                
                                # Handle mode setting
                                mode = series.get("mode", "markers")
                                
                                # Handle marker size (can be single value or list)
                                if isinstance(marker_size, list):
                                    sizes = marker_size
                                else:
                                    sizes = [marker_size] * len(y_vals)
                                
                                # Scale sizes for better visibility
                                scaled_sizes = [s * 2 for s in sizes]
                                
                                # Create enhanced scatter plot
                                scatter = ax1.scatter(x_values, y_vals, 
                                                     s=scaled_sizes, 
                                                     c=marker_color, 
                                                     alpha=marker_opacity,
                                                     edgecolors='white',
                                                     linewidth=1.5,
                                                     label=label,
                                                     zorder=3)
                                
                                # Add line if mode includes lines or if line properties are specified
                                if "lines" in mode or "line" in mode or line_config:
                                    # Convert dash style to matplotlib format
                                    if line_dash == "dash":
                                        linestyle = "--"
                                    elif line_dash == "dot":
                                        linestyle = ":"
                                    elif line_dash == "dashdot":
                                        linestyle = "-."
                                    else:
                                        linestyle = "-"  # solid
                                    
                                    ax1.plot(x_values, y_vals, 
                                            color=line_color, 
                                            linewidth=line_width, 
                                            linestyle=linestyle,
                                            alpha=marker_opacity,
                                            zorder=2)
                                
                                # Add subtle shadow effect for depth
                                ax1.scatter(x_values, y_vals, 
                                          s=scaled_sizes, 
                                          c='black', 
                                          alpha=0.1, 
                                          zorder=1)
                                
                                # Add data labels if text is provided
                                if text_labels and len(text_labels) == len(y_vals):
                                    for j, (x, y, text) in enumerate(zip(x_values, y_vals, text_labels)):
                                        if j < len(text_labels):
                                            # Determine label position based on textposition
                                            if "top" in text_position:
                                                y_offset = 2
                                            elif "bottom" in text_position:
                                                y_offset = -2
                                            else:
                                                y_offset = 0
                                            
                                            if "center" in text_position:
                                                ha = 'center'
                                            elif "left" in text_position:
                                                ha = 'left'
                                            elif "right" in text_position:
                                                ha = 'right'
                                            else:
                                                ha = 'center'
                                            
                                            # Add text label
                                            label_color = data_label_color or '#000000'
                                            ax1.text(x, y + y_offset, str(text), 
                                                    ha=ha, va='bottom' if y_offset > 0 else 'top',
                                                    fontsize=data_label_font_size or 10,
                                                    color=label_color,
                                                    fontweight='bold',
                                                    bbox=dict(boxstyle="round,pad=0.3", 
                                                             facecolor='white', 
                                                             alpha=0.9,
                                                             edgecolor='gray',
                                                             linewidth=0.5))
                        elif mpl_chart_type == "fill_between":
                            # Enhanced Area chart
                            # current_app.logger.debug(f"Processing area chart for series: {label}")
                            # current_app.logger.debug(f"X values: {x_vals}")
                            # current_app.logger.debug(f"Y values: {y_vals}")
                            
                            # Extract area-specific properties from series
                            fill_type = series.get("fill", "tozeroy")
                            line_color = series.get("line", {}).get("color", color)
                            line_width = series.get("line", {}).get("width", 2)
                            line_shape = series.get("line", {}).get("shape", "linear")
                            marker_symbol = series.get("marker", {}).get("symbol", "o")
                            marker_size = series.get("marker", {}).get("size", 6)
                            marker_color = series.get("marker", {}).get("color", line_color)
                            area_opacity = series.get("opacity", 0.6)
                            text_labels = series.get("text", [])
                            text_position = series.get("textposition", "top center")
                            
                            # Handle line shape
                            if line_shape == "spline":
                                linestyle = "-"
                                # For spline, we could add curve smoothing, but matplotlib doesn't have built-in splines
                            elif line_shape == "hv":
                                linestyle = "step"
                            elif line_shape == "vh":
                                linestyle = "step"
                            else:
                                linestyle = "-"  # linear
                            
                            # Create area fill based on fill type
                            if fill_type == "tozeroy":
                                # Fill from zero to y values
                                ax1.fill_between(x_vals, y_vals, alpha=area_opacity, label=label, color=line_color)
                            elif fill_type == "tonexty":
                                # Fill to next y values (for stacked areas)
                                if i == 0:
                                    # First series - fill from zero
                                    ax1.fill_between(x_vals, y_vals, alpha=area_opacity, label=label, color=line_color)
                                    # Store the cumulative values for next series
                                    if not hasattr(ax1, '_stacked_bottom'):
                                        ax1._stacked_bottom = {}
                                    ax1._stacked_bottom[label] = y_vals
                                else:
                                    # Get the bottom values from previous series
                                    prev_bottom = getattr(ax1, '_stacked_bottom', {}).get(series_data[i-1].get('name', f'series_{i-1}'), [0] * len(y_vals))
                                    # Calculate new bottom (cumulative)
                                    new_bottom = [b + y for b, y in zip(prev_bottom, y_vals)]
                                    # Fill between previous bottom and new bottom
                                    ax1.fill_between(x_vals, prev_bottom, new_bottom, alpha=area_opacity, label=label, color=line_color)
                                    # Store the new cumulative values
                                    ax1._stacked_bottom[label] = new_bottom
                            elif fill_type == "tonextx":
                                # Fill to next x values (horizontal stacking)
                                if i == 0:
                                    # First series - fill from zero
                                    ax1.fill_betweenx(y_vals, x_vals, alpha=area_opacity, label=label, color=line_color)
                                    # Store the cumulative values for next series
                                    if not hasattr(ax1, '_stacked_left'):
                                        ax1._stacked_left = {}
                                    ax1._stacked_left[label] = x_vals
                                else:
                                    # Get the left values from previous series
                                    prev_left = getattr(ax1, '_stacked_left', {}).get(series_data[i-1].get('name', f'series_{i-1}'), [0] * len(x_vals))
                                    # Calculate new left (cumulative)
                                    new_left = [l + x for l, x in zip(prev_left, x_vals)]
                                    # Fill between previous left and new left
                                    ax1.fill_betweenx(y_vals, prev_left, new_left, alpha=area_opacity, label=label, color=line_color)
                                    # Store the new cumulative values
                                    ax1._stacked_left[label] = new_left
                            else:
                                # Default fill
                                ax1.fill_between(x_vals, y_vals, alpha=area_opacity, label=label, color=line_color)
                            
                            # Add line on top of area
                            ax1.plot(x_vals, y_vals, color=line_color, linewidth=line_width, linestyle=linestyle, zorder=3)
                            
                            # Add markers if specified
                            if marker_symbol != "none":
                                ax1.scatter(x_vals, y_vals, color=marker_color, s=marker_size*20, 
                                          zorder=4, edgecolors='white', linewidth=1)
                            
                            # Add data labels if text is provided
                            if text_labels and len(text_labels) == len(y_vals):
                                for j, (x, y, text) in enumerate(zip(x_vals, y_vals, text_labels)):
                                    if j < len(text_labels):
                                        # Determine label position based on textposition
                                        if "top" in text_position:
                                            y_offset = 5
                                        elif "bottom" in text_position:
                                            y_offset = -5
                                        else:
                                            y_offset = 0
                                        
                                        if "center" in text_position:
                                            ha = 'center'
                                        elif "left" in text_position:
                                            ha = 'left'
                                        elif "right" in text_position:
                                            ha = 'right'
                                        else:
                                            ha = 'center'
                                        
                                        # Add text label
                                        label_color = data_label_color or '#000000'
                                        ax1.text(x, y + y_offset, str(text), 
                                               ha=ha, va='bottom' if y_offset > 0 else 'top',
                                               fontsize=data_label_font_size or 10,
                                               color=label_color,
                                               fontweight='bold',
                                               bbox=dict(boxstyle="round,pad=0.3", 
                                                        facecolor='white', 
                                                        alpha=0.9,
                                                        edgecolor='gray',
                                                        linewidth=0.5),
                                               zorder=5)
                        # Set axis labels, title, and legend for area chart (only once after all series)
                        if i == len(series_data) - 1 and mpl_chart_type == "fill_between":  # Only for area charts
                            # Set axis labels
                            x_axis_title = chart_meta.get("x_label", chart_config.get("x_axis_title", ""))
                            y_axis_title = chart_meta.get("primary_y_label", chart_config.get("primary_y_label", ""))
                            if x_axis_title:
                                # Use the actual distance value directly, not divided by 10
                                x_labelpad = x_axis_label_distance if x_axis_label_distance is not None else 5.0
                                # Make the distance effect much more pronounced by multiplying the value
                                x_labelpad = (x_axis_label_distance * 10) if x_axis_label_distance is not None else 50.0
                                ax1.set_xlabel(x_axis_title, fontsize=font_size or 12, color=font_color if font_color else 'black',
                                             fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None,
                                             labelpad=x_labelpad)
                            if y_axis_title:
                                # Use the actual distance value directly, not divided by 10
                                y_labelpad = y_axis_label_distance if y_axis_label_distance is not None else 5.0
                                # Make the distance effect much more pronounced by multiplying the value
                                y_labelpad = (y_axis_label_distance * 10) if y_axis_label_distance is not None else 50.0
                                ax1.set_ylabel(y_axis_title, fontsize=font_size or 12, color=font_color if font_color else 'black',
                                             fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None,
                                             labelpad=y_labelpad)
                             
                             # Set chart title
                            if title:
                                 ax1.set_title(title, fontsize=font_size or 14, weight='bold', pad=20,
                                             color=font_color if font_color else 'black',
                                             fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None)
                             
                             # Set legend
                            show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                            if show_legend:
                                 legend_position = chart_meta.get("legend_position", "top")
                                 legend_font_size = chart_meta.get("legend_font_size", 10)
                                 
                                 if legend_position == "bottom":
                                     ax1.legend(loc='lower center', bbox_to_anchor=(0.5, -0.15), fontsize=legend_font_size)
                                 elif legend_position == "top":
                                     ax1.legend(loc='upper center', bbox_to_anchor=(0.5, 1.02), fontsize=legend_font_size)
                                 else:
                                     ax1.legend(loc='best', fontsize=legend_font_size)

                            # Add data labels if text is provided
                            if text_labels and len(text_labels) == len(y_vals):
                                # ... existing data label code ...
                                ax1.annotate(text, (x, y + y_offset), ha=ha, va='center',
                                            fontsize=data_label_font_size or 10,
                                            color=label_color,
                                            fontweight='bold',
                                            bbox=dict(boxstyle="round,pad=0.3", 
                                                    facecolor='white', 
                                                    alpha=0.9,
                                                    edgecolor='gray',
                                                    linewidth=0.5),
                                            zorder=5)

                            # Set axis labels, title, and legend for area chart (only once after all series)
                            if i == len(series_data) - 1:  # Only set once after all series are processed
                                # Set axis labels
                                x_axis_title = chart_meta.get("x_label", chart_config.get("x_axis_title", ""))
                                y_axis_title = chart_meta.get("primary_y_label", chart_config.get("primary_y_label", ""))
                                if x_axis_title:
                                    # Use the actual distance value directly, not divided by 10
                                    x_labelpad = x_axis_label_distance if x_axis_label_distance is not None else 5.0
                                    # Make the distance effect much more pronounced by multiplying the value
                                    x_labelpad = (x_axis_label_distance * 10) if x_axis_label_distance is not None else 50.0
                                    ax1.set_xlabel(x_axis_title, fontsize=font_size or 12, color=font_color if font_color else 'black',
                                                fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None,
                                                labelpad=x_labelpad)
                                if y_axis_title:
                                    # Use the actual distance value directly, not divided by 10
                                    y_labelpad = y_axis_label_distance if y_axis_label_distance is not None else 5.0
                                    # Make the distance effect much more pronounced by multiplying the value
                                    y_labelpad = (y_axis_label_distance * 10) if y_axis_label_distance is not None else 50.0
                                    ax1.set_ylabel(y_axis_title, fontsize=font_size or 12, color=font_color if font_color else 'black',
                                                fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None,
                                                labelpad=y_labelpad)

                            # Set chart title
                            if title:
                                ax1.set_title(title, fontsize=font_size or 14, weight='bold', pad=20,
                                            color=font_color if font_color else 'black',
                                            fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None)
                            
                            # Set legend
                            show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                            if show_legend:
                                legend_position = chart_meta.get("legend_position", "top")
                                legend_font_size = chart_meta.get("legend_font_size", 10)
                                
                                if legend_position == "bottom":
                                    ax1.legend(loc='lower center', bbox_to_anchor=(0.5, -0.15), fontsize=legend_font_size)
                                elif legend_position == "top":
                                    ax1.legend(loc='upper center', bbox_to_anchor=(0.5, 1.02), fontsize=legend_font_size)
                                else:
                                    ax1.legend(loc='best', fontsize=legend_font_size)
                                    
                        elif mpl_chart_type == "hist":
                            # Histogram
                            ax1.hist(y_vals, bins=10, label=label, color=color, alpha=0.7)
                            
                        elif mpl_chart_type == "boxplot":
                            # Box plot
                            ax1.boxplot(y_vals, labels=[label], patch_artist=True)
                            if color:
                                ax1.findobj(plt.matplotlib.patches.Patch)[-1].set_facecolor(color)
                                
                        elif mpl_chart_type == "violinplot":
                            # Violin plot
                            ax1.violinplot(y_vals, positions=[i])
                            
                        elif mpl_chart_type == "imshow":
                            # Enhanced heatmap implementation
                            heatmap_data = series.get("z", series.get("values", []))
                            # current_app.logger.debug(f"Heatmap data found: {heatmap_data}")
                            
                            # Ensure we have valid heatmap data
                            if not heatmap_data or len(heatmap_data) == 0:
                                #current_app.logger.warning(f"⚠️ No heatmap data found in series: {series}")
                                # Create a default heatmap for testing
                                heatmap_data = [[1, 0, 1, 1], [1, 1, 1, 0], [0, 1, 1, 1]]
                                # current_app.logger.debug(f"Using default heatmap data: {heatmap_data}")
                            
                            # Ensure heatmap_data is a 2D array
                            if isinstance(heatmap_data[0], (int, float)):
                                # Flat list - reshape based on x_axis length
                                cols = len(x_values) if x_values else 4
                                rows = len(heatmap_data) // cols
                                if len(heatmap_data) % cols != 0:
                                    rows += 1
                                # Pad with zeros if needed
                                padded_data = heatmap_data + [0] * (rows * cols - len(heatmap_data))
                                heatmap_data = [padded_data[i:i+cols] for i in range(0, len(padded_data), cols)]
                                # current_app.logger.debug(f"Reshaped heatmap data: {heatmap_data}")
                                # current_app.logger.debug(f"X values: {x_values}")
                                # current_app.logger.debug(f"Cols: {cols}, Rows: {rows}")
                            
                            # Get colorscale from series or use default
                            colorscale = series.get("colorscale", "RdYlGn")
                            # current_app.logger.debug(f"Using colorscale: {colorscale}")
                            
                            # Create heatmap with enhanced styling
                            im = ax1.imshow(heatmap_data, cmap=colorscale, aspect='auto', 
                                           interpolation='nearest', alpha=0.8)
                            
                            # Axis labels from config
                            x_axis_title = chart_meta.get("x_label", chart_config.get("x_axis_title", ""))
                            y_axis_title = chart_meta.get("primary_y_label", chart_config.get("primary_y_label", ""))
                            if x_axis_title:
                                # Apply x_axis_label_distance for heatmaps
                                x_labelpad = x_axis_label_distance if x_axis_label_distance else 5.0  # Use the value directly
                                ax1.set_xlabel(x_axis_title, fontsize=font_size or 12, color=font_color if font_color else None,
                                               fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None,
                                               labelpad=x_labelpad)
                            if y_axis_title:
                                # Apply y_axis_label_distance for heatmaps
                                y_labelpad = y_axis_label_distance if y_axis_label_distance else 5.0  # Use the value directly
                                ax1.set_ylabel(y_axis_title, fontsize=font_size or 12, color=font_color if font_color else None,
                                               fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None,
                                               labelpad=y_labelpad)
                            
                            # Set axis labels
                            if x_values:
                                ax1.set_xticks(range(len(x_values)))
                                ax1.set_xticklabels(x_values, rotation=0, ha='center')
                            
                            # Create employee labels (Y-axis) - use series labels if available
                            y_labels = series.get("y", [])
                            if not y_labels:
                                num_employees = len(heatmap_data)
                                y_labels = [f"Employee {i+1}" for i in range(num_employees)]
                            
                            ax1.set_yticks(range(len(y_labels)))
                            ax1.set_yticklabels(y_labels)
                            
                            # Respect show_gridlines for heatmap explicitly (override style defaults)
                            if not chart_meta.get("show_gridlines", False):
                                try:
                                    ax1.grid(visible=False, which='both', axis='both')
                                except TypeError:
                                    # Fallback for older Matplotlib
                                    ax1.grid(False)
                                ax1.set_axisbelow(True)
                            
                            # Add colorbar if showscale is enabled
                            showscale = series.get("showscale", True)
                            if showscale:
                                cbar = plt.colorbar(im, ax=ax1, shrink=0.8)
                                cbar.set_label(series.get("name", "Value"), rotation=270, labelpad=15,
                                                    fontsize=chart_meta.get("legend_font_size", 10))
                            
                            # Add text annotations on heatmap cells
                            for i in range(len(heatmap_data)):
                                for j in range(len(heatmap_data[0])):
                                    value = heatmap_data[i][j]
                                    # Determine text color based on background
                                    if colorscale == "RdYlGn":
                                        text_color = 'white' if value > 0.5 else 'black'
                                    else:
                                        text_color = 'white' if value == 1 else 'black'
                                    
                                    # Show actual value or custom text
                                    if "text" in series and i < len(series["text"]) and j < len(series["text"][i]):
                                        cell_text = str(series["text"][i][j])
                                    else:
                                        cell_text = str(value)
                                    
                                    ax1.text(j, i, cell_text, 
                                           ha='center', va='center', color=text_color, 
                                           fontweight='bold', fontsize=10)
                            
                            # Set title
                            ax1.set_title(title, fontsize=font_size or 14, weight='bold', pad=20,
                                          color=font_color if font_color else None,
                                          fontname=chart_meta.get("font_family") if chart_meta.get("font_family") else None)
                        elif mpl_chart_type == "contour":
                            # Contour plot (simplified)
                            if len(y_vals) > 0:
                                # Create a simple 2D array for contour
                                contour_data = [y_vals] if len(y_vals) > 0 else [[0]]
                                ax1.contour(contour_data)
                                
                        else:
                            # Fallback to scatter for unknown types
                            #current_app.logger.warning(f"⚠️ Unknown matplotlib chart type '{series_type}', falling back to scatter")
                            ax1.scatter(x_values, y_vals, label=label, color=color, alpha=0.7)

                    # Add data labels to Matplotlib chart if enabled (skip for area charts as they have custom label handling)
                    if show_data_labels and (data_label_format or value_format or data_label_font_size or data_label_color) and chart_type != "area":
                        # current_app.logger.debug(f"Adding data labels to Matplotlib chart")
                        for i, series in enumerate(series_data):
                            series_type = series.get("type", "bar").lower()
                            y_vals = series.get("values")
                            value_range = series.get("value_range")
                            if value_range:
                                if isinstance(value_range, list):
                                    y_vals = value_range
                                else:
                                    y_vals = extract_values_from_range(value_range)
                            
                            if y_vals:
                                # Determine format to use
                                format_to_use = value_format if value_format else data_label_format
                                if not format_to_use:
                                    format_to_use = ".1f"  # Default format
                                
                                # Add data labels based on chart type
                                if series_type == "bar":
                                    for j, val in enumerate(y_vals):
                                        if j < len(x_values):
                                            # Format the value
                                            try:
                                                if format_to_use == ".1f":
                                                    formatted_val = f"{float(val):.1f}"
                                                elif format_to_use == ".0f":
                                                    formatted_val = f"{float(val):.0f}"
                                                elif format_to_use == ".0%":
                                                    formatted_val = f"{float(val):.0%}"
                                                else:
                                                    formatted_val = f"{float(val):{format_to_use}}"
                                            except:
                                                formatted_val = str(val)
                                            
                                            # Add text label on top of bar (ENHANCED VERSION)
                                            label_color = data_label_color or '#000000'
                                            ax1.text(j, val, formatted_val, 
                                                    ha='center', va='bottom', 
                                                    fontsize=data_label_font_size or int(title_fontsize * 0.9),  # Improved scaling
                                                    color=label_color,
                                                    fontweight='bold',
                                                    bbox=dict(boxstyle="round,pad=0.2", facecolor='white', alpha=0.9))
                                            # current_app.logger.debug(f"Added bar data label with color: {label_color}")
                                
                                elif series_type == "line":
                                    for j, val in enumerate(y_vals):
                                        if j < len(x_values):
                                            # For line charts, use secondary y-axis format if available
                                            line_format = secondary_y_axis_format if secondary_y_axis_format else format_to_use
                                            if not line_format:
                                                line_format = ".1f"  # Default format
                                            
                                            # Format the value
                                            try:
                                                if line_format == ".1f":
                                                    formatted_val = f"{float(val):.1f}"
                                                elif line_format == ".0f":
                                                    formatted_val = f"{float(val):.0f}"
                                                elif line_format == ".0%":
                                                    formatted_val = f"{float(val):.0%}"
                                                elif line_format == ".1%":
                                                    formatted_val = f"{float(val):.1%}"
                                                else:
                                                    formatted_val = f"{float(val):{line_format}}"
                                            except:
                                                formatted_val = str(val)
                                            
                                            # Add text label above line point (ENHANCED VERSION)
                                            label_color = data_label_color or '#000000'
                                            ax2.text(j, val, formatted_val, 
                                                    ha='center', va='bottom', 
                                                    fontsize=data_label_font_size or int(title_fontsize * 0.9),  # Improved scaling
                                                    color=label_color,
                                                    fontweight='bold',
                                                    bbox=dict(boxstyle="round,pad=0.2", facecolor='white', alpha=0.9))
                                            # current_app.logger.debug(f"Added line data label with color: {label_color}")

                                            # Set labels and styling
                        if chart_type != "pie" and chart_type != "area":
                            # Improved font size scaling for Matplotlib (moved outside condition)
                            title_fontsize = font_size or 52  # Increased default title size
                            label_fontsize = int((font_size or 52) * 0.9)  # Increased relative size for labels
                            
                            # Apply axis label distances using labelpad parameter
                            x_labelpad = x_axis_label_distance if x_axis_label_distance else 5.0  # Use the value directly
                            y_labelpad = y_axis_label_distance if y_axis_label_distance else 5.0  # Use the value directly
                            # Make the distance effect much more pronounced by multiplying the values
                            x_labelpad = (x_axis_label_distance * 10) if x_axis_label_distance is not None else 50.0
                            y_labelpad = (y_axis_label_distance * 10) if y_axis_label_distance is not None else 50.0
                            
                            # Debug logging
                            current_app.logger.info(f"🔍 Axis Label Distance Debug - X: {x_axis_label_distance} → {x_labelpad}, Y: {y_axis_label_distance} → {y_labelpad}")
                            
                            ax1.set_xlabel(chart_meta.get("x_label", chart_config.get("x_axis_title", "X")), 
                                         fontsize=label_fontsize, color=font_color, labelpad=x_labelpad)
                            ax1.set_ylabel(chart_meta.get("primary_y_label", chart_config.get("primary_y_label", "Primary Y")), 
                                         fontsize=label_fontsize, color=font_color, labelpad=y_labelpad)
                            if "secondary_y_label" in chart_meta or "secondary_y_label" in chart_config:
                                ax2.set_ylabel(chart_meta.get("secondary_y_label", chart_config.get("secondary_y_label", "Secondary Y")), 
                                             fontsize=label_fontsize, color=font_color, labelpad=y_labelpad)

                            # Apply axis scale type if provided
                            xaxis_type_cfg = chart_meta.get("xaxis_type")
                            yaxis_type_cfg = chart_meta.get("yaxis_type")
                            if xaxis_type_cfg in ("log", "log10"):
                                ax1.set_xscale('log')
                            if yaxis_type_cfg in ("log", "log10"):
                                ax1.set_yscale('log')

                        # Determine if this is a bubble chart early
                        is_bubble_chart = any(series.get("type", "").lower() == "bubble" for series in series_data)

                        # Set axis tick font size and control tick visibility
                        tick_fontsize = axis_tick_font_size or int(title_fontsize * 0.8)  # Improved tick size calculation
                        if axis_tick_font_size:
                            ax1.tick_params(axis='x', labelsize=axis_tick_font_size, colors=font_color)  # Removed rotation
                            ax1.tick_params(axis='y', labelsize=axis_tick_font_size, colors=font_color)
                            if ax2 and not is_bubble_chart:
                                ax2.tick_params(axis='y', labelsize=axis_tick_font_size, colors=font_color)
                        else:
                            ax1.tick_params(axis='x', labelsize=tick_fontsize, colors=font_color)  # Removed rotation
                            ax1.tick_params(axis='y', labelsize=tick_fontsize, colors=font_color)
                            if ax2 and not is_bubble_chart:
                                ax2.tick_params(axis='y', labelsize=tick_fontsize, colors=font_color)
                        
                        # Tick mark control for Matplotlib
                        if show_x_ticks is not None or show_y_ticks is not None:
                            if show_x_ticks is not None:
                                if not show_x_ticks:
                                    ax1.tick_params(axis='x', length=0)  # Hide tick marks
                                    ax1.set_xticklabels([])  # Hide tick labels
                                else:
                                    ax1.tick_params(axis='x', length=5)  # Show tick marks
                            if show_y_ticks is not None:
                                if not show_y_ticks:
                                    ax1.tick_params(axis='y', length=0)  # Hide tick marks
                                    ax1.set_yticklabels([])  # Hide tick labels
                                    if ax2 and not is_bubble_chart:
                                        ax2.tick_params(axis='y', length=0)  # Hide secondary y-axis tick marks
                                        ax2.set_yticklabels([])  # Hide secondary y-axis tick labels
                                else:
                                    ax1.tick_params(axis='y', length=5)  # Show tick marks
                                    if ax2 and not is_bubble_chart:
                                     ax2.tick_params(axis='y', length=5)  # Show secondary y-axis tick marks
                        
                        # Apply X-axis label distance using tick parameters
                        # if x_axis_label_distance:
                        #     # Use tick label padding to control distance
                        #     ax1.tick_params(axis='x', pad=5)  # Fixed: Use small default padding
                        #     # current_app.logger.debug(f"Applied X-axis tick padding: {x_axis_label_distance}")
                        
                        # Apply Y-axis label distance using tick parameters
                        # if y_axis_label_distance:
                        #     # Use tick label padding to control distance
                        #     ax1.tick_params(axis='y', pad=5)  # Fixed: Use small default padding
                        #     # current_app.logger.debug(f"Applied Y-axis tick padding: {y_axis_label_distance}")
                        
                        # Apply secondary y-axis formatting for Matplotlib
                        if ax2 and not is_bubble_chart:
                            if secondary_y_axis_format:
                                from matplotlib.ticker import FuncFormatter
                                def percentage_formatter(x, pos):
                                    return f'{x:.0%}'
                                ax2.yaxis.set_major_formatter(FuncFormatter(percentage_formatter))
                        if secondary_y_axis_min_max:
                            # Handle "auto" value for secondary y-axis min/max
                            if secondary_y_axis_min_max == "auto":
                                # Don't set range for auto - let Matplotlib auto-scale
                                pass
                            elif isinstance(secondary_y_axis_min_max, list) and len(secondary_y_axis_min_max) == 2:
                                ax2.set_ylim(secondary_y_axis_min_max)
                            else:
                                current_app.logger.warning(f"Invalid secondary_y_axis_min_max format: {secondary_y_axis_min_max}")
                        elif is_bubble_chart:
                            # current_app.logger.info(f"🎈 Skipping secondary Y-axis formatting for bubble chart")
                            # Explicitly hide secondary Y-axis for bubble charts
                            if ax2:
                                ax2.set_visible(False)
                                # current_app.logger.info(f"🎈 Secondary Y-axis hidden for bubble chart")
                        
                        # Gridlines
                        # current_app.logger.debug(f"Gridlines setting: {show_gridlines}")
                        if show_gridlines:
                            # Map gridline styles to valid Matplotlib linestyles
                            matplotlib_linestyle_map = {
                                "solid": "-",
                                "dashed": "--", 
                                "dash": "--",  # Map 'dash' to '--'
                                "dashdot": "-.",
                                "dotted": ":",
                                "dot": ":",
                                "dotdash": "-."
                            }
                            mapped_linestyle = matplotlib_linestyle_map.get(gridline_style, "--")
                            # Show both horizontal and vertical gridlines
                            ax1.grid(True, linestyle=mapped_linestyle, color=gridline_color if gridline_color else '#ccc', alpha=0.6, axis='both')
                            if ax2 and not is_bubble_chart:
                                ax2.grid(True, linestyle=mapped_linestyle, color=gridline_color if gridline_color else '#ccc', alpha=0.6, axis='both')
                        else:
                            ax1.grid(False)
                            if ax2:
                             ax2.grid(False)
                        
                        # Apply primary y-axis formatting for Matplotlib
                        if axis_tick_format:
                            from matplotlib.ticker import FuncFormatter
                            if "$" in axis_tick_format:
                                def currency_formatter(x, pos):
                                    return f'${x:,.0f}'
                                ax1.yaxis.set_major_formatter(FuncFormatter(currency_formatter))
                        if y_axis_min_max:
                            # current_app.logger.debug(f"Setting Matplotlib Y-axis range to: {y_axis_min_max}")
                            # Ensure the range is properly applied
                            if isinstance(y_axis_min_max, list) and len(y_axis_min_max) == 2:
                                # Check if the provided range is appropriate for the data
                                all_y_values = []
                                for series in series_data:
                                    y_vals = series.get("values", [])
                                    if y_vals:
                                        all_y_values.extend(y_vals)
                                
                                if all_y_values:
                                    data_min = min(all_y_values)
                                    data_max = max(all_y_values)
                                    data_range = data_max - data_min
                                    
                                    # Check if the provided range is reasonable (within 100x of data range for currency formatting)
                                    provided_range = y_axis_min_max[1] - y_axis_min_max[0]
                                    if provided_range > data_range * 100 and not axis_tick_format:
                                        # Auto-calculate appropriate range only if no special formatting is requested
                                        padding = data_range * 0.1  # 10% padding
                                        auto_min = max(0, data_min - padding)
                                        auto_max = data_max + padding
                                        ax1.set_ylim(auto_min, auto_max)
                                        # current_app.logger.debug(f"Auto-adjusted Y-axis range from {y_axis_min_max} to {[auto_min, auto_max]} (data range: {data_min}-{data_max})")
                                    else:
                                        # Use provided range (especially for currency formatting)
                                        ax1.set_ylim(y_axis_min_max[0], y_axis_min_max[1])
                                        # current_app.logger.debug(f"Applied Y-axis range: {y_axis_min_max[0]} to {y_axis_min_max[1]}")
                                else:
                                    # No data available, use provided range
                                    ax1.set_ylim(y_axis_min_max[0], y_axis_min_max[1])
                                    # current_app.logger.debug(f"Applied Y-axis range: {y_axis_min_max[0]} to {y_axis_min_max[1]}")
                            else:
                                current_app.logger.warning(f"Invalid Y-axis range format: {y_axis_min_max}")
                        
                        # Set title with improved font size
                        ax1.set_title(title, fontsize=title_fontsize, weight='bold', pad=20)
                        
                        # Legend
                        show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                        
                        legend_loc = 'best'
                        if show_legend:
                            handles = []
                            labels_list = []
                            # Use explicit proxy if present
                            proxy = getattr(ax1, "_bubble_proxy", None)
                            if proxy is not None:
                                handles.append(proxy)
                                labels_list.append(proxy.get_label())
                            else:
                                # Prefer stored bubble handles/labels
                                handles = getattr(ax1, "_bubble_handles", []) or []
                                labels_list = getattr(ax1, "_bubble_labels", []) or []
                                if not handles:
                                    handles, labels_list = ax1.get_legend_handles_labels()

                            # Combine with secondary axis if present
                            if 'ax2' in locals() and ax2 is not None:
                                h2, l2 = ax2.get_legend_handles_labels()
                                handles += h2
                                labels_list += l2

                            if legend_position:
                                loc_map = {
                                    "top": "upper center",
                                    "bottom": "lower center", 
                                    "left": "center left",
                                    "right": "center right",
                                }
                                legend_loc = loc_map.get(legend_position, 'best')
                            
                            if legend_position == "bottom":
                                ax1.legend(handles, labels_list, loc='lower center', bbox_to_anchor=(0.5, -0.15), fontsize=legend_font_size)
                            elif legend_position == "top":
                                ax1.legend(handles, labels_list, loc='upper center', bbox_to_anchor=(0.5, 1.02), fontsize=legend_font_size)
                            else:
                                ax1.legend(handles, labels_list, loc=legend_loc, fontsize=legend_font_size)
                                
                        # Set title with proper font attributes (ENHANCED VERSION)

                # Apply axis label distances for Matplotlib (skip heatmaps as they're handled specifically)
                # current_app.logger.debug(f"X-axis label distance: {x_axis_label_distance}")
                # current_app.logger.debug(f"Y-axis label distance: {y_axis_label_distance}")
                
                if (x_axis_label_distance or y_axis_label_distance) and chart_type != "heatmap":
                    # Get current subplot parameters
                    current_bottom = fig_mpl.subplotpars.bottom
                    current_left = fig_mpl.subplotpars.left
                    
                    if x_axis_label_distance:
                        # Convert the distance to a fraction of the figure height
                        # Higher x_axis_label_distance values will push labels further down
                        adjustment = x_axis_label_distance / 500.0  # Increased conversion factor for more visible effect
                        fig_mpl.subplots_adjust(bottom=current_bottom - adjustment)
                        # current_app.logger.debug(f"Applied X-axis adjustment: {adjustment}")
                    
                    if y_axis_label_distance:
                        # Convert the distance to a fraction of the figure width
                        adjustment = y_axis_label_distance / 500.0  # Increased conversion factor
                        fig_mpl.subplots_adjust(left=current_left - adjustment)
                        # current_app.logger.debug(f"Applied Y-axis adjustment: {adjustment}")
                
                # Apply tight_layout but preserve manual adjustments
                fig_mpl.tight_layout()
                
                # Re-apply manual adjustments after tight_layout with larger effect (skip heatmaps)
                if (x_axis_label_distance or y_axis_label_distance) and chart_type != "heatmap":
                    if x_axis_label_distance:
                        adjustment = x_axis_label_distance / 300.0  # Even larger effect after tight_layout
                        fig_mpl.subplots_adjust(bottom=fig_mpl.subplotpars.bottom - adjustment)
                        # Also re-apply x-axis label padding and nudge its position downward
                        try:
                            x_labelpad = x_axis_label_distance / 10.0 if x_axis_label_distance else 5.0
                            ax1.xaxis.labelpad = x_labelpad
                            # Move label further down in axes coordinates
                            x_label_nudge = -0.10 - (x_axis_label_distance / 1200.0)
                            ax1.xaxis.set_label_coords(0.5, x_label_nudge)
                        except Exception:
                            pass
                        # current_app.logger.debug(f"Re-applied X-axis adjustment: {adjustment}")
                    
                    if y_axis_label_distance:
                        adjustment = y_axis_label_distance / 300.0  # Even larger effect after tight_layout
                        fig_mpl.subplots_adjust(left=fig_mpl.subplotpars.left - adjustment)
                        # Ensure the y-label itself moves away from the axis ticks
                        try:
                            # Re-apply labelpad explicitly after tight_layout
                            y_labelpad = y_axis_label_distance / 10.0 if y_axis_label_distance else 5.0
                            ax1.yaxis.labelpad = y_labelpad
                            # Additionally, nudge the label position in axes coordinates for a clearer visual effect
                            # Negative x moves it further left; scale factor tuned for visibility
                            coord_nudge = -0.02 - (y_axis_label_distance / 1200.0)
                            ax1.yaxis.set_label_coords(coord_nudge, 0.5)
                        except Exception:
                            pass
                        # current_app.logger.debug(f"Re-applied Y-axis adjustment: {adjustment}")

                # Add annotations if specified
                if annotations:
                    # current_app.logger.debug(f"Adding {len(annotations)} annotations to chart")
                    
                    # For heatmaps, extract x and y values from series data
                    if chart_type == "heatmap" and series_data:
                        heatmap_x_values = series_data[0].get("x", []) if series_data else []
                        heatmap_y_values = series_data[0].get("y", []) if series_data else []
                    else:
                        heatmap_x_values = []
                        heatmap_y_values = []
                    
                    for annotation in annotations:
                        text = annotation.get("text", "")
                        x_value = annotation.get("x_value", 0)
                        y_value = annotation.get("y_value", 0)
                        
                        # Find x position based on x_value
                        x_pos = 0
                        if chart_type == "heatmap" and heatmap_x_values:
                            # For heatmaps, look in the heatmap x values
                            if x_value in heatmap_x_values:
                                x_pos = heatmap_x_values.index(x_value)
                        elif x_values and x_value in x_values:
                            # For other charts, look in the standard x_values
                            x_pos = x_values.index(x_value)
                        elif isinstance(x_value, str) and x_values:
                            # Try to find string value in x_values
                            try:
                                x_pos = x_values.index(x_value)
                            except ValueError:
                                # If not found, use first position
                                x_pos = 0
                                # Annotation x_value not found in x_values (skipping)
                        elif isinstance(x_value, (int, float)) and x_values:
                            # For numeric x_value, use the actual value for bubble charts
                            if chart_type == "bubble":
                                x_pos = x_value  # Use actual x value for bubble charts
                            else:
                                # For other charts, find closest position
                                try:
                                    x_pos = int(x_value)
                                    if x_pos >= len(x_values):
                                        x_pos = len(x_values) - 1
                                except (ValueError, TypeError):
                                    x_pos = 0
                        
                        # Find y position based on y_value
                        y_pos = y_value
                        if chart_type == "heatmap" and heatmap_y_values:
                            # For heatmaps, look in the heatmap y values
                            if y_value in heatmap_y_values:
                                y_pos = heatmap_y_values.index(y_value)
                        
                        # Auto-adjust Y position if it's outside the data range
                        all_y_values = []
                        for series in series_data:
                            y_vals = series.get("values", [])
                            if y_vals:
                                all_y_values.extend(y_vals)
                        
                        if all_y_values:
                            data_min = min(all_y_values)
                            data_max = max(all_y_values)
                            data_range = data_max - data_min
                            
                            # If annotation Y value is way outside data range, adjust it
                            if y_pos > data_max * 2 or y_pos < data_min * 0.5:
                                # Position annotation above the highest data point
                                y_pos = data_max + (data_range * 0.1)  # 10% above max
                                # current_app.logger.debug(f"Auto-adjusted annotation Y position from {y_value} to {y_pos} (data range: {data_min}-{data_max})")
                        
                        # Add annotation text with better positioning for bubble charts
                        if chart_type == "bubble":
                            # For bubble charts, position annotation above the bubble
                            annotation_y_offset = 2000  # Fixed offset for bubble charts
                            ax1.annotate(text, xy=(x_pos, y_pos), xytext=(x_pos, y_pos + annotation_y_offset),
                                       arrowprops=dict(arrowstyle='->', color='red', lw=1.5),
                                       fontsize=12, color='red', weight='bold',
                                       ha='center', va='bottom',
                                       bbox=dict(boxstyle="round,pad=0.3", facecolor='white', alpha=0.9, edgecolor='red'))
                        else:
                            # For other charts, use original positioning
                         ax1.annotate(text, xy=(x_pos, y_pos), xytext=(x_pos, y_pos + (data_range * 0.05) if 'data_range' in locals() else y_pos + 50),
                                   arrowprops=dict(arrowstyle='->', color='red'),
                                   fontsize=12, color='red', weight='bold',
                                   ha='center', va='bottom')
                        # current_app.logger.debug(f"Added annotation: '{text}' at position ({x_pos}, {y_pos})")

                # Apply margin settings to Matplotlib (FIXED VERSION)
                if margin:
                    # current_app.logger.debug(f"Applying margin to Matplotlib: {margin}")
                    # Use figure padding instead of subplot adjustments
                    # Convert margin values to fractions of figure size (more appropriate scaling)
                    # Use a smaller conversion factor to avoid invalid subplot parameters
                    conversion_factor = 1000.0  # Larger denominator for smaller values
                    left_margin_fraction = margin.get("l", 0) / conversion_factor
                    right_margin_fraction = margin.get("r", 0) / conversion_factor
                    top_margin_fraction = margin.get("t", 0) / conversion_factor
                    bottom_margin_fraction = margin.get("b", 0) / conversion_factor
                    
                    # Calculate subplot parameters with validation
                    left_pos = max(0.1, left_margin_fraction)  # Minimum 0.1 to avoid edge
                    right_pos = min(0.9, 1.0 - right_margin_fraction)  # Maximum 0.9 to avoid edge
                    top_pos = min(0.9, 1.0 - top_margin_fraction)  # Maximum 0.9 to avoid edge
                    bottom_pos = max(0.1, bottom_margin_fraction)  # Minimum 0.1 to avoid edge
                    
                    # Ensure left < right and bottom < top
                    if left_pos >= right_pos:
                        left_pos = 0.1
                        right_pos = 0.9
                        #current_app.logger.warning(f"Invalid margin: left >= right, using default values")
                    if bottom_pos >= top_pos:
                        bottom_pos = 0.1
                        top_pos = 0.9
                        #current_app.logger.warning(f"Invalid margin: bottom >= top, using default values")
                    
                    # Apply margins using figure padding
                    fig_mpl.subplots_adjust(
                        left=left_pos,
                        right=right_pos,
                        top=top_pos,
                        bottom=bottom_pos
                    )
                    # current_app.logger.debug(f"Applied margins (fractions): left={left_pos}, right={right_pos}, top={top_pos}, bottom={bottom_pos}")
                
                # Adjust layout to accommodate legend position
                if show_legend and legend_position == "bottom":
                    # Add extra space at bottom for legend
                    fig_mpl.subplots_adjust(bottom=fig_mpl.subplotpars.bottom + 0.15)
                    # current_app.logger.debug("Added extra bottom space for legend")

                tmpfile = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                # Use different bbox_inches parameter based on legend position
                if show_legend and legend_position == "bottom":
                    # For bottom legend, use 'tight' but with extra padding
                    plt.savefig(tmpfile.name, bbox_inches='tight', pad_inches=0.3, dpi=200)
                    # current_app.logger.debug(f"Saved Matplotlib chart with bottom legend using extra padding")
                else:
                    # For other positions, use standard tight layout
                    plt.savefig(tmpfile.name, bbox_inches='tight', dpi=200)
                                        # Only log legend position if legend_loc is defined (for charts that have legends)
                    if 'legend_loc' in locals():
                        # Chart saved with legend
                        pass
                    else:
                        # current_app.logger.debug(f"Saved Matplotlib chart without legend")
                        plt.close(fig_mpl)

                return tmpfile.name

            except Exception as e:
                import traceback
                
                # Create user-friendly error message
                error_type = type(e).__name__
                error_msg = str(e)
                
                # Simplify common error messages
                if "JSONDecodeError" in error_type:
                    user_message = "Invalid JSON format in chart configuration"
                elif "KeyError" in error_type:
                    user_message = f"Missing required field: {error_msg}"
                elif "ValueError" in error_type:
                    user_message = f"Invalid value: {error_msg}"
                elif "IndexError" in error_type:
                    user_message = "Data array is empty or has wrong dimensions"
                elif "TypeError" in error_type:
                    user_message = f"Wrong data type: {error_msg}"
                elif "FileNotFoundError" in error_type:
                    user_message = "Excel file or sheet not found"
                elif "openpyxl" in error_msg.lower():
                    user_message = "Excel file format error - check if file is corrupted"
                elif "pandas" in error_msg.lower():
                    user_message = "Data reading error - check Excel file format"
                else:
                    user_message = error_msg
                
                error_details = {
                    "chart_tag": chart_tag,
                    "error_type": error_type,
                    "user_message": user_message,
                    "technical_message": error_msg,
                    "chart_type": chart_type if 'chart_type' in locals() else "unknown",
                    "data_points": len(series_data) if 'series_data' in locals() else 0,
                    "timestamp": datetime.utcnow().isoformat()
                }
                
                # Simple console logging
                current_app.logger.error(f"❌ Chart '{chart_tag}' failed: {user_message}")
                
                # Store error details for frontend (project-specific)
                if not hasattr(current_app, 'chart_errors'):
                    current_app.chart_errors = {}
                if project_id not in current_app.chart_errors:
                    current_app.chart_errors[project_id] = {}
                current_app.chart_errors[project_id][chart_tag] = error_details
                
                # Also store a simplified version for report generation errors
                if not hasattr(current_app, 'report_generation_errors'):
                    current_app.report_generation_errors = {}
                if project_id not in current_app.report_generation_errors:
                    current_app.report_generation_errors[project_id] = {}
                current_app.report_generation_errors[project_id][chart_tag] = {
                    "error": user_message,
                    "chart_type": chart_type if 'chart_type' in locals() else "unknown",
                    "timestamp": datetime.utcnow().isoformat()
                }
                
                return None

        # Insert charts into paragraphs
        chart_errors = []
        
        # COMPREHENSIVE TEXT REPLACEMENT - Process ALL document elements
        process_entire_document()
        
        # Process charts in paragraphs
        for para_idx, para in enumerate(doc.paragraphs):
            full_text = ''.join(run.text for run in para.runs)
            chart_placeholders = re.findall(r"\$\{(section\d+_chart)\}", full_text, flags=re.IGNORECASE)
            for tag in chart_placeholders:
                if tag.lower() in chart_attr_map:
                    try:
                        chart_img = generate_chart({}, tag)
                        if chart_img:
                            para.text = re.sub(rf"\$\{{{tag}\}}", "", para.text, flags=re.IGNORECASE)
                            para.add_run().add_picture(chart_img, width=Inches(5.5))
                        else:
                            # Chart generation failed, add error placeholder
                            error_msg = f"[Chart failed: {tag}]"
                            para.text = re.sub(rf"\$\{{{tag}\}}", error_msg, para.text, flags=re.IGNORECASE)
                            
                            # Get the specific error from chart_errors if available
                            specific_error = "Chart could not be generated"
                            if hasattr(current_app, 'chart_errors') and project_id in current_app.chart_errors:
                                if tag in current_app.chart_errors[project_id]:
                                    specific_error = current_app.chart_errors[project_id][tag].get('user_message', specific_error)
                            elif hasattr(current_app, 'report_generation_errors') and project_id in current_app.report_generation_errors:
                                if tag in current_app.report_generation_errors[project_id]:
                                    specific_error = current_app.report_generation_errors[project_id][tag].get('error', specific_error)
                            
                            chart_errors.append({
                                "tag": tag,
                                "error": specific_error
                            })
                    except Exception as e:
                        current_app.logger.error(f"⚠️ Failed to insert chart for tag {tag}: {e}")
                        error_msg = f"[Chart failed: {tag}]"
                        para.text = re.sub(rf"\$\{{{tag}\}}", error_msg, para.text, flags=re.IGNORECASE)
                        chart_errors.append({
                            "tag": tag,
                            "error": f"Chart insertion failed: {str(e)}"
                        })

        # Insert charts into tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        full_text = ''.join(run.text for run in para.runs)
                        chart_placeholders = re.findall(r"\$\{(section\d+_chart)\}", full_text, flags=re.IGNORECASE)
                        for tag in chart_placeholders:
                            if tag.lower() in chart_attr_map:
                                try:
                                    chart_img = generate_chart({}, tag)
                                    if chart_img:
                                        para.text = re.sub(rf"\$\{{{tag}\}}", "", para.text, flags=re.IGNORECASE)
                                        para.add_run().add_picture(chart_img, width=Inches(5.5))
                                    else:
                                        # Chart generation failed, add error placeholder
                                        error_msg = f"[Chart failed: {tag}]"
                                        para.text = re.sub(rf"\$\{{{tag}\}}", error_msg, para.text, flags=re.IGNORECASE)
                                        
                                        # Get the specific error from chart_errors if available
                                        specific_error = "Chart could not be generated"
                                        if hasattr(current_app, 'chart_errors') and project_id in current_app.chart_errors:
                                            if tag in current_app.chart_errors[project_id]:
                                                specific_error = current_app.chart_errors[project_id][tag].get('user_message', specific_error)
                                        elif hasattr(current_app, 'report_generation_errors') and project_id in current_app.report_generation_errors:
                                            if tag in current_app.report_generation_errors[project_id]:
                                                specific_error = current_app.report_generation_errors[project_id][tag].get('error', specific_error)
                                        
                                        chart_errors.append({
                                            "tag": tag,
                                            "error": specific_error
                                        })
                                except Exception as e:
                                    current_app.logger.error(f"⚠️ Failed to insert chart in table for tag {tag}: {e}")
                                    error_msg = f"[Chart failed: {tag}]"
                                    para.text = re.sub(rf"\$\{{{tag}\}}", error_msg, para.text, flags=re.IGNORECASE)
                                    chart_errors.append({
                                        "tag": tag,
                                        "error": f"Chart insertion failed: {str(e)}"
                                    })

        # Save report to temporary location
        import tempfile
        temp_dir = tempfile.mkdtemp()
        output_path = os.path.join(temp_dir, f'output_report_{project_id}.docx')
        doc.save(output_path)
        current_app.logger.info(f"✅ Report generated successfully")
        
        # Store chart errors for this report generation
        if not hasattr(current_app, 'report_errors'):
            current_app.report_errors = {}
        current_app.report_errors[project_id] = {
            "chart_errors": chart_errors,
            "generated_at": datetime.utcnow().isoformat()
        }
        
        # Clear old chart errors for this project when new report is generated
        if hasattr(current_app, 'chart_errors') and project_id in current_app.chart_errors:
            current_app.chart_errors[project_id] = {}
        
        return output_path

    except Exception as e:
        current_app.logger.error(f"❌ Failed to generate report: {e}")
        import traceback
        current_app.logger.error(traceback.format_exc())
        return None

# Helper function no longer needed - files are now stored in database

@projects_bp.route('/api/projects', methods=['GET'])
@login_required
def get_projects():
    # Access MongoDB via current_app.mongo.db
    projects = list(current_app.mongo.db.projects.find({'user_id': current_user.get_id()}))
    for project in projects:
        project['id'] = str(project['_id'])
        del project['_id']
        # Remove binary file content to prevent JSON serialization error
        if 'file_content' in project:
            del project['file_content']
    return jsonify({'projects': projects})

@projects_bp.route('/api/projects', methods=['POST'])
@login_required
def create_project():
    name = request.form.get('name')
    description = request.form.get('description')
    file = request.files.get('file') 

    if not name or not description:
        return jsonify({'error': 'Missing required fields (name or description)'}), 400

    file_name = None
    file_content = None
    if file:
        if not allowed_file(file.filename): 
            return jsonify({'error': 'File type not allowed'}), 400
        file_name = secure_filename(file.filename)
        file_content = file.read()  # Read file content into memory

    project = {
        'name': name,
        'description': description,
        'user_id': current_user.get_id(),
        'file_name': file_name,
        'file_content': file_content,  # Store file content in database
        'created_at': datetime.utcnow().isoformat() 
    }
    # Access MongoDB via current_app.mongo.db
    project_id = current_app.mongo.db.projects.insert_one(project).inserted_id
    
    # Create a copy for JSON response without binary content
    project_response = {
        'id': str(project_id),
        'name': project['name'],
        'description': project['description'],
        'user_id': project['user_id'],
        'file_name': project['file_name'],
        'created_at': project['created_at']
    }

    return jsonify({'message': 'Project created successfully', 'project': project_response}), 201

@projects_bp.route('/api/projects/<project_id>/upload_report', methods=['POST'])
@login_required
def upload_report(project_id):
    current_app.logger.info(f"📤 Upload request received for project: {project_id}")
    
    if 'report_file' not in request.files:
        current_app.logger.error(f"❌ No report_file in request.files: {list(request.files.keys())}")
        return jsonify({'error': 'No report file provided'}), 400

    report_file = request.files['report_file']
    current_app.logger.info(f"📁 File received: {report_file.filename}")

    if report_file.filename == '':
        current_app.logger.error(f"❌ Empty filename")
        return jsonify({'error': 'No selected report file'}), 400

    if not allowed_report_file(report_file.filename):
        current_app.logger.error(f"❌ File type not allowed: {report_file.filename}")
        return jsonify({'error': 'Report file type not allowed. Only .xlsx or .csv are accepted.'}), 400

    try:
        project_id_obj = ObjectId(project_id)
        current_app.logger.info(f"✅ Valid project ID: {project_id}")
    except Exception as e:
        current_app.logger.error(f"❌ Invalid project ID: {project_id}, error: {e}")
        return jsonify({'error': 'Invalid project ID'}), 400

    project = current_app.mongo.db.projects.find_one({'_id': project_id_obj, 'user_id': current_user.get_id()})
    if not project:
        current_app.logger.error(f"❌ Project not found or unauthorized: {project_id}")
        return jsonify({'error': 'Project not found or unauthorized'}), 404

    current_app.logger.info(f"✅ Project found: {project.get('name', 'Unknown')}")

    # Handle both old (file_path) and new (file_name/file_content) project formats
    template_file_name = project.get('file_name')
    template_file_content = project.get('file_content')
    
    # Backward compatibility: if new format not found, try old format
    if not template_file_name or not template_file_content:
        old_file_path = project.get('file_path')
        if old_file_path:
            # Convert old format to new format
            abs_file_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), old_file_path)
            if os.path.exists(abs_file_path):
                try:
                    with open(abs_file_path, 'rb') as f:
                        template_file_content = f.read()
                    template_file_name = os.path.basename(old_file_path)
                    
                    # Update project to new format
                    current_app.mongo.db.projects.update_one(
                        {'_id': project_id_obj},
                        {'$set': {
                            'file_name': template_file_name,
                            'file_content': template_file_content
                        }}
                    )
                    current_app.logger.info(f"🔄 Migrated project from old format to new format")
                except Exception as e:
                    current_app.logger.error(f"❌ Failed to migrate old project format: {e}")
                    return jsonify({'error': 'Failed to load template file. Please re-upload the template.'}), 400
            else:
                current_app.logger.error(f"❌ Old template file not found: {abs_file_path}")
                return jsonify({'error': 'Template file not found. Please re-upload the template.'}), 400
        else:
            current_app.logger.error(f"❌ No template file found in project")
            return jsonify({'error': 'Word template file not found for this project. Please upload it during project creation.'}), 400
    
    current_app.logger.info(f"📄 Template file name: {template_file_name}")
    
    # Create temporary file from database content
    temp_template_dir = tempfile.mkdtemp()
    temp_template_path = os.path.join(temp_template_dir, template_file_name)
    with open(temp_template_path, 'wb') as f:
        f.write(template_file_content)
    current_app.logger.info(f"📄 Temporary template created: {temp_template_path}")
    
    # Save the uploaded report data file temporarily
    report_data_filename = secure_filename(report_file.filename)
    temp_dir = tempfile.mkdtemp()
    temp_report_data_path = os.path.join(temp_dir, report_data_filename)
    report_file.save(temp_report_data_path)

    # Clear any existing errors for this project before starting new generation
    if hasattr(current_app, 'chart_errors') and project_id in current_app.chart_errors:
        current_app.chart_errors[project_id] = {}

    # Generate the report
    current_app.logger.info(f"🔄 Starting report generation...")
    generated_report_path = _generate_report(project_id, temp_template_path, temp_report_data_path)
    
    # Clean up the temporary files and directories
    import shutil
    shutil.rmtree(temp_dir)
    shutil.rmtree(temp_template_dir)
    current_app.logger.info(f"🧹 Temporary files cleaned up")

    if generated_report_path:
        current_app.logger.info(f"✅ Report generated successfully: {generated_report_path}")
        # Update project with generated report path
        current_app.mongo.db.projects.update_one(
            {'_id': project_id_obj},
            {'$set': {'generated_report_path': generated_report_path, 'report_generated_at': datetime.utcnow().isoformat()}}
        )
        return jsonify({'message': 'Report generated successfully', 'report_path': generated_report_path}), 200
    else:
        current_app.logger.error(f"❌ Report generation failed")
        return jsonify({'error': 'Failed to generate report'}), 500

@projects_bp.route('/api/reports/<project_id>/download', methods=['GET'])
@login_required
def download_report(project_id):
    try:
        project_id_obj = ObjectId(project_id)
    except:
        return jsonify({'error': 'Invalid project ID'}), 400

    project = current_app.mongo.db.projects.find_one({'_id': project_id_obj, 'user_id': current_user.get_id()})
    if not project:
        return jsonify({'error': 'Project not found or unauthorized'}), 404

    generated_report_path = project.get('generated_report_path')
    if not generated_report_path or not os.path.exists(generated_report_path):
        return jsonify({'error': 'Generated report not found for this project'}), 404

    # Send file and clean up after download
    try:
        response = send_file(generated_report_path, as_attachment=True)
        
        # Clean up the temporary file after sending
        def cleanup_after_response(response):
            try:
                if os.path.exists(generated_report_path):
                    os.remove(generated_report_path)
                    # Also remove the temp directory if it's empty
                    temp_dir = os.path.dirname(generated_report_path)
                    if os.path.exists(temp_dir) and not os.listdir(temp_dir):
                        os.rmdir(temp_dir)
            except Exception as e:
                current_app.logger.warning(f"⚠️ Failed to cleanup temporary report file: {e}")
            return response
        
        response.call_on_close(lambda: cleanup_after_response(response))
        return response
        
    except Exception as e:
        # Clean up on error too
        try:
            if os.path.exists(generated_report_path):
                os.remove(generated_report_path)
                temp_dir = os.path.dirname(generated_report_path)
                if os.path.exists(temp_dir) and not os.listdir(temp_dir):
                    os.rmdir(temp_dir)
        except:
            pass
        raise e

@projects_bp.route('/api/reports/<chart_filename>/download_html', methods=['GET'])
@login_required
def download_chart_html(chart_filename):
    # Charts are embedded in Word documents, no separate HTML files needed
    return jsonify({'error': 'Chart HTML files are not available - charts are embedded in Word documents'}), 404

@projects_bp.route('/api/reports/batch_reports_<project_id>.zip', methods=['GET'])
@login_required
def download_batch_reports(project_id):
    """Download batch reports ZIP file"""
    try:
        # Validate project ID
        project_id_obj = ObjectId(project_id)
        
        # Check if project exists and belongs to user
        project = current_app.mongo.db.projects.find_one({'_id': project_id_obj, 'user_id': current_user.get_id()})
        if not project:
            return jsonify({'error': 'Project not found or unauthorized'}), 404
        
        # Construct the ZIP file path
        zip_filename = f'batch_reports_{project_id}.zip'
        zip_path = os.path.join(tempfile.gettempdir(), zip_filename)
        
        # Check if file exists
        if not os.path.exists(zip_path):
            return jsonify({'error': 'Batch reports file not found. Please regenerate the reports.'}), 404
        
        # Send file
        def cleanup_after_response(response):
            try:
                # Clean up the ZIP file after sending
                if os.path.exists(zip_path):
                    os.remove(zip_path)
                    current_app.logger.info(f"Cleaned up batch reports ZIP: {zip_path}")
            except Exception as e:
                current_app.logger.warning(f"Failed to clean up batch reports ZIP {zip_path}: {e}")
            return response
        
        response = send_file(
            zip_path,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )
        
        # Add cleanup callback
        response.call_on_close(lambda: cleanup_after_response(response))
        
        current_app.logger.info(f"Downloading batch reports ZIP: {zip_filename}")
        return response
        
    except Exception as e:
        current_app.logger.error(f"Error downloading batch reports: {e}")
        return jsonify({'error': 'Failed to download batch reports'}), 500

@projects_bp.route('/api/projects/<project_id>/chart_errors', methods=['GET'])
@login_required
def get_chart_errors(project_id):
    try:
        project_id_obj = ObjectId(project_id)
    except:
        return jsonify({'error': 'Invalid project ID'}), 400

    project = current_app.mongo.db.projects.find_one({'_id': project_id_obj, 'user_id': current_user.get_id()})
    if not project:
        return jsonify({'error': 'Project not found or unauthorized'}), 404

    # Get chart errors for this project
    chart_errors = getattr(current_app, 'chart_errors', {}).get(project_id, {})
    report_errors = getattr(current_app, 'report_errors', {}).get(project_id, {})
    report_generation_errors = getattr(current_app, 'report_generation_errors', {}).get(project_id, {})
    
    # Combine both types of errors
    all_errors = {
        "chart_generation_errors": chart_errors,
        "report_generation_errors": report_errors.get("chart_errors", []),
        "report_generation_errors_detailed": report_generation_errors,
        "report_generated_at": report_errors.get("generated_at")
    }
    
    return jsonify(all_errors)

@projects_bp.route('/api/projects/<project_id>/clear_errors', methods=['POST'])
@login_required
def clear_project_errors(project_id):
    try:
        project_id_obj = ObjectId(project_id)
    except:
        return jsonify({'error': 'Invalid project ID'}), 400

    project = current_app.mongo.db.projects.find_one({'_id': project_id_obj, 'user_id': current_user.get_id()})
    if not project:
        return jsonify({'error': 'Project not found or unauthorized'}), 404

    # Clear chart errors for this project
    if hasattr(current_app, 'chart_errors') and project_id in current_app.chart_errors:
        current_app.chart_errors[project_id] = {}
    
    # Clear report errors for this project
    if hasattr(current_app, 'report_errors') and project_id in current_app.report_errors:
        current_app.report_errors[project_id] = {}
    
    return jsonify({'message': 'Project errors cleared successfully'})

@projects_bp.route('/api/projects/<project_id>/upload_zip', methods=['POST'])
@login_required
def upload_zip_and_generate_reports(project_id):
    if 'zip_file' not in request.files:
        return jsonify({'error': 'No zip file provided'}), 400

    zip_file = request.files['zip_file']
    if not zip_file.filename.endswith('.zip'):
        return jsonify({'error': 'Only .zip files are allowed'}), 400

    # Clear any existing errors for this project before starting new generation
    if hasattr(current_app, 'chart_errors') and project_id in current_app.chart_errors:
        current_app.chart_errors[project_id] = {}

    # Prepare temp directories
    temp_dir = tempfile.mkdtemp()
    extracted_dir = os.path.join(temp_dir, 'extracted')
    os.makedirs(extracted_dir, exist_ok=True)

    # Save and extract ZIP
    zip_path = os.path.join(temp_dir, secure_filename(zip_file.filename))
    zip_file.save(zip_path)
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extracted_dir)

    # Find all Excel files
    excel_files = [os.path.join(extracted_dir, f) for f in os.listdir(extracted_dir) if f.endswith('.xlsx') or f.endswith('.xls')]

    # Prepare temporary output folders
    output_folder_name = os.path.join(temp_dir, 'reports_by_name')
    output_folder_code = os.path.join(temp_dir, 'reports_by_code')
    os.makedirs(output_folder_name, exist_ok=True)
    os.makedirs(output_folder_code, exist_ok=True)

    generated_files = []
    total_files = len(excel_files)
    current_app.logger.info(f"Starting batch processing of {total_files} Excel files")
    
    for idx, excel_path in enumerate(excel_files, 1):
        # Extract report name and code from Excel file
        report_name, report_code = extract_report_info_from_excel(excel_path)
        
        current_app.logger.info(f"Processing file {idx}/{total_files}: {report_name} (Code: {report_code})")

        # Generate report
        project = current_app.mongo.db.projects.find_one({'_id': ObjectId(project_id)})
        
        # Handle both old and new project formats
        template_file_name = project.get('file_name')
        template_file_content = project.get('file_content')
        
        # Backward compatibility: if new format not found, try old format
        if not template_file_name or not template_file_content:
            old_file_path = project.get('file_path')
            if old_file_path:
                abs_file_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), old_file_path)
                if os.path.exists(abs_file_path):
                    with open(abs_file_path, 'rb') as f:
                        template_file_content = f.read()
                    template_file_name = os.path.basename(old_file_path)
                else:
                    current_app.logger.error(f"❌ Old template file not found for batch processing: {abs_file_path}")
                    continue
            else:
                current_app.logger.error(f"❌ No template file found for batch processing")
                continue
        
        # Create temporary template file
        temp_template_dir = tempfile.mkdtemp()
        temp_template_path = os.path.join(temp_template_dir, template_file_name)
        with open(temp_template_path, 'wb') as f:
            f.write(template_file_content)
        
        output_path = _generate_report(f"{project_id}_{idx}", temp_template_path, excel_path)
        
        # Clean up temporary template
        shutil.rmtree(temp_template_dir)
        if output_path:
            # Save in both folders with proper naming
            name_file_path = os.path.join(output_folder_name, f"{report_name}.docx")
            code_file_path = os.path.join(output_folder_code, f"{report_code}.docx")
            
            shutil.copy(output_path, name_file_path)
            shutil.copy(output_path, code_file_path)
            
            generated_files.append({'name': report_name, 'code': report_code})
            current_app.logger.info(f"✅ Successfully generated report {idx}/{total_files}: {report_name} -> {report_code}")
        else:
            current_app.logger.error(f"❌ Failed to generate report {idx}/{total_files}: {report_name}")
        
        # Log progress
        current_app.logger.info(f"Progress: {idx}/{total_files} reports processed")

    # Create zip file in temporary location with both folder structures
    zip_output_path = os.path.join(temp_dir, f'batch_reports_{project_id}.zip')
    with zipfile.ZipFile(zip_output_path, 'w') as zipf:
        # Add files from both folders to maintain the folder structure
        for file_info in generated_files:
            # Add file from reports_by_name folder
            name_file_path = os.path.join(output_folder_name, f"{file_info['name']}.docx")
            if os.path.exists(name_file_path):
                zipf.write(name_file_path, arcname=f"reports_by_name/{file_info['name']}.docx")
            
            # Add file from reports_by_code folder
            code_file_path = os.path.join(output_folder_code, f"{file_info['code']}.docx")
            if os.path.exists(code_file_path):
                zipf.write(code_file_path, arcname=f"reports_by_code/{file_info['code']}.docx")

    # Move zip to a temporary location that will be cleaned up after download
    final_zip_path = os.path.join(tempfile.gettempdir(), f'batch_reports_{project_id}.zip')
    shutil.move(zip_output_path, final_zip_path)
    
    # Clean up temp directory
    shutil.rmtree(temp_dir)

    current_app.logger.info(f"Batch processing complete. Generated {len(generated_files)} out of {total_files} reports")

    return jsonify({
        'message': f'Generated {len(generated_files)} out of {total_files} reports.',
        'download_zip': final_zip_path,
        'reports': generated_files,
        'total_files': total_files,
        'processed_files': len(generated_files),
        'success_rate': f"{len(generated_files)}/{total_files}"
    })

@projects_bp.route('/api/projects/<project_id>', methods=['PUT'])
@login_required
def update_project(project_id):
    try:
        project_id_obj = ObjectId(project_id)
    except Exception as e:
        return jsonify({'error': 'Invalid project ID'}), 400

    # Check if project exists and belongs to user
    project = current_app.mongo.db.projects.find_one({'_id': project_id_obj, 'user_id': current_user.get_id()})
    if not project:
        return jsonify({'error': 'Project not found or unauthorized'}), 404

    # Get form data
    name = request.form.get('name')
    description = request.form.get('description')
    file = request.files.get('file')

    # Validate required fields
    if not name or not description:
        return jsonify({'error': 'Missing required fields (name or description)'}), 400

    # Prepare update data
    update_data = {
        'name': name,
        'description': description,
        'updated_at': datetime.utcnow().isoformat()
    }

    # Handle file upload if provided
    if file:
        if not allowed_file(file.filename):
            return jsonify({'error': 'File type not allowed. Only .doc or .docx files are accepted.'}), 400
        
        # Read new file content
        file_name = secure_filename(file.filename)
        file_content = file.read()
        update_data['file_name'] = file_name
        update_data['file_content'] = file_content

    # Update project in database
    result = current_app.mongo.db.projects.update_one(
        {'_id': project_id_obj, 'user_id': current_user.get_id()},
        {'$set': update_data}
    )

    if result.modified_count == 0:
        return jsonify({'error': 'Failed to update project'}), 500

    # Get updated project
    updated_project = current_app.mongo.db.projects.find_one({'_id': project_id_obj})
    updated_project['id'] = str(updated_project['_id'])
    del updated_project['_id']

    return jsonify({'message': 'Project updated successfully', 'project': updated_project})

@projects_bp.route('/api/projects/<project_id>', methods=['DELETE'])
@login_required
def delete_project(project_id):
    try:
        project_id_obj = ObjectId(project_id)
    except Exception as e:
        return jsonify({'error': 'Invalid project ID'}), 400

    # Check if project exists and belongs to user
    project = current_app.mongo.db.projects.find_one({'_id': project_id_obj, 'user_id': current_user.get_id()})
    if not project:
        return jsonify({'error': 'Project not found or unauthorized'}), 404

    # File content is stored in database, no file system cleanup needed

    # Delete project from database
    result = current_app.mongo.db.projects.delete_one({'_id': project_id_obj, 'user_id': current_user.get_id()})

    if result.deleted_count == 0:
        return jsonify({'error': 'Failed to delete project'}), 500

    return jsonify({'message': 'Project deleted successfully'})

@projects_bp.route('/api/projects/<project_id>', methods=['GET'])
@login_required
def get_project(project_id):
    try:
        project_id_obj = ObjectId(project_id)
    except Exception as e:
        return jsonify({'error': 'Invalid project ID'}), 400

    # Check if project exists and belongs to user
    project = current_app.mongo.db.projects.find_one({'_id': project_id_obj, 'user_id': current_user.get_id()})
    if not project:
        return jsonify({'error': 'Project not found or unauthorized'}), 404

    project['id'] = str(project['_id'])
    del project['_id']
    
    return jsonify({'project': project})
