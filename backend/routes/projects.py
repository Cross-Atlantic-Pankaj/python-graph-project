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
import openpyxl
import tempfile
import re
import zipfile
import shutil
import plotly.graph_objects as go
from plotly.subplots import make_subplots


import re

# Define a constant for the section1_chart attribut

# Define a custom UPLOAD_FOLDER for this blueprint or ensure it's globally configured
UPLOAD_FOLDER = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'uploads')
# Ensure the upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif', 'csv', 'xlsx', 'docx'}
ALLOWED_REPORT_EXTENSIONS = {'csv', 'xlsx'}

projects_bp = Blueprint('projects', __name__)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def allowed_report_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_REPORT_EXTENSIONS

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
    # Bar chart for breakdown
    fig.add_trace(go.Bar(
        x=other_labels,
        y=other_values,
        marker_color=other_colors,
        text=[f"{v}{value_format}" for v in other_values],
        textposition="auto",
        name="Breakdown of 'Other'"
    ), row=1, col=2)
    fig.update_layout(
        title_text=title,
        showlegend=False,
        height=500,
        width=900
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
        current_app.logger.debug(f"Generating report for project: {project_id}")

        df = pd.read_excel(data_file_path, sheet_name=0)  # Use first sheet
        df.columns = df.columns.str.strip().str.replace(" ", "_").str.replace("__", "_")

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
        for _, row in df.iterrows():
            chart_tag = row.get("Chart_Tag")
            if not isinstance(chart_tag, str) or not chart_tag:
                continue
            section_prefix = chart_tag.replace('_chart', '').lower()

            # Process all columns systematically
            for col in df.columns:
                col_lower = col.lower().strip()
                
                # Handle Chart Data columns (for values)
                if col_lower.startswith("chart_data_y"):
                    year = col.replace("Chart_Data_", "").replace("chart_data_", "")
                    key = f"{section_prefix}_{year.lower()}"
                    value = row[col]
                    if pd.notna(value) and str(value).strip():
                        flat_data_map[key] = str(value).strip()
                    else:
                        current_app.logger.warning(f"⚠️ Empty value for {col} -> {key}")
                
                # Handle Growth columns (for growth rates)
                elif col_lower.startswith("growth_y"):
                    year = col.replace("Growth_", "").replace("growth_", "")
                    key = f"{section_prefix}_{year.lower()}_kpi2"
                    value = row[col]
                    if pd.notna(value) and str(value).strip():
                        # Format percentage values properly
                        try:
                            # Convert to float first to handle any numeric format
                            float_val = float(value)
                            # If the value is between 0 and 1, it's likely a decimal percentage
                            if 0 <= float_val <= 1:
                                # Convert decimal to percentage (e.g., 0.2 -> 20%)
                                percentage_val = f"{float_val * 100:.1f}%"
                                flat_data_map[key] = percentage_val
                            else:
                                # If it's already a larger number, assume it's already a percentage
                                flat_data_map[key] = f"{float_val:.1f}%"
                        except (ValueError, TypeError):
                            # If conversion fails, use the original value as string
                            flat_data_map[key] = str(value).strip()
                    else:
                        current_app.logger.warning(f"⚠️ Empty value for {col} -> {key}")

            # Handle CAGR
            if pd.notna(row.get("Chart_Data_CAGR")):
                key = f"{section_prefix}_cgrp"
                value = row["Chart_Data_CAGR"]
                if str(value).strip():
                    # Format CAGR percentage values properly
                    try:
                        # Convert to float first to handle any numeric format
                        float_val = float(value)
                        # If the value is between 0 and 1, it's likely a decimal percentage
                        if 0 <= float_val <= 1:
                            # Convert decimal to percentage (e.g., 0.105 -> 10.5%)
                            percentage_val = f"{float_val * 100:.1f}%"
                            flat_data_map[key] = percentage_val
                        else:
                            # If it's already a larger number, assume it's already a percentage
                            flat_data_map[key] = f"{float_val:.1f}%"
                    except (ValueError, TypeError):
                        # If conversion fails, use the original value as string
                        flat_data_map[key] = str(value).strip()
                else:
                    current_app.logger.warning(f"⚠️ Empty value for Chart_Data_CAGR -> {key}")

        # Ensure all keys are lowercase
        flat_data_map = {k.lower(): v for k, v in flat_data_map.items()}
        
        # Data mapping completed silently

        doc = Document(template_path)

        def replace_text_in_paragraph(paragraph):
            # First, try to replace placeholders that are contained within single runs
            for run in paragraph.runs:
                original_text = run.text
                matches = re.findall(r"\$\{(.*?)\}", run.text)
                for match in matches:
                    key_lower = match.lower().strip()
                    val = flat_data_map.get(key_lower) or text_map.get(key_lower)
                    if val:
                        pattern = re.compile(re.escape(f"${{{match}}}"), re.IGNORECASE)
                        run.text = pattern.sub(val, run.text)
                    else:
                        current_app.logger.warning(f"⚠️ No value found for placeholder ${{{match}}} (key: {key_lower})")
            
            # Then, handle placeholders that might be split across multiple runs
            # Get the full paragraph text to find all placeholders
            full_text = paragraph.text
            all_matches = re.findall(r"\$\{(.*?)\}", full_text)
            
            for match in all_matches:
                key_lower = match.lower().strip()
                val = flat_data_map.get(key_lower) or text_map.get(key_lower)
                if val:
                    placeholder = f"${{{match}}}"
                    if placeholder in full_text:
                        # Clear all runs and recreate with replacement
                        for run in paragraph.runs:
                            run.text = ""
                        # Set the first run to the replaced text
                        if paragraph.runs:
                            paragraph.runs[0].text = full_text.replace(placeholder, str(val))
                        else:
                            # If no runs exist, create one
                            paragraph.add_run(full_text.replace(placeholder, str(val)))
                        break  # Only process one placeholder at a time to avoid conflicts

        def replace_text_in_tables():
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            replace_text_in_paragraph(para)
            
            # Table processing completed silently


                            
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

                # --- Extract custom fields from chart_config ---
                bar_colors = chart_config.get("bar_colors")
                bar_width = chart_config.get("bar_width")
                orientation = chart_config.get("orientation")
                bar_border_color = chart_config.get("bar_border_color")
                bar_border_width = chart_config.get("bar_border_width")
                font_family = chart_config.get("font_family") or chart_meta.get("font_family")
                font_size = chart_config.get("font_size") or chart_meta.get("font_size")
                font_color = chart_config.get("font_color") or chart_meta.get("font_color")
                legend_position = chart_config.get("legend_position")
                legend_font_size = chart_config.get("legend_font_size")
                show_gridlines = chart_config.get("show_gridlines") if "show_gridlines" in chart_config else chart_meta.get("show_gridlines")
                # Ensure show_gridlines is a boolean
                if isinstance(show_gridlines, str):
                    show_gridlines = show_gridlines.strip().lower() == "true"
                elif show_gridlines is None:
                    show_gridlines = True  # Default to showing gridlines if not specified
                gridline_color = chart_config.get("gridline_color")
                gridline_style = chart_config.get("gridline_style")
                chart_background = chart_config.get("chart_background")
                plot_background = chart_config.get("plot_background")
                data_label_format = chart_config.get("data_label_format") or chart_meta.get("data_label_format")
                data_label_font_size = chart_config.get("data_label_font_size") or chart_meta.get("data_label_font_size")
                data_label_color = chart_config.get("data_label_color") or chart_meta.get("data_label_color")
                axis_tick_format = chart_config.get("axis_tick_format")
                y_axis_min_max = chart_config.get("y_axis_min_max") or chart_meta.get("y_axis_min_max")
                current_app.logger.debug(f"Y-axis min/max from config: {y_axis_min_max}")
                secondary_y_axis_format = chart_config.get("secondary_y_axis_format") or chart_meta.get("secondary_y_axis_format")
                secondary_y_axis_min_max = chart_config.get("secondary_y_axis_min_max") or chart_meta.get("secondary_y_axis_min_max")
                sort_order = chart_config.get("sort_order")
                data_grouping = chart_config.get("data_grouping")
                annotations = chart_config.get("annotations", [])
                axis_tick_font_size = chart_config.get("axis_tick_font_size")
                
                # --- Extract margin settings ---
                margin = chart_config.get("margin") or chart_meta.get("margin")
                x_axis_label_distance = chart_config.get("x_axis_label_distance") or chart_meta.get("x_axis_label_distance")
                y_axis_label_distance = chart_config.get("y_axis_label_distance") or chart_meta.get("y_axis_label_distance")
                axis_tick_distance = chart_config.get("axis_tick_distance") or chart_meta.get("axis_tick_distance")
                figsize = chart_config.get("figsize") or chart_meta.get("figsize")

                # --- Excel range extraction helpers ---
                def extract_excel_range(sheet, cell_range):
                    # cell_range: e.g., 'E23:E29' or 'AA20:AA23'
                    from openpyxl.utils import range_boundaries
                    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
                    values = []
                    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                        for cell in row:
                            values.append(cell.value)
                    return values

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
                                    current_app.logger.warning(f"⚠️ Failed to extract {k} from {v}: {e}")
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
                    x_values = series_meta.get("x_axis", [])
                
                colors = series_meta.get("colors", [])

                # --- Plotly interactive chart generation ---
                fig = go.Figure()

                # --- Bar of Pie chart special handling ---
                if chart_type in ["bar of pie", "bar_of_pie"]:
                    # Re-fetch after extraction to ensure we have the extracted lists
                    other_labels = chart_meta.get("other_labels", [])
                    other_values = chart_meta.get("other_values", [])
                    other_colors = chart_meta.get("other_colors", [])
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
                            hovertemplate=f"<b>{label}</b><br>%{{label}}: %{{value}}{value_format}<extra></extra>"
                        ))
                
                # Handle stacked column, area, and other multi-series charts
                else:
                    for i, series in enumerate(series_data):
                        label = series.get("name", f"Series {i+1}")
                        series_type = series.get("type", "bar").lower()
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
                                    "textinfo": "label+percent+value" if chart_meta.get("data_labels", True) else "none",
                                    "textposition": "outside",
                                    "hole": 0.4 if series_type == "donut" else 0.0
                                }
                                
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
                                
                            else:
                                # Generic handling for other chart types
                                fig.add_trace(plotly_chart_class(**trace_kwargs,
                                    hovertemplate=f"<b>{label}</b><br>Value: %{{y}}<extra></extra>"
                                ))
                        else:
                            # Fallback to scatter if chart type not recognized
                            current_app.logger.warning(f"⚠️ Unknown chart type '{series_type}', falling back to scatter")
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
                if show_legend:
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
                    if legend_font_size:
                        layout_updates.setdefault("legend", {})["font"] = {"size": legend_font_size}
                else:
                    layout_updates["showlegend"] = False
                
                # Bar mode for stacked charts
                if chart_type == "stacked_column":
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
                            current_app.logger.debug(f"Setting Y-axis range to: {y_axis_min_max}")
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
                        dash_map = {"solid": "solid", "dot": "dot", "dash": "dash"}
                        dash_style = dash_map.get(gridline_style, "solid")
                        layout_updates["xaxis"] = layout_updates.get("xaxis", {})
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["xaxis"]["griddash"] = dash_style
                        layout_updates["yaxis"]["griddash"] = dash_style
                
                # Data labels (for bar/line traces)
                show_data_labels = chart_meta.get("data_labels", True)
                value_format = chart_meta.get("value_format", "")
                
                # Enable data labels if any data label settings are provided
                if data_label_format or data_label_font_size or data_label_color:
                    show_data_labels = True
                
                # Debug logging for data labels
                current_app.logger.debug(f"Data labels enabled: {show_data_labels}")
                current_app.logger.debug(f"Data label format: {data_label_format}")
                current_app.logger.debug(f"Data label font size: {data_label_font_size}")
                current_app.logger.debug(f"Data label color: {data_label_color}")
                current_app.logger.debug(f"Value format: {value_format}")
                current_app.logger.debug(f"Chart config keys: {list(chart_config.keys())}")
                current_app.logger.debug(f"Chart meta keys: {list(chart_meta.keys())}")
                
                if show_data_labels and (data_label_format or value_format or data_label_font_size or data_label_color):
                    current_app.logger.debug(f"Processing {len(fig.data)} traces for data labels")
                    for i, trace in enumerate(fig.data):
                        current_app.logger.debug(f"Trace {i}: type={trace.type}, mode={getattr(trace, 'mode', 'N/A')}")
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

                fig.update_layout(**layout_updates)

                # --- Matplotlib static chart for DOCX ---
                if chart_type == "pie":
                    # Check if this is an expanded pie chart
                    expanded_segment = chart_meta.get("expanded_segment")
                    
                    if expanded_segment and len(series_data) == 1:
                        # Create subplot for expanded pie chart
                        fig_mpl, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 8), dpi=200)
                        
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
                        
                        # Create pie chart
                        wedges, texts, autotexts = ax1.pie(values, labels=labels, autopct='%1.1f%%', 
                                                          colors=color, startangle=90)
                        
                        # Style the text
                        for autotext in autotexts:
                            autotext.set_color('white')
                            autotext.set_fontweight('bold')
                        
                        ax1.set_title(title, fontsize=font_size or 14, weight='bold', pad=20)
                        
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
                        fig_mpl, ax = plt.subplots(figsize=(10, 8), dpi=200)
                        
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
                            
                            # Create pie chart
                            wedges, texts, autotexts = ax.pie(values, labels=labels, autopct='%1.1f%%', 
                                                             colors=color, startangle=90)
                            
                            # Style the text
                            for autotext in autotexts:
                                autotext.set_color('white')
                                autotext.set_fontweight('bold')
                            
                            ax.set_title(title, fontsize=font_size or 14, weight='bold', pad=20)
                        
                elif chart_type in ["bar of pie", "bar_of_pie"]:
                    # Matplotlib version of bar of pie
                    fig_mpl, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 8), dpi=200)
                    
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
                    # Bar chart
                    if other_colors:
                        ax2.bar(other_labels, other_values, color=other_colors, alpha=0.7)
                    else:
                        ax2.bar(other_labels, other_values, color=colors if colors else None, alpha=0.7)
                    ax2.set_title("Breakdown of 'Other'", fontsize=font_size or 12, weight='bold')
                    ax2.set_ylabel("Value")
                    for i, v in enumerate(other_values):
                        ax2.text(i, v, f"{v}", ha='center', va='bottom', fontweight='bold')
                
                else:
                    # Bar, line, area charts
                    fig_mpl, ax1 = plt.subplots(figsize=(10, 6), dpi=200)
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
                        
                        mpl_chart_type = chart_type_mapping_mpl.get(series_type, "scatter")
                        
                        if mpl_chart_type == "bar":
                            if chart_type == "stacked_column":
                                # For stacked column, use bottom parameter
                                if i == 0:
                                    ax1.bar(x_values, y_vals, label=label, color=color, alpha=0.7)
                                    bottom_vals = y_vals
                                else:
                                    ax1.bar(x_values, y_vals, bottom=bottom_vals, label=label, color=color, alpha=0.7)
                                    bottom_vals = [sum(x) for x in zip(bottom_vals, y_vals)]
                            else:
                                if isinstance(color, list):
                                    for j, val in enumerate(y_vals):
                                        bar_color = color[j % len(color)]
                                        ax1.bar(x_values[j], val, color=bar_color, alpha=0.7)
                                else:
                                    ax1.bar(x_values, y_vals, label=label, color=color, alpha=0.7)
                                    
                        elif mpl_chart_type == "barh":
                            # Horizontal bar chart
                            if isinstance(color, list):
                                for j, val in enumerate(y_vals):
                                    bar_color = color[j % len(color)]
                                    ax1.barh(x_values[j], val, color=bar_color, alpha=0.7)
                            else:
                                ax1.barh(x_values, y_vals, label=label, color=color, alpha=0.7)
                                
                        elif mpl_chart_type == "plot":
                            # Line chart
                            marker = 'o' if series_type == "scatter_line" else None
                            ax2.plot(x_values, y_vals, label=label, color=color, marker=marker, linewidth=2)
                            
                        elif mpl_chart_type == "scatter":
                            # Scatter plot
                            if series_type == "bubble" and "size" in series:
                                sizes = series.get("size", [20] * len(y_vals))
                                ax1.scatter(x_values, y_vals, s=sizes, label=label, color=color, alpha=0.7)
                            else:
                                ax1.scatter(x_values, y_vals, label=label, color=color, alpha=0.7)
                                
                        elif mpl_chart_type == "fill_between":
                            # Area chart
                            ax1.fill_between(x_values, y_vals, alpha=0.6, label=label, color=color)
                            ax1.plot(x_values, y_vals, color=color, linewidth=2)
                            
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
                            # Heatmap (simplified)
                            heatmap_data = series.get("values", [])
                            if len(heatmap_data) > 0:
                                ax1.imshow(heatmap_data, cmap='viridis', aspect='auto')
                                
                        elif mpl_chart_type == "contour":
                            # Contour plot (simplified)
                            if len(y_vals) > 0:
                                # Create a simple 2D array for contour
                                contour_data = [y_vals] if len(y_vals) > 0 else [[0]]
                                ax1.contour(contour_data)
                                
                        else:
                            # Fallback to scatter for unknown types
                            current_app.logger.warning(f"⚠️ Unknown matplotlib chart type '{series_type}', falling back to scatter")
                            ax1.scatter(x_values, y_vals, label=label, color=color, alpha=0.7)

                    # Add data labels to Matplotlib chart if enabled
                    if show_data_labels and (data_label_format or value_format or data_label_font_size or data_label_color):
                        current_app.logger.debug(f"Adding data labels to Matplotlib chart")
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
                                            
                                            # Add text label on top of bar
                                            ax1.text(j, val, formatted_val, 
                                                    ha='center', va='bottom', 
                                                    fontsize=data_label_font_size or 10,
                                                    color=data_label_color or '#000000',
                                                    fontweight='bold')
                                
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
                                            
                                            # Add text label above line point
                                            ax2.text(j, val, formatted_val, 
                                                    ha='center', va='bottom', 
                                                    fontsize=data_label_font_size or 10,
                                                    color=data_label_color or '#000000',
                                                    fontweight='bold')

                    # Set labels and styling
                    if chart_type != "pie":
                        ax1.set_xlabel(chart_meta.get("x_label", chart_config.get("x_axis_title", "X")), fontsize=font_size or 11)
                        ax1.set_ylabel(chart_meta.get("primary_y_label", chart_config.get("primary_y_label", "Primary Y")), fontsize=font_size or 11)
                        if "secondary_y_label" in chart_meta or "secondary_y_label" in chart_config:
                            ax2.set_ylabel(chart_meta.get("secondary_y_label", chart_config.get("secondary_y_label", "Secondary Y")), fontsize=font_size or 11)

                        # Set axis tick font size if provided
                        if axis_tick_font_size:
                            ax1.tick_params(axis='x', labelsize=axis_tick_font_size, rotation=45)
                            ax1.tick_params(axis='y', labelsize=axis_tick_font_size)
                            ax2.tick_params(axis='y', labelsize=axis_tick_font_size)
                        else:
                            ax1.tick_params(axis='x', rotation=45)
                        
                        # Apply X-axis label distance using tick parameters
                        if x_axis_label_distance:
                            # Use tick label padding to control distance
                            ax1.tick_params(axis='x', pad=x_axis_label_distance)
                            current_app.logger.debug(f"Applied X-axis tick padding: {x_axis_label_distance}")
                        
                        # Apply secondary y-axis formatting for Matplotlib
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
                        
                        # Gridlines
                        current_app.logger.debug(f"Gridlines setting: {show_gridlines}")
                        if show_gridlines:
                            ax1.grid(True, linestyle=gridline_style or '--', color=gridline_color or '#ccc', alpha=0.6)
                            ax2.grid(True, linestyle=gridline_style or '--', color=gridline_color or '#ccc', alpha=0.6)
                        else:
                            ax1.grid(False)
                            ax2.grid(False)
                        
                        # Apply primary y-axis formatting for Matplotlib
                        if axis_tick_format:
                            from matplotlib.ticker import FuncFormatter
                            if "$" in axis_tick_format:
                                def currency_formatter(x, pos):
                                    return f'${x:,.0f}'
                                ax1.yaxis.set_major_formatter(FuncFormatter(currency_formatter))
                        if y_axis_min_max:
                            current_app.logger.debug(f"Setting Matplotlib Y-axis range to: {y_axis_min_max}")
                            # Ensure the range is properly applied
                            if isinstance(y_axis_min_max, list) and len(y_axis_min_max) == 2:
                                ax1.set_ylim(y_axis_min_max[0], y_axis_min_max[1])
                                current_app.logger.debug(f"Applied Y-axis range: {y_axis_min_max[0]} to {y_axis_min_max[1]}")
                            else:
                                current_app.logger.warning(f"Invalid Y-axis range format: {y_axis_min_max}")
                        
                        # Legend
                        show_legend = chart_meta.get("showlegend", chart_meta.get("legend", True))
                        if show_legend:
                            ax1.legend(loc='best')
                        
                        ax1.set_title(title, fontsize=font_size or 14, weight='bold')
                
                # Apply axis label distances for Matplotlib
                current_app.logger.debug(f"X-axis label distance: {x_axis_label_distance}")
                current_app.logger.debug(f"Y-axis label distance: {y_axis_label_distance}")
                
                if x_axis_label_distance or y_axis_label_distance:
                    # Get current subplot parameters
                    current_bottom = fig_mpl.subplotpars.bottom
                    current_left = fig_mpl.subplotpars.left
                    
                    if x_axis_label_distance:
                        # Convert the distance to a fraction of the figure height
                        # Higher x_axis_label_distance values will push labels further down
                        adjustment = x_axis_label_distance / 500.0  # Increased conversion factor for more visible effect
                        fig_mpl.subplots_adjust(bottom=current_bottom - adjustment)
                        current_app.logger.debug(f"Applied X-axis adjustment: {adjustment}")
                    
                    if y_axis_label_distance:
                        # Convert the distance to a fraction of the figure width
                        adjustment = y_axis_label_distance / 500.0  # Increased conversion factor
                        fig_mpl.subplots_adjust(left=current_left - adjustment)
                        current_app.logger.debug(f"Applied Y-axis adjustment: {adjustment}")
                
                # Apply tight_layout but preserve manual adjustments
                fig_mpl.tight_layout()
                
                # Re-apply manual adjustments after tight_layout with larger effect
                if x_axis_label_distance or y_axis_label_distance:
                    if x_axis_label_distance:
                        adjustment = x_axis_label_distance / 300.0  # Even larger effect after tight_layout
                        fig_mpl.subplots_adjust(bottom=fig_mpl.subplotpars.bottom - adjustment)
                        current_app.logger.debug(f"Re-applied X-axis adjustment: {adjustment}")
                    
                    if y_axis_label_distance:
                        adjustment = y_axis_label_distance / 300.0  # Even larger effect after tight_layout
                        fig_mpl.subplots_adjust(left=fig_mpl.subplotpars.left - adjustment)
                        current_app.logger.debug(f"Re-applied Y-axis adjustment: {adjustment}")

                tmpfile = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.savefig(tmpfile.name, bbox_inches='tight')
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
        for para_idx, para in enumerate(doc.paragraphs):
            replace_text_in_paragraph(para)
        
        # Paragraph processing completed silently
        
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

        replace_text_in_tables()

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

# Helper to get absolute path from DB (relative) path

def get_abs_path_from_db_path(db_path):
    # If already absolute, return as is (for backward compatibility)
    if os.path.isabs(db_path):
        return db_path
    # Use the directory containing this file as the root
    return os.path.join(os.path.abspath(os.path.dirname(__file__)), db_path)

@projects_bp.route('/api/projects', methods=['GET'])
@login_required
def get_projects():
    # Access MongoDB via current_app.mongo.db
    projects = list(current_app.mongo.db.projects.find({'user_id': current_user.get_id()}))
    for project in projects:
        project['id'] = str(project['_id'])
        del project['_id'] 
    return jsonify({'projects': projects})

@projects_bp.route('/api/projects', methods=['POST'])
@login_required
def create_project():
    name = request.form.get('name')
    description = request.form.get('description')
    file = request.files.get('file') 

    if not name or not description:
        return jsonify({'error': 'Missing required fields (name or description)'}), 400

    file_path = None
    if file:
        if not allowed_file(file.filename): 
            return jsonify({'error': 'File type not allowed'}), 400
        os.makedirs(UPLOAD_FOLDER, exist_ok=True) 
        filename = secure_filename(file.filename)
        # Store relative path in DB
        file_path = os.path.join('uploads', filename)
        abs_file_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), file_path)
        file.save(abs_file_path)

    project = {
        'name': name,
        'description': description,
        'user_id': current_user.get_id(),
        'file_path': file_path,
        'created_at': datetime.utcnow().isoformat() 
    }
    # Access MongoDB via current_app.mongo.db
    project_id = current_app.mongo.db.projects.insert_one(project).inserted_id
    project['id'] = str(project_id)
    del project['_id']

    return jsonify({'message': 'Project created successfully', 'project': project}), 201

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

    template_file_path = project.get('file_path')
    abs_template_file_path = get_abs_path_from_db_path(template_file_path)
    current_app.logger.info(f"📄 Template path: {template_file_path}")
    current_app.logger.info(f"📄 Absolute template path: {abs_template_file_path}")
    
    if not template_file_path or not os.path.exists(abs_template_file_path):
        current_app.logger.error(f"❌ Template file not found: {abs_template_file_path}")
        return jsonify({'error': 'Word template file not found for this project. Please upload it during project creation.'}), 400
    
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
    generated_report_path = _generate_report(project_id, abs_template_file_path, temp_report_data_path)
    
    # Clean up the temporary uploaded report data file and directory
    import shutil
    shutil.rmtree(temp_dir)
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
        # You may want to extract report name/code from the Excel file or filename
        report_name = os.path.splitext(os.path.basename(excel_path))[0]
        report_code = f"{project_id}_{idx}"

        current_app.logger.info(f"Processing file {idx}/{total_files}: {report_name}")

        # Generate report
        template_file_path = current_app.mongo.db.projects.find_one({'_id': ObjectId(project_id)})['file_path']
        abs_template_file_path = get_abs_path_from_db_path(template_file_path)
        output_path = _generate_report(f"{project_id}_{idx}", abs_template_file_path, excel_path)
        if output_path:
            # Save in both folders
            shutil.copy(output_path, os.path.join(output_folder_name, f"{report_name}.docx"))
            shutil.copy(output_path, os.path.join(output_folder_code, f"{report_code}.docx"))
            generated_files.append({'name': report_name, 'code': report_code})
            current_app.logger.info(f"✅ Successfully generated report {idx}/{total_files}: {report_name}")
        else:
            current_app.logger.error(f"❌ Failed to generate report {idx}/{total_files}: {report_name}")
        
        # Log progress
        current_app.logger.info(f"Progress: {idx}/{total_files} reports processed")

    # Create zip file in temporary location
    zip_output_path = os.path.join(temp_dir, f'batch_reports_{project_id}.zip')
    with zipfile.ZipFile(zip_output_path, 'w') as zipf:
        for file_info in generated_files:
            file_path = os.path.join(output_folder_name, f"{file_info['name']}.docx")
            zipf.write(file_path, arcname=f"{file_info['name']}.docx")

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
