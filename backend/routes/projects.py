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
# Ensure the upload and reports directories exist
os.makedirs(os.path.join(UPLOAD_FOLDER, 'reports'), exist_ok=True)
os.makedirs(os.path.join(UPLOAD_FOLDER, 'temp_reports'), exist_ok=True)
os.makedirs(os.path.join(UPLOAD_FOLDER, 'charts'), exist_ok=True)

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

        df = pd.read_excel(data_file_path, sheet_name='sample')
        df.columns = df.columns.str.strip().str.replace(" ", "_").str.replace("__", "_")

        text_map = {str(k).strip().lower(): str(v).strip() for k, v in zip(df["Text_Tag"], df["Text"]) if pd.notna(k) and pd.notna(v)}
        chart_attr_map = {str(k).strip().lower(): str(v).strip() for k, v in zip(df["Chart_Tag"], df["Chart_Attributes"]) if pd.notna(k) and pd.notna(v)}
        chart_type_map = {str(k).strip().lower(): str(v).strip() for k, v in zip(df["Chart_Tag"], df["Chart_Type"]) if pd.notna(k) and pd.notna(v)}

        flat_data_map = {}
        for _, row in df.iterrows():
            chart_tag = row.get("Chart_Tag")
            if not isinstance(chart_tag, str) or not chart_tag:
                continue
            section_prefix = chart_tag.replace('_chart', '').lower()

            for col in df.columns:
                if col.startswith("Chart_Data_Y"):
                    year = col.replace("Chart_Data_", "")
                    flat_data_map[f"{section_prefix}_{year.lower()}"] = str(row[col])
                elif col.startswith("Growth_Y"):
                    year = col.replace("Growth_", "")
                    flat_data_map[f"{section_prefix}_{year.lower()}_kpi2"] = str(row[col])

            if pd.notna(row.get("Chart_Data_CAGR")):
                flat_data_map[f"{section_prefix}_cgrp"] = str(row["Chart_Data_CAGR"])

        flat_data_map = {k.lower(): v for k, v in flat_data_map.items()}

        doc = Document(template_path)

        def replace_text_in_paragraph(paragraph):
            for run in paragraph.runs:
                original_text = run.text
                matches = re.findall(r"\$\{(.*?)\}", run.text)
                for match in matches:
                    key_lower = match.lower()
                    val = flat_data_map.get(key_lower) or text_map.get(key_lower)
                    if val:
                        pattern = re.compile(re.escape(f"${{{match}}}"), re.IGNORECASE)
                        run.text = pattern.sub(val, run.text)
                        current_app.logger.debug(f"✅ Replaced ${{{match}}} with '{val}' in run: {original_text}")

        def replace_text_in_tables():
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            replace_text_in_paragraph(para)

        def generate_chart(data_dict, chart_tag):
            import plotly.graph_objects as go
            import matplotlib.pyplot as plt
            from openpyxl.utils import column_index_from_string
            import numpy as np
            import os
            import tempfile
            import json
            import re

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
                gridline_color = chart_config.get("gridline_color")
                gridline_style = chart_config.get("gridline_style")
                chart_background = chart_config.get("chart_background")
                plot_background = chart_config.get("plot_background")
                data_label_format = chart_config.get("data_label_format")
                data_label_font_size = chart_config.get("data_label_font_size")
                data_label_color = chart_config.get("data_label_color")
                axis_tick_format = chart_config.get("axis_tick_format")
                y_axis_min_max = chart_config.get("y_axis_min_max")
                sort_order = chart_config.get("sort_order")
                data_grouping = chart_config.get("data_grouping")
                annotations = chart_config.get("annotations", [])
                axis_tick_font_size = chart_config.get("axis_tick_font_size")

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

                # If chart_meta specifies source_sheet and ranges, extract data from Excel
                extracted_x_axis = None
                extracted_series_data = None
                chart_type_from_meta = chart_meta.get("chart_type", "").lower()
                
                if "source_sheet" in chart_meta:
                    wb = openpyxl.load_workbook(data_file_path, data_only=True)
                    sheet = wb[chart_meta["source_sheet"]]
                    
                    # Generic chart type extraction
                    chart_extraction_mapping = {
                        # Single series charts
                        "pie": "single_series",
                        "donut": "single_series", 
                        "histogram": "single_series",
                        "box": "single_series",
                        "violin": "single_series",
                        "waterfall": "single_series",
                        "funnel": "single_series",
                        "sunburst": "single_series",
                        "treemap": "single_series",
                        "icicle": "single_series",
                        "sankey": "single_series",
                        "indicator": "single_series",
                        
                        # Multi-series charts
                        "bar": "multi_series",
                        "column": "multi_series",
                        "stacked_column": "multi_series", 
                        "horizontal_bar": "multi_series",
                        "line": "multi_series",
                        "scatter": "multi_series",
                        "scatter_line": "multi_series",
                        "area": "multi_series",
                        "filled_area": "multi_series",
                        "bubble": "multi_series",
                        "heatmap": "multi_series",
                        "contour": "multi_series",
                        "scatter3d": "multi_series",
                        "surface": "multi_series",
                        "mesh3d": "multi_series",
                        "candlestick": "multi_series",
                        "ohlc": "multi_series",
                        "scattergeo": "multi_series",
                        "choropleth": "multi_series"
                    }
                    
                    extraction_type = chart_extraction_mapping.get(chart_type_from_meta, "multi_series")
                    
                    if extraction_type == "single_series":
                        # Single series chart extraction (pie, histogram, etc.)
                        labels = extract_excel_range(sheet, chart_meta["category_range"]) if "category_range" in chart_meta else None
                        values = extract_excel_range(sheet, chart_meta["value_range"]) if "value_range" in chart_meta else None
                        extracted_series_data = [{
                            "name": chart_meta.get("chart_title", chart_type_from_meta.title()),
                            "type": chart_type_from_meta,
                            "labels": labels,
                            "values": values,
                            "marker": {"color": series_meta.get("colors")}
                        }]
                        
                    elif extraction_type == "multi_series":
                        # Multi-series chart extraction (bar, line, area, etc.)
                        x_axis = extract_excel_range(sheet, chart_meta["category_range"]) if "category_range" in chart_meta else None
                        series_labels = series_meta.get("labels", [])
                        series_colors = series_meta.get("colors", [])
                        value_ranges = chart_meta.get("value_range", [])
                        extracted_series_data = []
                        
                        for idx, rng in enumerate(value_ranges):
                            values = extract_excel_range(sheet, rng)
                            extracted_series_data.append({
                                "name": series_labels[idx] if idx < len(series_labels) else f"Series {idx+1}",
                                "type": chart_type_from_meta,
                                "values": values,
                                "marker": {"color": series_colors[idx] if idx < len(series_colors) else None}
                            })
                        extracted_x_axis = x_axis

                # Use extracted data if present, otherwise fall back to series_meta
                if extracted_series_data is not None:
                    series_data = extracted_series_data
                    chart_type = chart_type_from_meta  # Use chart type from meta
                else:
                    series_data = series_meta.get("data", [])
                    chart_type = chart_type_map.get(chart_tag_lower, "").lower().strip()
                
                if extracted_x_axis is not None:
                    x_values = extracted_x_axis
                else:
                    x_values = series_meta.get("x_axis", [])
                
                colors = series_meta.get("colors", [])

                # --- Plotly interactive chart generation ---
                fig = go.Figure()

                def extract_values_from_range(cell_range):
                    start_cell, end_cell = cell_range.split(":")
                    start_col, start_row = re.match(r"([A-Z]+)(\d+)", start_cell).groups()
                    end_col, end_row = re.match(r"([A-Z]+)(\d+)", end_cell).groups()

                    start_col_idx = column_index_from_string(start_col) - 1
                    end_col_idx = column_index_from_string(end_col) - 1
                    start_row_idx = int(start_row) - 2
                    end_row_idx = int(end_row) - 2

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
                                if chart_type == "stacked_column":
                                    trace_kwargs["barmode"] = "stack"
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
                                if "size" not in trace_kwargs and len(y_vals) > 0:
                                    # Create size based on values if not provided
                                    max_val = max(y_vals) if y_vals else 1
                                    trace_kwargs["size"] = [abs(v/max_val) * 20 for v in y_vals]
                                
                                trace_kwargs["mode"] = "markers"
                                
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
                    layout_updates["yaxis_title"] = chart_meta.get("primary_y_label", chart_config.get("y_axis_title", "Y"))
                
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
                show_legend = chart_meta.get("legend", True)
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
                        layout_updates["yaxis"]["range"] = y_axis_min_max
                    if axis_tick_format:
                        layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                        layout_updates["yaxis"]["tickformat"] = axis_tick_format
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
                
                if show_data_labels and (data_label_format or value_format or data_label_font_size or data_label_color):
                    for trace in fig.data:
                        if trace.type in ['bar', 'scatter']:
                            # Use value_format from chart_meta if available, otherwise use data_label_format
                            format_to_use = value_format if value_format else data_label_format
                            if format_to_use:
                                if trace.type == 'bar':
                                    trace.update(texttemplate=f"%{{y:{format_to_use}}}", textposition="auto")
                                elif trace.type == 'scatter':
                                    trace.update(texttemplate=f"%{{y:{format_to_use}}}", textposition="top center")
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

                fig.update_layout(**layout_updates)

                interactive_path = os.path.join(UPLOAD_FOLDER, 'reports', f'interactive_{chart_tag_lower}.html')
                fig.write_html(interactive_path)
                current_app.logger.debug(f"\U0001F310 Interactive chart saved to: {interactive_path}")

                # --- Matplotlib static chart for DOCX ---
                if chart_type == "pie":
                    # Check if this is an expanded pie chart
                    expanded_segment = chart_meta.get("expanded_segment")
                    
                    if expanded_segment and len(series_data) == 1:
                        # Create subplot for expanded pie chart
                        fig_mpl, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 8), dpi=200)
                        
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
                        
                else:
                    # Bar, line, area charts
                    fig_mpl, ax1 = plt.subplots(figsize=(10, 6), dpi=200)
                    ax2 = ax1.twinx()

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
                            if len(y_vals) > 0:
                                # Create a simple 2D array for heatmap
                                heatmap_data = [y_vals] if len(y_vals) > 0 else [[0]]
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

                    # Set labels and styling
                    if chart_type != "pie":
                        ax1.set_xlabel(chart_meta.get("x_label", chart_config.get("x_axis_title", "X")), fontsize=font_size or 11)
                        ax1.set_ylabel(chart_meta.get("primary_y_label", chart_config.get("y_axis_title", "Primary Y")), fontsize=font_size or 11)
                        if "secondary_y_label" in chart_meta or "secondary_y_label" in chart_config:
                            ax2.set_ylabel(chart_meta.get("secondary_y_label", chart_config.get("secondary_y_label", "Secondary Y")), fontsize=font_size or 11)

                        # Set axis tick font size if provided
                        if axis_tick_font_size:
                            ax1.tick_params(axis='x', labelsize=axis_tick_font_size, rotation=45)
                            ax1.tick_params(axis='y', labelsize=axis_tick_font_size)
                            ax2.tick_params(axis='y', labelsize=axis_tick_font_size)
                        else:
                            ax1.tick_params(axis='x', rotation=45)
                        
                        # Gridlines
                        if show_gridlines:
                            ax1.grid(True, linestyle=gridline_style or '--', color=gridline_color or '#ccc', alpha=0.6)
                            ax2.grid(True, linestyle=gridline_style or '--', color=gridline_color or '#ccc', alpha=0.6)
                        else:
                            ax1.grid(False)
                            ax2.grid(False)
                        
                        # Legend
                        if chart_meta.get("legend", True):
                            ax1.legend(loc='best')
                        
                        ax1.set_title(title, fontsize=font_size or 14, weight='bold')
                
                fig_mpl.tight_layout()

                tmpfile = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.savefig(tmpfile.name, bbox_inches='tight')
                plt.close(fig_mpl)

                return tmpfile.name

            except Exception as e:
                current_app.logger.error(f"❌ Chart generation failed for {chart_tag}: {e}")
                return None

        # Insert charts into paragraphs
        for para in doc.paragraphs:
            replace_text_in_paragraph(para)
            full_text = ''.join(run.text for run in para.runs)
            chart_placeholders = re.findall(r"\$\{(section\d+_chart)\}", full_text, flags=re.IGNORECASE)
            for tag in chart_placeholders:
                if tag.lower() in chart_attr_map:
                    try:
                        chart_img = generate_chart({}, tag)
                        if chart_img:
                            para.text = re.sub(rf"\$\{{{tag}\}}", "", para.text, flags=re.IGNORECASE)
                            para.add_run().add_picture(chart_img, width=Inches(5.5))
                    except Exception as e:
                        current_app.logger.error(f"⚠️ Failed to insert chart for tag {tag}: {e}")

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
                                except Exception as e:
                                    current_app.logger.error(f"⚠️ Failed to insert chart in table for tag {tag}: {e}")

        output_path = os.path.join(UPLOAD_FOLDER, 'reports', f'output_report_{project_id}.docx')
        doc.save(output_path)
        current_app.logger.debug(f"✅ Report saved to: {output_path}")
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
    if 'report_file' not in request.files:
        return jsonify({'error': 'No report file provided'}), 400

    report_file = request.files['report_file']

    if report_file.filename == '':
        return jsonify({'error': 'No selected report file'}), 400

    if not allowed_report_file(report_file.filename):
        return jsonify({'error': 'Report file type not allowed. Only .xlsx or .csv are accepted.'}), 400

    try:
        project_id_obj = ObjectId(project_id)
    except:
        return jsonify({'error': 'Invalid project ID'}), 400

    project = current_app.mongo.db.projects.find_one({'_id': project_id_obj, 'user_id': current_user.get_id()})
    if not project:
        return jsonify({'error': 'Project not found or unauthorized'}), 404

    template_file_path = project.get('file_path')
    abs_template_file_path = get_abs_path_from_db_path(template_file_path)
    if not template_file_path or not os.path.exists(abs_template_file_path):
        return jsonify({'error': 'Word template file not found for this project. Please upload it during project creation.'}), 400
    
    # Save the uploaded report data file temporarily
    report_data_filename = secure_filename(report_file.filename)
    temp_upload_folder = os.path.join(current_app.root_path, UPLOAD_FOLDER, 'temp_reports')
    os.makedirs(temp_upload_folder, exist_ok=True)
    temp_report_data_path = os.path.join(temp_upload_folder, report_data_filename)
    report_file.save(temp_report_data_path)

    # Generate the report
    generated_report_path = _generate_report(project_id, abs_template_file_path, temp_report_data_path)
    
    # Clean up the temporary uploaded report data file
    os.remove(temp_report_data_path)

    if generated_report_path:
        # Update project with generated report path
        current_app.mongo.db.projects.update_one(
            {'_id': project_id_obj},
            {'$set': {'generated_report_path': generated_report_path, 'report_generated_at': datetime.utcnow().isoformat()}}
        )
        return jsonify({'message': 'Report generated successfully', 'report_path': generated_report_path}), 200
    else:
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
    abs_generated_report_path = get_abs_path_from_db_path(generated_report_path)
    if not generated_report_path or not os.path.exists(abs_generated_report_path):
        return jsonify({'error': 'Generated report not found for this project'}), 404

    return send_file(abs_generated_report_path, as_attachment=True)

@projects_bp.route('/api/reports/<chart_filename>/download_html', methods=['GET'])
@login_required
def download_chart_html(chart_filename):
    chart_path = os.path.join(UPLOAD_FOLDER, 'reports', chart_filename)
    if not os.path.exists(chart_path):
        return jsonify({'error': 'Chart HTML file not found'}), 404
    return send_file(chart_path, as_attachment=True)

@projects_bp.route('/api/projects/<project_id>/upload_zip', methods=['POST'])
@login_required
def upload_zip_and_generate_reports(project_id):
    if 'zip_file' not in request.files:
        return jsonify({'error': 'No zip file provided'}), 400

    zip_file = request.files['zip_file']
    if not zip_file.filename.endswith('.zip'):
        return jsonify({'error': 'Only .zip files are allowed'}), 400

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

    # Prepare output folders
    output_folder_name = os.path.join(UPLOAD_FOLDER, 'reports_by_name')
    output_folder_code = os.path.join(UPLOAD_FOLDER, 'reports_by_code')
    os.makedirs(output_folder_name, exist_ok=True)
    os.makedirs(output_folder_code, exist_ok=True)

    generated_files = []
    for idx, excel_path in enumerate(excel_files, 1):
        # You may want to extract report name/code from the Excel file or filename
        report_name = os.path.splitext(os.path.basename(excel_path))[0]
        report_code = f"{project_id}_{idx}"

        # Generate report
        template_file_path = current_app.mongo.db.projects.find_one({'_id': ObjectId(project_id)})['file_path']
        abs_template_file_path = get_abs_path_from_db_path(template_file_path)
        output_path = _generate_report(f"{project_id}_{idx}", abs_template_file_path, excel_path)
        if output_path:
            # Save in both folders
            shutil.copy(output_path, os.path.join(output_folder_name, f"{report_name}.docx"))
            shutil.copy(output_path, os.path.join(output_folder_code, f"{report_code}.docx"))
            generated_files.append({'name': report_name, 'code': report_code})
        # Log progress
        current_app.logger.info(f"Generated {idx} of {len(excel_files)} reports")

    # Optionally, zip all generated reports for download
    zip_output_path = os.path.join(UPLOAD_FOLDER, 'reports', f'batch_reports_{project_id}.zip')
    with zipfile.ZipFile(zip_output_path, 'w') as zipf:
        for file_info in generated_files:
            file_path = os.path.join(output_folder_name, f"{file_info['name']}.docx")
            zipf.write(file_path, arcname=f"{file_info['name']}.docx")

    # Clean up temp
    shutil.rmtree(temp_dir)

    return jsonify({
        'message': f'Generated {len(generated_files)} reports.',
        'download_zip': zip_output_path,
        'reports': generated_files
    })
