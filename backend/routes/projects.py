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
                if "source_sheet" in chart_meta:
                    wb = openpyxl.load_workbook(data_file_path, data_only=True)
                    sheet = wb[chart_meta["source_sheet"]]
                    # Pie chart extraction
                    if chart_meta.get("chart_type", "").lower() == "pie":
                        labels = extract_excel_range(sheet, chart_meta["category_range"]) if "category_range" in chart_meta else None
                        values = extract_excel_range(sheet, chart_meta["value_range"]) if "value_range" in chart_meta else None
                        extracted_series_data = [{
                            "name": chart_meta.get("chart_title", "Pie"),
                            "type": "pie",
                            "labels": labels,
                            "values": values,
                            "marker": {"color": series_meta.get("colors")}
                        }]
                    # Stacked column extraction
                    elif chart_meta.get("chart_type", "").lower() == "stacked_column":
                        x_axis = extract_excel_range(sheet, chart_meta["category_range"]) if "category_range" in chart_meta else None
                        series_labels = series_meta.get("labels", [])
                        series_colors = series_meta.get("colors", [])
                        value_ranges = chart_meta.get("value_range", [])
                        extracted_series_data = []
                        for idx, rng in enumerate(value_ranges):
                            values = extract_excel_range(sheet, rng)
                            extracted_series_data.append({
                                "name": series_labels[idx] if idx < len(series_labels) else f"Series {idx+1}",
                                "type": "bar",
                                "values": values,
                                "marker": {"color": series_colors[idx] if idx < len(series_colors) else None}
                            })
                        extracted_x_axis = x_axis

                # Use extracted data if present
                if extracted_series_data is not None:
                    series_data = extracted_series_data
                if extracted_x_axis is not None:
                    x_values = extracted_x_axis

                # --- Plotly interactive chart generation ---
                fig = go.Figure()
                x_values = series_meta.get("x_axis", [])
                series_data = series_meta.get("data", [])
                colors = series_meta.get("colors", [])

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
                    if color:
                        if isinstance(color, list) and series_type == "bar":
                            trace_kwargs["marker"] = dict(color=color)
                        elif isinstance(color, str):
                            trace_kwargs["marker_color"] = color
                    if bar_width and series_type == "bar":
                        trace_kwargs["width"] = bar_width
                    if orientation and series_type == "bar":
                        trace_kwargs["orientation"] = orientation[0].lower() if isinstance(orientation, str) else orientation
                    if bar_border_color and series_type == "bar":
                        trace_kwargs["marker_line_color"] = bar_border_color
                    if bar_border_width and series_type == "bar":
                        trace_kwargs["marker_line_width"] = bar_border_width

                    if series_type == "bar":
                        fig.add_trace(go.Bar(**trace_kwargs,
                            hovertemplate=f"<b>{label}</b><br>Year: %{{x}}<br>Value: %{{y}}<extra></extra>"
                        ))
                    elif series_type == "line":
                        fig.add_trace(go.Scatter(**trace_kwargs,
                            mode='lines+markers',
                            hovertemplate=f"<b>{label}</b><br>Year: %{{x}}<br>Value: %{{y}}<extra></extra>"
                        ))
                    elif series_type == "area":
                        # Area chart: use Scatter with fill='tozeroy'
                        fig.add_trace(go.Scatter(**trace_kwargs,
                            mode='lines',
                            fill='tozeroy',
                            hovertemplate=f"<b>{label}</b><br>Year: %{{x}}<br>Value: %{{y}}<extra></extra>"
                        ))
                    elif series_type == "pie":
                        # Pie chart: use go.Pie, x_vals as labels, y_vals as values
                        pie_kwargs = {
                            "labels": x_vals,
                            "values": y_vals,
                            "name": label
                        }
                        # Pie colors
                        if color:
                            pie_kwargs["marker"] = dict(colors=color) if isinstance(color, list) else dict(colors=[color])
                        fig.add_trace(go.Pie(**pie_kwargs,
                            hovertemplate=f"<b>{label}</b><br>%{{label}}: %{{value}}<extra></extra>"
                        ))

                # --- Layout updates ---
                layout_updates = {}
                # Title and axis labels
                layout_updates["title"] = title
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
                # Legend
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
                # Axis min/max and tick format
                if y_axis_min_max:
                    layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                    layout_updates["yaxis"]["range"] = y_axis_min_max
                if axis_tick_format:
                    layout_updates["yaxis"] = layout_updates.get("yaxis", {})
                    layout_updates["yaxis"]["tickformat"] = axis_tick_format
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
                if data_label_format or data_label_font_size or data_label_color:
                    for trace in fig.data:
                        if data_label_format:
                            if trace.type == 'bar':
                                trace.update(texttemplate=f"%{{y:{data_label_format}}}", textposition="auto")
                            elif trace.type == 'scatter':
                                trace.update(texttemplate=f"%{{y:{data_label_format}}}", textposition="top center")
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

                # --- Matplotlib static chart for DOCX (basic mapping) ---
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

                    if series_type == "bar":
                        if isinstance(color, list):
                            for j, val in enumerate(y_vals):
                                bar_color = color[j % len(color)]
                                ax1.bar(x_values[j], val, color=bar_color, alpha=0.7)
                        else:
                            ax1.bar(x_values, y_vals, label=label, color=color, alpha=0.7)
                    elif series_type == "line":
                        ax2.plot(x_values, y_vals, label=label, color=color, marker='o', linewidth=2)

                ax1.set_xlabel(chart_meta.get("x_label", chart_config.get("x_axis_title", "X")), fontsize=font_size or 11)
                ax1.set_ylabel(chart_meta.get("primary_y_label", chart_config.get("y_axis_title", "Primary Y")), fontsize=font_size or 11)
                if "secondary_y_label" in chart_meta or "secondary_y_label" in chart_config:
                    ax2.set_ylabel(chart_meta.get("secondary_y_label", chart_config.get("secondary_y_label", "Secondary Y")), fontsize=font_size or 11)

                ax1.tick_params(axis='x', rotation=45)
                if show_gridlines is not None:
                    ax1.grid(bool(show_gridlines), linestyle=gridline_style or '--', color=gridline_color or '#ccc', alpha=0.6)
                    ax2.grid(bool(show_gridlines), linestyle=gridline_style or '--', color=gridline_color or '#ccc', alpha=0.6)
                else:
                    ax1.grid(False)
                    ax2.grid(False)
                fig_mpl.suptitle(title, fontsize=font_size or 14, weight='bold')
                fig_mpl.tight_layout()

                tmpfile = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.savefig(tmpfile.name, bbox_inches='tight')
                plt.close(fig_mpl)

                return tmpfile.name

            except Exception as e:
                current_app.logger.error(f"❌ Chart generation failed for {chart_tag}: {e}")
                return None

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
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)

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
    if not template_file_path or not os.path.exists(template_file_path):
        return jsonify({'error': 'Word template file not found for this project. Please upload it during project creation.'}), 400
    
    # Save the uploaded report data file temporarily
    report_data_filename = secure_filename(report_file.filename)
    temp_upload_folder = os.path.join(current_app.root_path, UPLOAD_FOLDER, 'temp_reports')
    os.makedirs(temp_upload_folder, exist_ok=True)
    temp_report_data_path = os.path.join(temp_upload_folder, report_data_filename)
    report_file.save(temp_report_data_path)

    # Generate the report
    generated_report_path = _generate_report(project_id, template_file_path, temp_report_data_path)
    
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

    if not generated_report_path or not os.path.exists(generated_report_path):
        return jsonify({'error': 'Generated report not found for this project'}), 404

    return send_file(generated_report_path, as_attachment=True)

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
        output_path = _generate_report(f"{project_id}_{idx}", template_file_path, excel_path)
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
