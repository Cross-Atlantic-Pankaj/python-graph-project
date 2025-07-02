#!/usr/bin/env python3
"""
Test script to demonstrate chart generation with the provided configurations
"""

import json
import tempfile
import os
from routes.projects import create_expanded_pie_chart
import plotly.graph_objects as go
from plotly.subplots import make_subplots

def test_stacked_column_chart():
    """Test stacked column chart configuration"""
    config = {
        "chart_meta": {
            "chart_type": "stacked_column",
            "legend": True,
            "data_labels": True,
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
    
    # Mock data for testing
    categories = ["2020", "2021", "2022", "2023"]
    series_data = [
        [25.5, 28.2, 30.1, 32.8],  # Commission-Based
        [15.3, 18.7, 22.4, 25.9],  # Service/Refurbishment
        [8.2, 12.1, 15.6, 18.3]    # Ad-Based
    ]
    
    fig = go.Figure()
    
    for i, (label, values, color) in enumerate(zip(config["series"]["labels"], series_data, config["series"]["colors"])):
        fig.add_trace(go.Bar(
            x=categories,
            y=values,
            name=label,
            marker_color=color,
            text=[f"{v:.1f}%" for v in values],
            textposition="auto"
        ))
    
    fig.update_layout(
        title="Revenue by Year - Stacked Column",
        barmode="stack",
        xaxis_title="Year",
        yaxis_title="Revenue (%)",
        showlegend=True,
        height=500
    )
    
    # Save the chart
    fig.write_html("test_stacked_column.html")
    print("✅ Stacked column chart saved as test_stacked_column.html")

def test_pie_chart():
    """Test pie chart configuration"""
    config = {
        "chart_meta": {
            "chart_type": "pie",
            "legend": True,
            "data_labels": True,
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
    
    # Mock data for testing
    labels = config["series"]["labels"]
    values = [25.5, 18.3, 15.7, 12.4, 10.2, 8.9, 9.0]
    colors = config["series"]["colors"]
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        textinfo="label+percent+value",
        textposition="outside",
        marker=dict(colors=colors)
    )])
    
    fig.update_layout(
        title="Market Distribution - Pie Chart",
        showlegend=True,
        height=600
    )
    
    # Save the chart
    fig.write_html("test_pie_chart.html")
    print("✅ Pie chart saved as test_pie_chart.html")

def test_expanded_pie_chart():
    """Test expanded pie chart configuration"""
    config = {
        "chart_meta": {
            "chart_type": "pie",
            "legend": True,
            "data_labels": True,
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
    
    # Mock data for testing
    labels = config["series"]["labels"]
    values = [22.5, 28.3, 15.7, 12.4, 8.2, 7.9, 5.0]
    colors = config["series"]["colors"]
    expanded_segment = config["chart_meta"]["expanded_segment"]
    
    fig = create_expanded_pie_chart(
        labels=labels,
        values=values,
        colors=colors,
        expanded_segment=expanded_segment,
        title="Product Category Distribution - Expanded Pie Chart",
        value_format="%"
    )
    
    # Save the chart
    fig.write_html("test_expanded_pie_chart.html")
    print("✅ Expanded pie chart saved as test_expanded_pie_chart.html")

def test_area_chart():
    """Test area chart configuration"""
    config = {
        "chart_meta": {
            "chart_type": "area",
            "legend": True,
            "data_labels": True,
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
    
    # Mock data for testing
    categories = ["Q1", "Q2", "Q3", "Q4"]
    series_data = [
        [100, 120, 140, 160],  # Revenue
        [80, 95, 110, 125],    # Costs
        [20, 25, 30, 35]       # Profit
    ]
    
    fig = go.Figure()
    
    for i, (label, values, color) in enumerate(zip(config["series"]["labels"], series_data, config["series"]["colors"])):
        fig.add_trace(go.Scatter(
            x=categories,
            y=values,
            name=label,
            fill='tozeroy',
            mode='lines',
            line=dict(color=color),
            fillcolor=color
        ))
    
    fig.update_layout(
        title="Quarterly Performance - Area Chart",
        xaxis_title="Quarter",
        yaxis_title="Value ($M)",
        showlegend=True,
        height=500
    )
    
    # Save the chart
    fig.write_html("test_area_chart.html")
    print("✅ Area chart saved as test_area_chart.html")

if __name__ == "__main__":
    print("🧪 Testing chart generation with provided configurations...")
    
    try:
        test_stacked_column_chart()
        test_pie_chart()
        test_expanded_pie_chart()
        test_area_chart()
        
        print("\n🎉 All chart tests completed successfully!")
        print("📁 Check the generated HTML files in the current directory")
        
    except Exception as e:
        print(f"❌ Error during testing: {e}")
        import traceback
        traceback.print_exc() 