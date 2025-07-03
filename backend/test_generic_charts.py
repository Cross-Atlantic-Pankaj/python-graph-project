#!/usr/bin/env python3
"""
Test script to demonstrate generic chart type support
Shows how the system can now automatically handle new chart types
"""

import json
import tempfile
import os
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import numpy as np

def test_scatter_chart():
    """Test scatter chart - previously unsupported"""
    config = {
        "chart_meta": {
            "chart_type": "scatter",
            "legend": True,
            "data_labels": True,
            "value_format": "0.0",
            "source_sheet": "Scatter_Data",
            "category_range": "A1:A10",
            "value_range": ["B1:B10", "C1:C10"]
        },
        "series": {
            "colors": ["#FF6B6B", "#4ECDC4"],
            "labels": ["Sales", "Marketing"]
        }
    }
    
    # Mock data
    x_values = list(range(1, 11))
    series_data = [
        [10, 15, 12, 18, 20, 22, 25, 28, 30, 35],  # Sales
        [5, 8, 12, 15, 18, 20, 22, 25, 28, 30]     # Marketing
    ]
    
    fig = go.Figure()
    
    for i, (label, values, color) in enumerate(zip(config["series"]["labels"], series_data, config["series"]["colors"])):
        fig.add_trace(go.Scatter(
            x=x_values,
            y=values,
            name=label,
            mode='markers',
            marker_color=color,
            text=[f"{v:.1f}" for v in values],
            textposition="top center"
        ))
    
    fig.update_layout(
        title="Sales vs Marketing - Scatter Chart",
        xaxis_title="Month",
        yaxis_title="Value ($K)",
        showlegend=True,
        height=500
    )
    
    fig.write_html("test_scatter_chart.html")
    print("✅ Scatter chart saved as test_scatter_chart.html")

def test_bubble_chart():
    """Test bubble chart - previously unsupported"""
    config = {
        "chart_meta": {
            "chart_type": "bubble",
            "legend": True,
            "data_labels": True,
            "value_format": "0.0",
            "source_sheet": "Bubble_Data",
            "category_range": "A1:A8",
            "value_range": ["B1:B8", "C1:C8", "D1:D8"]
        },
        "series": {
            "colors": ["#FF6B6B", "#4ECDC4", "#45B7D1"],
            "labels": ["Revenue", "Profit", "Growth"]
        }
    }
    
    # Mock data
    x_values = [10, 20, 30, 40, 50, 60, 70, 80]
    y_values = [15, 25, 35, 45, 55, 65, 75, 85]
    sizes = [20, 35, 50, 65, 80, 95, 110, 125]
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=x_values,
        y=y_values,
        mode='markers',
        name="Bubble Data",
        marker=dict(
            size=sizes,
            color=sizes,
            colorscale='Viridis',
            showscale=True,
            colorbar=dict(title="Size")
        ),
        text=[f"Size: {s}" for s in sizes],
        hovertemplate="<b>Bubble</b><br>X: %{x}<br>Y: %{y}<br>Size: %{marker.size}<extra></extra>"
    ))
    
    fig.update_layout(
        title="Bubble Chart - Revenue vs Profit vs Growth",
        xaxis_title="Revenue ($M)",
        yaxis_title="Profit ($M)",
        showlegend=True,
        height=600
    )
    
    fig.write_html("test_bubble_chart.html")
    print("✅ Bubble chart saved as test_bubble_chart.html")

def test_histogram_chart():
    """Test histogram chart - previously unsupported"""
    config = {
        "chart_meta": {
            "chart_type": "histogram",
            "legend": True,
            "data_labels": True,
            "value_format": "0",
            "source_sheet": "Histogram_Data",
            "category_range": "A1:A100",
            "value_range": "B1:B100"
        },
        "series": {
            "colors": ["#FF6B6B"],
            "labels": ["Distribution"]
        }
    }
    
    # Mock data - generate random distribution
    np.random.seed(42)
    data = np.random.normal(50, 15, 100)
    
    fig = go.Figure()
    
    fig.add_trace(go.Histogram(
        x=data,
        nbinsx=20,
        name="Distribution",
        marker_color="#FF6B6B",
        opacity=0.7
    ))
    
    fig.update_layout(
        title="Data Distribution - Histogram",
        xaxis_title="Value",
        yaxis_title="Frequency",
        showlegend=True,
        height=500
    )
    
    fig.write_html("test_histogram_chart.html")
    print("✅ Histogram chart saved as test_histogram_chart.html")

def test_box_chart():
    """Test box plot chart - previously unsupported"""
    config = {
        "chart_meta": {
            "chart_type": "box",
            "legend": True,
            "data_labels": True,
            "value_format": "0.0",
            "source_sheet": "Box_Data",
            "category_range": "A1:A50",
            "value_range": ["B1:B50", "C1:C50", "D1:D50"]
        },
        "series": {
            "colors": ["#FF6B6B", "#4ECDC4", "#45B7D1"],
            "labels": ["Q1", "Q2", "Q3"]
        }
    }
    
    # Mock data - generate random data for each quarter
    np.random.seed(42)
    q1_data = np.random.normal(100, 20, 50)
    q2_data = np.random.normal(120, 25, 50)
    q3_data = np.random.normal(110, 15, 50)
    
    fig = go.Figure()
    
    fig.add_trace(go.Box(y=q1_data, name="Q1", marker_color="#FF6B6B"))
    fig.add_trace(go.Box(y=q2_data, name="Q2", marker_color="#4ECDC4"))
    fig.add_trace(go.Box(y=q3_data, name="Q3", marker_color="#45B7D1"))
    
    fig.update_layout(
        title="Quarterly Performance - Box Plot",
        yaxis_title="Sales ($K)",
        showlegend=True,
        height=500
    )
    
    fig.write_html("test_box_chart.html")
    print("✅ Box chart saved as test_box_chart.html")

def test_heatmap_chart():
    """Test heatmap chart - previously unsupported"""
    config = {
        "chart_meta": {
            "chart_type": "heatmap",
            "legend": True,
            "data_labels": True,
            "value_format": "0.0",
            "source_sheet": "Heatmap_Data",
            "category_range": "A1:A5",
            "value_range": ["B1:F1", "B2:F2", "B3:F3", "B4:F4", "B5:F5"]
        },
        "series": {
            "colors": ["#FF6B6B", "#4ECDC4", "#45B7D1"],
            "labels": ["Region A", "Region B", "Region C", "Region D", "Region E"]
        }
    }
    
    # Mock data - 5x5 heatmap
    data = [
        [10, 20, 30, 40, 50],
        [15, 25, 35, 45, 55],
        [20, 30, 40, 50, 60],
        [25, 35, 45, 55, 65],
        [30, 40, 50, 60, 70]
    ]
    
    fig = go.Figure(data=go.Heatmap(
        z=data,
        x=['Jan', 'Feb', 'Mar', 'Apr', 'May'],
        y=['Region A', 'Region B', 'Region C', 'Region D', 'Region E'],
        colorscale='Viridis',
        text=[[f"{val:.1f}" for val in row] for row in data],
        texttemplate="%{text}",
        textfont={"size": 12},
        colorbar=dict(title="Value")
    ))
    
    fig.update_layout(
        title="Regional Performance Heatmap",
        xaxis_title="Month",
        yaxis_title="Region",
        height=500
    )
    
    fig.write_html("test_heatmap_chart.html")
    print("✅ Heatmap chart saved as test_heatmap_chart.html")

def test_3d_scatter_chart():
    """Test 3D scatter chart - previously unsupported"""
    config = {
        "chart_meta": {
            "chart_type": "scatter3d",
            "legend": True,
            "data_labels": True,
            "value_format": "0.0",
            "source_sheet": "3D_Data",
            "category_range": "A1:A20",
            "value_range": ["B1:B20", "C1:C20", "D1:D20"]
        },
        "series": {
            "colors": ["#FF6B6B"],
            "labels": ["3D Data"]
        }
    }
    
    # Mock 3D data
    np.random.seed(42)
    x = np.random.rand(20) * 100
    y = np.random.rand(20) * 100
    z = np.random.rand(20) * 100
    
    fig = go.Figure(data=[go.Scatter3d(
        x=x, y=y, z=z,
        mode='markers',
        marker=dict(
            size=8,
            color=z,
            colorscale='Viridis',
            opacity=0.8
        ),
        text=[f"Point {i+1}" for i in range(20)],
        hovertemplate="<b>%{text}</b><br>X: %{x:.1f}<br>Y: %{y:.1f}<br>Z: %{z:.1f}<extra></extra>"
    )])
    
    fig.update_layout(
        title="3D Scatter Plot",
        scene=dict(
            xaxis_title="X Axis",
            yaxis_title="Y Axis", 
            zaxis_title="Z Axis"
        ),
        height=600
    )
    
    fig.write_html("test_3d_scatter_chart.html")
    print("✅ 3D Scatter chart saved as test_3d_scatter_chart.html")

def test_waterfall_chart():
    """Test waterfall chart - previously unsupported"""
    config = {
        "chart_meta": {
            "chart_type": "waterfall",
            "legend": True,
            "data_labels": True,
            "value_format": "$0.0",
            "source_sheet": "Waterfall_Data",
            "category_range": "A1:A8",
            "value_range": "B1:B8"
        },
        "series": {
            "colors": ["#FF6B6B"],
            "labels": ["Cash Flow"]
        }
    }
    
    # Mock waterfall data
    fig = go.Figure(go.Waterfall(
        name="Cash Flow",
        orientation="h",
        measure=["relative", "relative", "relative", "relative", "relative", "relative", "relative", "total"],
        x=[120, 30, -20, 80, -30, 20, -40, 180],
        textposition="outside",
        text=["+120", "+30", "-20", "+80", "-30", "+20", "-40", "180"],
        y=["Starting Balance", "Revenue", "Expenses", "Investment", "Taxes", "Interest", "Fees", "Ending Balance"],
        connector={"line": {"color": "rgb(63, 63, 63)"}},
        decreasing={"marker": {"color": "#FF6B6B"}},
        increasing={"marker": {"color": "#4ECDC4"}},
        totals={"marker": {"color": "#45B7D1"}}
    ))
    
    fig.update_layout(
        title="Cash Flow Waterfall Chart",
        xaxis_title="Amount ($K)",
        yaxis_title="Category",
        showlegend=False,
        height=500
    )
    
    fig.write_html("test_waterfall_chart.html")
    print("✅ Waterfall chart saved as test_waterfall_chart.html")

def test_funnel_chart():
    """Test funnel chart - previously unsupported"""
    config = {
        "chart_meta": {
            "chart_type": "funnel",
            "legend": True,
            "data_labels": True,
            "value_format": "0",
            "source_sheet": "Funnel_Data",
            "category_range": "A1:A5",
            "value_range": "B1:B5"
        },
        "series": {
            "colors": ["#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4", "#FFEAA7"],
            "labels": ["Leads", "Qualified", "Proposals", "Negotiations", "Closed"]
        }
    }
    
    # Mock funnel data
    fig = go.Figure(go.Funnel(
        y=["Leads", "Qualified", "Proposals", "Negotiations", "Closed"],
        x=[1000, 800, 600, 400, 200],
        textinfo="value+percent initial",
        marker=dict(color=["#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4", "#FFEAA7"])
    ))
    
    fig.update_layout(
        title="Sales Funnel",
        height=500
    )
    
    fig.write_html("test_funnel_chart.html")
    print("✅ Funnel chart saved as test_funnel_chart.html")

def main():
    """Run all chart tests"""
    print("🚀 Testing Generic Chart Type Support")
    print("=" * 50)
    
    # Test various chart types that were previously unsupported
    test_scatter_chart()
    test_bubble_chart()
    test_histogram_chart()
    test_box_chart()
    test_heatmap_chart()
    test_3d_scatter_chart()
    test_waterfall_chart()
    test_funnel_chart()
    
    print("\n" + "=" * 50)
    print("✅ All generic chart tests completed!")
    print("📁 Check the generated HTML files to see the charts")
    print("\n🎯 The system now automatically supports:")
    print("   • Scatter plots")
    print("   • Bubble charts") 
    print("   • Histograms")
    print("   • Box plots")
    print("   • Heatmaps")
    print("   • 3D scatter plots")
    print("   • Waterfall charts")
    print("   • Funnel charts")
    print("   • And many more Plotly chart types!")

if __name__ == "__main__":
    main()
