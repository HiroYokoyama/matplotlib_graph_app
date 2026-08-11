# -*- coding: utf-8 -*-
import pytest
import os
import json
import pandas as pd
import tkinter as tk
from unittest.mock import patch, MagicMock

import matplotlib
matplotlib.use('Agg')

from hygrapher.main import GraphApp as GraphApp2D
from hygrapher.main_3d import GraphApp as GraphApp3D


@pytest.fixture
def sample_csv(tmp_path):
    p = tmp_path / "sample.csv"
    p.write_text("Time,Val1,Val2,Err\n1,10,100,0.5\n2,20,200,1.0\n3,30,300,1.5\n4,40,400,2.0\n")
    return str(p)


def test_2d_app_full_coverage(sample_csv, tmp_path):
    try:
        app = GraphApp2D()
    except tk.TclError:
        pytest.skip("Tkinter display not available")

    # Load data
    app.load_data(file_path=sample_csv)
    assert app.df is not None
    assert len(app.df) == 4

    # Multi X-tabs
    app.add_x_tab()
    assert len(app.x_tab_widgets) == 2

    # Set variables & options
    app.x_axis_var.set("Time")
    app.title_var.set("Headless 2D Test")
    app.xlabel_var.set("Time (s)")
    app.ylabel_var.set("Val1 Label")
    app.ylabel2_var.set("Val2 Label")

    init_tab = app.x_tab_widgets[0]
    init_tab["y1_listbox"].select_set(0)  # Val1
    init_tab["y2_listbox"].select_set(1)  # Val2

    # Test plot types: line, scatter, bar, step, area
    for p_type in ["line", "scatter", "bar", "step", "area"]:
        app.plot_type_var.set(p_type)
        app.plot_graph()

    # Style editor interactions
    app.combined_style_target_var.set("(T1-Y1) Val1")
    app.on_combined_series_select()
    app.current_style_color_var.set("#FF0000")
    app.current_style_linestyle_var.set("--")
    app.current_style_linewidth_var.set(2.5)
    app.on_style_editor_change()
    app.on_style_editor_color_auto()

    # Advanced options: smoothing, errorbar, log scale, minor ticks, data filter
    app.enable_smoothing_var.set(True)
    app.smoothing_window_var.set(2)
    app.enable_errorbar_var.set(True)
    app.errorbar_column_var.set("Err")
    app.enable_annotation_var.set(True)
    app.x_log_scale_var.set(True)
    app.y1_log_scale_var.set(True)
    app.xtick_minor_show_var.set(True)
    app.data_filter_enabled_var.set(True)
    app.filter_column_var.set("Val1")
    app.filter_min_var.set("15")
    app.filter_max_var.set("35")

    app.plot_graph()

    # Export graph & data
    exp_graph_path = tmp_path / "output.png"
    app.fig.savefig(str(exp_graph_path))
    assert exp_graph_path.exists()

    exp_data_path = tmp_path / "filtered.csv"
    with patch("tkinter.filedialog.asksaveasfilename", return_value=str(exp_data_path)):
        app.export_filtered_data()
    assert exp_data_path.exists()

    # Save and Load project
    proj_path = tmp_path / "project.pmggrp"
    app.current_project_path = str(proj_path)
    app.overwrite_save()
    assert proj_path.exists()

    # Drag & Drop handling
    class MockDropEvent:
        data = str(proj_path)

    app.on_drop(MockDropEvent())

    # Tab removal & reset/clear
    app.remove_x_tab(app.x_tab_widgets[1]["tab_frame"])
    assert len(app.x_tab_widgets) == 1

    with patch("tkinter.messagebox.askyesno", return_value=True):
        app.reset_settings()
        app.clear_all()

    app.destroy()


def test_3d_app_full_coverage(sample_csv, tmp_path):
    try:
        app = GraphApp3D()
    except tk.TclError:
        pytest.skip("Tkinter display not available")

    app.load_data(file_path=sample_csv)
    assert app.df is not None

    app.x_axis_var.set("Time")
    app.y_axis_var.set("Val1")
    app.y_listbox.select_set(0)  # Val2

    for p_type in ["surface", "wireframe", "scatter3d", "line3d"]:
        app.plot_type_var.set(p_type)
        app.plot_graph()

    proj_path = tmp_path / "project_3d.pmggrp"
    app.current_project_path = str(proj_path)
    app.overwrite_save()
    assert proj_path.exists()

    class MockDropEvent:
        data = str(proj_path)

    app.on_drop(MockDropEvent())

    with patch("tkinter.messagebox.askyesno", return_value=True):
        app.reset_settings()
        app.clear_all()

    app.destroy()
