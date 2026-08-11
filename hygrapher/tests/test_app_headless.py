# -*- coding: utf-8 -*-
"""
Headless PyQt6 application tests for HyGrapher 2D & 3D.
"""

import os
os.environ["QT_QPA_PLATFORM"] = "offscreen"

import pytest
import pandas as pd

from PyQt6.QtWidgets import QApplication
import matplotlib
matplotlib.use('Agg')

from hygrapher.main import GraphApp as GraphApp2D
from hygrapher.main_3d import GraphApp as GraphApp3D


@pytest.fixture(scope="session", autouse=True)
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    yield app
    app.processEvents()


@pytest.fixture
def sample_csv(tmp_path):
    p = tmp_path / "sample.csv"
    p.write_text("Time,Val1,Val2,Err\n1,10,100,0.5\n2,25,200,1.0\n3,15,300,1.5\n4,40,250,2.0\n5,30,500,2.5\n")
    return str(p)


def test_2d_app_pyqt6(sample_csv, tmp_path):
    app = GraphApp2D()

    # Load data
    app.load_data(file_path=sample_csv)
    assert app.df is not None
    assert len(app.df) == 5

    # Multi X-tabs
    app.add_x_tab()
    assert len(app.x_tab_widgets) == 2

    # Set variables & options
    app.title_input.setText("Headless 2D Test")
    app.xlabel_input.setText("Time (s)")
    app.ylabel_input.setText("Val1 Label")
    app.ylabel2_input.setText("Val2 Label")

    init_tab = app.x_tab_widgets[0]
    init_tab["y1_listbox"].item(1).setSelected(True)  # Val1
    init_tab["y2_listbox"].item(2).setSelected(True)  # Val2

    # Test plot types
    for p_type in ["line", "scatter", "bar", "step", "stem", "area", "pie", "box", "violin", "heatmap", "polar"]:
        idx = app.plot_type_combo.findText(p_type)
        if idx >= 0:
            app.plot_type_combo.setCurrentIndex(idx)
            app.plot_graph()

    # Style editor interactions
    app.update_style_editor_targets()
    if app.combined_style_target_combo.count() > 0:
        app.combined_style_target_combo.setCurrentIndex(0)
        app.on_combined_series_select()
        app.style_color_input.setText("#FF0000")
        app.on_style_editor_change()
        app.on_style_editor_color_auto()

    # Advanced options
    app.enable_smoothing_check.setChecked(True)
    app.smoothing_window_spin.setValue(2)
    app.enable_errorbar_check.setChecked(True)
    app.enable_annotation_check.setChecked(True)
    app.x_log_check.setChecked(True)
    app.y1_log_check.setChecked(True)
    app.data_filter_check.setChecked(True)
    app.filter_min_input.setText("15")
    app.filter_max_input.setText("35")

    app.plot_graph()

    # Export graph & data
    exp_graph_path = tmp_path / "output.png"
    app.fig.savefig(str(exp_graph_path))
    assert exp_graph_path.exists()

    exp_data_path = tmp_path / "filtered.csv"
    df_filtered = app.data_mgr.get_filtered_df(filter_enabled=True, filter_column="Val1", min_val_str="15", max_val_str="35")
    df_filtered.to_csv(str(exp_data_path), index=False)
    assert exp_data_path.exists()

    # Save project
    proj_path = tmp_path / "project.pmggrp"
    app.current_project_path = str(proj_path)
    app.overwrite_save()
    assert proj_path.exists()

    # Tab removal & reset/clear
    app.remove_x_tab(app.x_tab_widgets[1]["tab_widget"])
    assert len(app.x_tab_widgets) == 1

    app.reset_settings()
    app.clear_all()
    app.close()
    app.deleteLater()
    QApplication.processEvents()


def test_3d_app_pyqt6(sample_csv, tmp_path):
    app = GraphApp3D()

    app.load_data(file_path=sample_csv)
    assert app.df is not None

    if app.z_listbox.count() >= 3:
        app.x_axis_combo.setCurrentIndex(0)
        app.y_axis_combo.setCurrentIndex(1)
        app.z_listbox.item(2).setSelected(True)

    app.resolution_spin.setValue(10)

    for p_type in ["surface", "wireframe", "contour3d", "scatter3d", "line3d"]:
        idx = app.plot_type_combo.findText(p_type)
        if idx >= 0:
            app.plot_type_combo.setCurrentIndex(idx)
            app.plot_graph()

    proj_path = tmp_path / "project_3d.pmggrp"
    app.current_project_path = str(proj_path)
    app.overwrite_save()
    assert proj_path.exists()

    app.reset_settings()
    app.clear_all()
    app.close()
    app.deleteLater()
    QApplication.processEvents()
