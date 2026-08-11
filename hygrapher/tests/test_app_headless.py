# -*- coding: utf-8 -*-
"""
Granular PyQt6 Headless Test Suite for HyGrapher 2D & 3D.
"""

import os

os.environ["QT_QPA_PLATFORM"] = "offscreen"

import pytest

from PyQt6.QtWidgets import QApplication
import matplotlib

matplotlib.use("Agg")

from hygrapher.main import GraphApp as GraphApp2D
from hygrapher.main_3d import GraphApp3D


@pytest.fixture(scope="session", autouse=True)
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    yield app
    app.processEvents()


@pytest.fixture(scope="module")
def sample_csv(tmp_path_factory):
    p = tmp_path_factory.mktemp("data") / "sample.csv"
    p.write_text(
        "Time,Val1,Val2,Err\n1,10,100,0.5\n2,25,200,1.0\n3,15,300,1.5\n4,40,250,2.0\n5,30,500,2.5\n"
    )
    return str(p)


@pytest.fixture(scope="module")
def app2d(sample_csv):
    app = GraphApp2D()
    app.load_data(file_path=sample_csv)
    init_tab = app.x_tab_widgets[0]
    init_tab["y1_listbox"].item(1).setSelected(True)  # Val1
    init_tab["y2_listbox"].item(2).setSelected(True)  # Val2
    yield app
    app.close()
    app.deleteLater()
    QApplication.processEvents()


@pytest.fixture(scope="module")
def app3d(sample_csv):
    app = GraphApp3D()
    app.load_data(file_path=sample_csv)
    if app.z_listbox.count() >= 3:
        app.x_axis_combo.setCurrentIndex(0)
        app.y_axis_combo.setCurrentIndex(1)
        app.z_listbox.item(2).setSelected(True)
    app.resolution_spin.setValue(10)
    yield app
    app.close()
    app.deleteLater()
    QApplication.processEvents()


# ── 2D Plot Types Parameterized Tests (12 Tests) ──────────────────────────────
@pytest.mark.parametrize(
    "plot_type",
    [
        "line",
        "scatter",
        "bar",
        "step",
        "stem",
        "area",
        "pie",
        "box",
        "violin",
        "heatmap",
        "contour",
        "polar",
    ],
)
def test_2d_plot_type(app2d, plot_type):
    idx = app2d.plot_type_combo.findText(plot_type)
    assert idx >= 0
    app2d.plot_type_combo.setCurrentIndex(idx)
    if plot_type == "contour":
        init_tab = app2d.x_tab_widgets[0]
        init_tab["y1_listbox"].item(1).setSelected(True)
        init_tab["y1_listbox"].item(2).setSelected(True)
    app2d.plot_graph()


# ── 3D Plot Types Parameterized Tests (5 Tests) ───────────────────────────────
@pytest.mark.parametrize(
    "plot_type", ["surface", "wireframe", "contour3d", "scatter3d", "line3d"]
)
def test_3d_plot_type(app3d, plot_type):
    idx = app3d.plot_type_combo.findText(plot_type)
    assert idx >= 0
    app3d.plot_type_combo.setCurrentIndex(idx)
    app3d.plot_graph()


# ── Series Style Editor Tests ────────────────────────────────────────────────
def test_2d_style_editor(app2d):
    app2d.update_style_editor_targets()
    assert app2d.combined_style_target_combo.count() > 0
    app2d.combined_style_target_combo.setCurrentIndex(0)
    app2d.on_combined_series_select()
    app2d.style_color_input.setText("#FF0000")
    app2d.style_linestyle_combo.setCurrentIndex(1)
    app2d.style_linewidth_spin.setValue(3.0)
    app2d.on_style_editor_change()
    app2d.plot_graph()

    app2d.on_style_editor_color_auto()
    assert app2d.style_color_input.text() == "Auto"


# ── Advanced Data Filtering Tests ─────────────────────────────────────────────
def test_2d_advanced_data_filter(app2d):
    app2d.data_filter_check.setChecked(True)
    app2d.filter_column_combo.setCurrentIndex(1)  # Val1
    app2d.filter_min_input.setText("15")
    app2d.filter_max_input.setText("35")
    app2d.plot_graph()

    df_filtered = app2d.data_mgr.get_filtered_df(
        filter_enabled=True, filter_column="Val1", min_val_str="15", max_val_str="35"
    )
    assert len(df_filtered) == 3


# ── Line Smoothing & Annotations Tests ────────────────────────────────────────
def test_2d_smoothing_and_annotations(app2d):
    app2d.enable_smoothing_check.setChecked(True)
    app2d.smoothing_window_spin.setValue(2)
    app2d.enable_annotation_check.setChecked(True)
    app2d.enable_errorbar_check.setChecked(True)
    app2d.plot_graph()


# ── Subplot Mode & Inversion Tests ────────────────────────────────────────────
def test_2d_subplot_and_axis_inversion(app2d):
    app2d.subplot_mode_check.setChecked(True)
    app2d.x_log_check.setChecked(True)
    app2d.y1_log_check.setChecked(True)
    app2d.y1_invert_check.setChecked(True)
    app2d.y2_invert_check.setChecked(True)
    app2d.plot_graph()


# ── Spines & Grid Styling Tests ───────────────────────────────────────────────
def test_2d_spines_and_grid(app2d):
    app2d.spine_top_check.setChecked(False)
    app2d.spine_right_check.setChecked(False)
    app2d.grid_check.setChecked(False)
    app2d.rotate_labels_check.setChecked(True)
    app2d.rotation_angle_spin.setValue(90)
    app2d.plot_graph()


# ── X-Tab Management Tests ───────────────────────────────────────────────────
def test_2d_multi_x_tab_management(app2d):
    assert len(app2d.x_tab_widgets) == 1
    app2d.add_x_tab()
    assert len(app2d.x_tab_widgets) == 2
    app2d.remove_x_tab(app2d.x_tab_widgets[1]["tab_widget"])
    assert len(app2d.x_tab_widgets) == 1


# ── Table View Data Editing Tests ─────────────────────────────────────────────
def test_2d_table_editing(app2d):
    assert app2d.data_table.rowCount() == 5
    item = app2d.data_table.item(0, 1)
    item.setText("999")
    app2d.get_data_from_table()
    assert app2d.df.iat[0, 1] == "999"


# ── Project File I/O Tests ────────────────────────────────────────────────────
def test_2d_project_save_load(app2d, tmp_path):
    app2d.title_input.setText("Round Trip Title")
    app2d.plot_type_combo.setCurrentIndex(app2d.plot_type_combo.findText("scatter"))
    app2d.grid_check.setChecked(False)
    app2d.spine_top_check.setChecked(False)
    app2d.x_log_check.setChecked(True)

    proj_path = tmp_path / "test_project.pmggrp"
    app2d.current_project_path = str(proj_path)
    app2d.overwrite_save()
    assert proj_path.exists()

    # Clear the widgets so the load actually has to restore them.
    app2d.title_input.setText("")
    app2d.plot_type_combo.setCurrentIndex(0)
    app2d.grid_check.setChecked(True)
    app2d.spine_top_check.setChecked(True)
    app2d.x_log_check.setChecked(False)

    app2d.load_project_file(str(proj_path))
    assert app2d.df is not None
    assert app2d.title_input.text() == "Round Trip Title"
    assert app2d.plot_type_combo.currentText() == "scatter"
    assert app2d.grid_check.isChecked() is False
    assert app2d.spine_top_check.isChecked() is False
    assert app2d.x_log_check.isChecked() is True


def test_3d_project_save_load(app3d, tmp_path):
    proj_path = tmp_path / "test_project_3d.pmggrp"
    app3d.current_project_path = str(proj_path)
    app3d.overwrite_save()
    assert proj_path.exists()

    app3d.load_project_file(str(proj_path))
    assert app3d.df is not None


# ── Reset & Clear All Tests ──────────────────────────────────────────────────
def test_2d_reset_and_clear(app2d):
    app2d.title_input.setText("Some Title")
    app2d.reset_settings()
    assert app2d.title_input.text() == ""
    assert app2d.df is None


def test_3d_reset_and_clear(app3d):
    app3d.elev_spin.setValue(45)
    app3d.reset_settings()
    assert app3d.elev_spin.value() == 30
    assert app3d.df is None
