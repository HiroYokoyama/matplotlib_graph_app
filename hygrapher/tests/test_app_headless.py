# -*- coding: utf-8 -*-
"""
Granular PyQt6 Headless Test Suite for HyGrapher 2D & 3D.
"""

import os

os.environ["QT_QPA_PLATFORM"] = "offscreen"

import pytest
from unittest.mock import MagicMock

from PyQt6.QtCore import QUrl
from PyQt6.QtGui import QCloseEvent
from PyQt6.QtWidgets import QApplication, QDialog, QInputDialog, QMessageBox
import matplotlib

matplotlib.use("Agg")

from hygrapher.main import GraphApp as GraphApp2D
from hygrapher.main_3d import GraphApp3D
from hygrapher.import_dialog import ImportPreviewDialog


def _fake_drop_event(file_path):
    """A MagicMock standing in for QDropEvent (real QDropEvent objects
    can't be constructed directly from Python)."""
    event = MagicMock()
    event.mimeData.return_value.hasUrls.return_value = True
    event.mimeData.return_value.urls.return_value = [QUrl.fromLocalFile(file_path)]
    return event


@pytest.fixture
def auto_accept_import_dialog(monkeypatch):
    """ImportPreviewDialog.exec() opens a real modal event loop, which would
    hang forever in a headless test with nothing to click Accept. Patch it to
    accept immediately with whatever header_row the test set up beforehand
    (default: 0, from the dialog's own __init__)."""
    monkeypatch.setattr(
        ImportPreviewDialog, "exec", lambda self: QDialog.DialogCode.Accepted
    )


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
    # app2d is module-scoped and shared; pin a rectangular plot type since a
    # prior parametrized test may have left it on "polar" (different spines).
    app2d.plot_type_combo.setCurrentIndex(app2d.plot_type_combo.findText("line"))
    app2d.spine_top_check.setChecked(False)
    app2d.spine_right_check.setChecked(False)
    app2d.grid_check.setChecked(False)
    app2d.rotate_labels_check.setChecked(True)
    app2d.rotation_angle_spin.setValue(90)
    app2d.plot_graph()

    # Regression: these settings previously had no effect at all.
    assert app2d.ax.spines["top"].get_visible() is False
    assert app2d.ax.spines["right"].get_visible() is False


def test_2d_axis_settings_actually_apply(app2d):
    """Log scale / invert / Y2 label were wired into the UI but never read
    by plot_graph(); this guards against that regressing again."""
    app2d.plot_type_combo.setCurrentIndex(app2d.plot_type_combo.findText("line"))
    app2d.x_log_check.setChecked(True)
    app2d.y1_log_check.setChecked(True)
    app2d.y1_invert_check.setChecked(True)
    app2d.ylabel2_input.setText("Custom Y2")
    app2d.plot_graph()

    assert app2d.ax.get_xscale() == "log"
    assert app2d.ax.get_yscale() == "log"
    if app2d.ax2 is not None:
        assert app2d.ax2.get_ylabel() == "Custom Y2"

    # Reset for subsequent tests in this module-scoped fixture.
    app2d.x_log_check.setChecked(False)
    app2d.y1_log_check.setChecked(False)
    app2d.y1_invert_check.setChecked(False)


def test_2d_major_tick_interval_applies(app2d):
    """apply_major_ticker() used to be called with the wrong arguments
    (an Axes instead of an Axis, and the axis name instead of the interval
    text), so the interval boxes silently did nothing."""
    import matplotlib.ticker as mticker

    app2d.x_log_check.setChecked(False)
    app2d.xtick_major_interval_input.setText("2.5")
    app2d.plot_graph()

    locator = app2d.ax.xaxis.get_major_locator()
    assert isinstance(locator, mticker.MultipleLocator)

    app2d.xtick_major_interval_input.setText("")


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


# ── Row/Column Insert & Delete Tests ─────────────────────────────────────────
def test_2d_insert_and_delete_row_is_undoable(sample_csv):
    app = GraphApp2D()
    try:
        app.load_data(file_path=sample_csv)
        original_rows = app.data_table.rowCount()

        app.data_table.setCurrentCell(1, 0)
        app.insert_row_above()
        assert app.data_table.rowCount() == original_rows + 1
        assert len(app.df) == original_rows + 1

        app.undo_stack.undo()
        assert app.data_table.rowCount() == original_rows
        assert len(app.df) == original_rows

        app.data_table.setCurrentCell(0, 0)
        app.delete_selected_row()
        assert app.data_table.rowCount() == original_rows - 1

        app.undo_stack.undo()
        assert app.data_table.rowCount() == original_rows
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_2d_cannot_delete_last_row(sample_csv, monkeypatch):
    app = GraphApp2D()
    try:
        app.load_data(file_path=sample_csv)
        warned = []
        monkeypatch.setattr(QMessageBox, "warning", lambda *a, **k: warned.append(1))

        while app.data_table.rowCount() > 1:
            app.data_table.setCurrentCell(0, 0)
            app.delete_selected_row()
        app.data_table.setCurrentCell(0, 0)
        app.delete_selected_row()

        assert app.data_table.rowCount() == 1
        assert warned
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_2d_insert_and_delete_column_is_undoable(sample_csv, monkeypatch):
    app = GraphApp2D()
    try:
        app.load_data(file_path=sample_csv)
        original_cols = app.data_table.columnCount()

        monkeypatch.setattr(
            QInputDialog, "getText", staticmethod(lambda *a, **k: ("NewCol", True))
        )
        app.data_table.setCurrentCell(0, 0)
        app.insert_column_left()
        assert app.data_table.columnCount() == original_cols + 1
        assert "NewCol" in app.df.columns
        x_combo = app.x_tab_widgets[0]["x_combo"]
        assert "NewCol" in [x_combo.itemText(i) for i in range(x_combo.count())]

        app.undo_stack.undo()
        assert app.data_table.columnCount() == original_cols
        assert "NewCol" not in app.df.columns

        app.data_table.setCurrentCell(0, 1)
        app.delete_selected_column()
        assert app.data_table.columnCount() == original_cols - 1

        app.undo_stack.undo()
        assert app.data_table.columnCount() == original_cols
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_2d_insert_column_cancelled_dialog_does_nothing(sample_csv, monkeypatch):
    app = GraphApp2D()
    try:
        app.load_data(file_path=sample_csv)
        original_cols = app.data_table.columnCount()
        monkeypatch.setattr(
            QInputDialog, "getText", staticmethod(lambda *a, **k: ("", False))
        )
        app.data_table.setCurrentCell(0, 0)
        app.insert_column_left()
        assert app.data_table.columnCount() == original_cols
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_2d_insert_column_duplicate_name_rejected(sample_csv, monkeypatch):
    """A duplicate column name would make df[col] return a DataFrame
    instead of a Series, crashing plot_graph() with AttributeError on
    .str; must be rejected instead of silently corrupting the data."""
    app = GraphApp2D()
    try:
        app.load_data(file_path=sample_csv)
        original_cols = app.data_table.columnCount()
        existing_name = app.data_table.horizontalHeaderItem(0).text()

        warned = []
        monkeypatch.setattr(QMessageBox, "warning", lambda *a, **k: warned.append(1))
        monkeypatch.setattr(
            QInputDialog,
            "getText",
            staticmethod(lambda *a, **k: (existing_name, True)),
        )
        app.data_table.setCurrentCell(0, 0)
        app.insert_column_left()

        assert app.data_table.columnCount() == original_cols
        assert warned
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_2d_cannot_delete_last_column(sample_csv, monkeypatch):
    app = GraphApp2D()
    try:
        app.load_data(file_path=sample_csv)
        warned = []
        monkeypatch.setattr(QMessageBox, "warning", lambda *a, **k: warned.append(1))

        while app.data_table.columnCount() > 1:
            app.data_table.setCurrentCell(0, 0)
            app.delete_selected_column()
        app.data_table.setCurrentCell(0, 0)
        app.delete_selected_column()

        assert app.data_table.columnCount() == 1
        assert warned
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_2d_table_context_menu_noop_without_data():
    app = GraphApp2D()
    try:
        # No columns yet: must return before menu.exec() (which would
        # otherwise block on a real modal event loop in a headless test).
        app.show_table_context_menu(app.data_table.rect().center())
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_3d_insert_and_delete_row_is_undoable(sample_csv):
    app = GraphApp3D()
    try:
        app.load_data(file_path=sample_csv)
        original_rows = app.data_table.rowCount()

        app.data_table.setCurrentCell(1, 0)
        app.insert_row_above()
        assert app.data_table.rowCount() == original_rows + 1

        app.undo_stack.undo()
        assert app.data_table.rowCount() == original_rows

        app.data_table.setCurrentCell(0, 0)
        app.delete_selected_row()
        assert app.data_table.rowCount() == original_rows - 1

        app.undo_stack.undo()
        assert app.data_table.rowCount() == original_rows
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_3d_insert_and_delete_column_is_undoable(sample_csv, monkeypatch):
    app = GraphApp3D()
    try:
        app.load_data(file_path=sample_csv)
        original_cols = app.data_table.columnCount()

        monkeypatch.setattr(
            QInputDialog, "getText", staticmethod(lambda *a, **k: ("NewCol", True))
        )
        app.data_table.setCurrentCell(0, 0)
        app.insert_column_left()
        assert app.data_table.columnCount() == original_cols + 1
        assert "NewCol" in [
            app.x_axis_combo.itemText(i) for i in range(app.x_axis_combo.count())
        ]

        app.undo_stack.undo()
        assert app.data_table.columnCount() == original_cols

        app.data_table.setCurrentCell(0, 1)
        app.delete_selected_column()
        assert app.data_table.columnCount() == original_cols - 1

        app.undo_stack.undo()
        assert app.data_table.columnCount() == original_cols
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_3d_insert_column_duplicate_name_rejected(sample_csv, monkeypatch):
    app = GraphApp3D()
    try:
        app.load_data(file_path=sample_csv)
        original_cols = app.data_table.columnCount()
        existing_name = app.data_table.horizontalHeaderItem(0).text()

        warned = []
        monkeypatch.setattr(QMessageBox, "warning", lambda *a, **k: warned.append(1))
        monkeypatch.setattr(
            QInputDialog,
            "getText",
            staticmethod(lambda *a, **k: (existing_name, True)),
        )
        app.data_table.setCurrentCell(0, 0)
        app.insert_column_left()

        assert app.data_table.columnCount() == original_cols
        assert warned
    finally:
        app.deleteLater()
        QApplication.processEvents()


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


def test_3d_load_project_file_tracks_current_path(app3d, tmp_path):
    """Regression: load_project_file() used to forget to set
    current_project_path, so overwrite_save() after a load would always
    prompt Save As instead of silently overwriting the loaded file."""
    proj_path = tmp_path / "tracked.pmggrp"
    app3d.current_project_path = str(proj_path)
    app3d.overwrite_save()

    app3d.current_project_path = None
    app3d.load_project_file(str(proj_path))
    assert app3d.current_project_path == str(proj_path)


# ── Drag & Drop Tests ─────────────────────────────────────────────────────────
def test_2d_drop_event_loads_data_file(app2d, sample_csv, auto_accept_import_dialog):
    app2d.df = None
    app2d.dropEvent(_fake_drop_event(sample_csv))
    assert app2d.df is not None


def test_2d_drop_event_loads_project_file(app2d, tmp_path):
    proj_path = tmp_path / "dropped.pmggrp"
    app2d.title_input.setText("Dropped Project")
    app2d.current_project_path = str(proj_path)
    app2d.overwrite_save()

    app2d.title_input.setText("")
    # Project files (.pmggrp) skip the import-preview dialog entirely.
    app2d.dropEvent(_fake_drop_event(str(proj_path)))
    assert app2d.title_input.text() == "Dropped Project"


def test_2d_drop_event_ignores_unsupported_extension(app2d, tmp_path):
    other_file = tmp_path / "not_supported.exe"
    other_file.write_bytes(b"\x00")
    app2d.df = None

    app2d.dropEvent(_fake_drop_event(str(other_file)))
    assert app2d.df is None  # silently ignored, no crash


def test_3d_drop_event_loads_data_file(app3d, sample_csv, auto_accept_import_dialog):
    app3d.df = None
    app3d.dropEvent(_fake_drop_event(sample_csv))
    assert app3d.df is not None


# ── Import Preview (header-row selection) Tests ─────────────────────────────
def test_2d_import_skips_title_row_above_header(app2d, tmp_path, monkeypatch):
    """A file with a title line above the real header should load correctly
    once the user (or, here, the pre-set dialog state) picks row 1 as header."""
    messy_file = tmp_path / "messy.csv"
    messy_file.write_text("Experiment Run 42\nTime,Val1,Val2\n1,10,100\n2,20,200\n")

    def fake_exec(self):
        self.header_row = 1
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(ImportPreviewDialog, "exec", fake_exec)

    app2d.df = None
    app2d.import_data_interactive(file_path=str(messy_file))

    assert app2d.df is not None
    assert list(app2d.df.columns) == ["Time", "Val1", "Val2"]
    assert len(app2d.df) == 2


def test_2d_import_no_header_auto_names_columns(app2d, tmp_path, monkeypatch):
    headerless_file = tmp_path / "headerless.csv"
    headerless_file.write_text("1,10,100\n2,20,200\n")

    def fake_exec(self):
        self.header_row = None
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(ImportPreviewDialog, "exec", fake_exec)

    app2d.df = None
    app2d.import_data_interactive(file_path=str(headerless_file))

    assert app2d.df is not None
    assert list(app2d.df.columns) == ["Column1", "Column2", "Column3"]
    assert len(app2d.df) == 2


def test_2d_import_cancelled_dialog_does_not_load(app2d, tmp_path, monkeypatch):
    data_file = tmp_path / "data.csv"
    data_file.write_text("A,B\n1,2\n")

    monkeypatch.setattr(
        ImportPreviewDialog, "exec", lambda self: QDialog.DialogCode.Rejected
    )

    app2d.df = None
    app2d.import_data_interactive(file_path=str(data_file))
    assert app2d.df is None


# ── Undo/Redo Tests ───────────────────────────────────────────────────────────
def test_2d_undo_redo_cell_edit(app2d):
    app2d.undo_stack.clear()
    item = app2d.data_table.item(0, 1)
    original_text = item.text()

    item.setText("999")
    assert app2d.undo_stack.count() == 1
    assert app2d.data_table.item(0, 1).text() == "999"

    app2d.undo_stack.undo()
    assert app2d.data_table.item(0, 1).text() == original_text

    app2d.undo_stack.redo()
    assert app2d.data_table.item(0, 1).text() == "999"

    app2d.undo_stack.undo()  # restore for later tests in this module


def test_2d_undo_does_not_record_noop_edits(app2d):
    """Re-committing the same text (e.g. clicking into a cell and pressing
    Enter without changing anything) shouldn't create undo history."""
    app2d.undo_stack.clear()
    item = app2d.data_table.item(0, 0)
    item.setText(item.text())
    assert app2d.undo_stack.count() == 0


def test_3d_undo_redo_cell_edit(app3d):
    app3d.undo_stack.clear()
    item = app3d.data_table.item(0, 0)
    original_text = item.text()

    item.setText("42")
    assert app3d.undo_stack.count() == 1

    app3d.undo_stack.undo()
    assert app3d.data_table.item(0, 0).text() == original_text
    app3d.undo_stack.redo()
    assert app3d.data_table.item(0, 0).text() == "42"
    app3d.undo_stack.undo()


# ── Window Title Tests ───────────────────────────────────────────────────────
def test_2d_window_title_shows_unsaved_marker_after_edit(app2d):
    app2d.undo_stack.clear()
    app2d.update_window_title()
    assert "*" not in app2d.windowTitle()

    item = app2d.data_table.item(0, 0)
    old_text = item.text()
    item.setText(old_text + "_edited")
    assert "*" in app2d.windowTitle()

    app2d.undo_stack.undo()
    assert "*" not in app2d.windowTitle()


def test_2d_window_title_updates_on_project_save_and_load(tmp_path):
    app = GraphApp2D()
    try:
        assert app.current_project_path is None
        assert "HyGrapher" in app.windowTitle()

        proj_path = tmp_path / "titled_project.pmggrp"
        app.current_project_path = str(proj_path)
        app.overwrite_save()
        assert proj_path.name in app.windowTitle()
        assert "*" not in app.windowTitle()

        app.current_project_path = None
        app.load_project_file(str(proj_path))
        assert proj_path.name in app.windowTitle()
        assert app.current_project_path == str(proj_path)
    finally:
        app.close()
        app.deleteLater()
        QApplication.processEvents()


def test_3d_window_title_updates_on_project_save_and_load(tmp_path):
    app = GraphApp3D()
    try:
        assert app.current_project_path is None
        assert "HyGrapher" in app.windowTitle()

        proj_path = tmp_path / "titled_3d_project.pmggrp"
        app.current_project_path = str(proj_path)
        app.overwrite_save()
        assert proj_path.name in app.windowTitle()

        app.current_project_path = None
        app.load_project_file(str(proj_path))
        assert proj_path.name in app.windowTitle()
    finally:
        app.close()
        app.deleteLater()
        QApplication.processEvents()


# ── Unsaved-Changes-on-Close Tests ──────────────────────────────────────────
def test_2d_close_with_no_unsaved_changes_closes_immediately(sample_csv):
    app = GraphApp2D()
    try:
        app.load_data(file_path=sample_csv)
        assert app.undo_stack.isClean()

        event = QCloseEvent()
        app.closeEvent(event)
        assert event.isAccepted()
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_2d_close_with_unsaved_changes_cancel_keeps_window_open(
    sample_csv, monkeypatch
):
    app = GraphApp2D()
    try:
        app.load_data(file_path=sample_csv)
        app.data_table.item(0, 0).setText("edited")
        assert not app.undo_stack.isClean()

        monkeypatch.setattr(
            QMessageBox,
            "question",
            lambda *a, **k: QMessageBox.StandardButton.Cancel,
        )
        event = QCloseEvent()
        app.closeEvent(event)
        assert not event.isAccepted()
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_2d_close_with_unsaved_changes_discard_closes(sample_csv, monkeypatch):
    app = GraphApp2D()
    try:
        app.load_data(file_path=sample_csv)
        app.data_table.item(0, 0).setText("edited")

        monkeypatch.setattr(
            QMessageBox,
            "question",
            lambda *a, **k: QMessageBox.StandardButton.Discard,
        )
        event = QCloseEvent()
        app.closeEvent(event)
        assert event.isAccepted()
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_2d_close_with_unsaved_changes_save_then_closes(
    sample_csv, tmp_path, monkeypatch
):
    app = GraphApp2D()
    try:
        app.load_data(file_path=sample_csv)
        app.data_table.item(0, 0).setText("edited")
        app.current_project_path = str(tmp_path / "autosave.pmggrp")

        monkeypatch.setattr(
            QMessageBox, "question", lambda *a, **k: QMessageBox.StandardButton.Save
        )
        event = QCloseEvent()
        app.closeEvent(event)
        assert event.isAccepted()
        assert (tmp_path / "autosave.pmggrp").exists()
    finally:
        app.deleteLater()
        QApplication.processEvents()


def test_3d_close_with_unsaved_changes_cancel_keeps_window_open(
    sample_csv, monkeypatch
):
    app = GraphApp3D()
    try:
        app.load_data(file_path=sample_csv)
        app.data_table.item(0, 0).setText("edited")
        assert not app.undo_stack.isClean()

        monkeypatch.setattr(
            QMessageBox,
            "question",
            lambda *a, **k: QMessageBox.StandardButton.Cancel,
        )
        event = QCloseEvent()
        app.closeEvent(event)
        assert not event.isAccepted()
    finally:
        app.deleteLater()
        QApplication.processEvents()


# ── 2D → 3D Mode Switch Tests ────────────────────────────────────────────────
def test_open_in_3d_mode_carries_over_loaded_data(sample_csv):
    app = GraphApp2D()
    try:
        app.load_data(file_path=sample_csv)
        app.open_in_3d_mode()
        try:
            assert app.win_3d.df is not None
            assert list(app.win_3d.df.columns) == list(app.df.columns)
            assert app.win_3d.data_table.rowCount() == app.data_table.rowCount()
        finally:
            app.win_3d.deleteLater()
    finally:
        app.deleteLater()
        QApplication.processEvents()


# ── Reset Button Wiring Tests ────────────────────────────────────────────────
def test_3d_reset_btn_is_wired_to_reset_settings(app3d, sample_csv):
    app3d.elev_spin.setValue(99)
    app3d.reset_btn.click()
    assert app3d.elev_spin.value() == 30

    # app3d is a module-scoped fixture; reload data so later tests in this
    # module that expect it to already have data keep working.
    app3d.load_data(file_path=sample_csv)
    if app3d.z_listbox.count() >= 3:
        app3d.x_axis_combo.setCurrentIndex(0)
        app3d.y_axis_combo.setCurrentIndex(1)
        app3d.z_listbox.item(2).setSelected(True)


# ── Minor Tick Tests ─────────────────────────────────────────────────────────
def test_2d_minor_tick_interval_applies(app2d):
    import matplotlib.ticker as mticker

    app2d.plot_type_combo.setCurrentIndex(app2d.plot_type_combo.findText("line"))
    app2d.x_log_check.setChecked(False)
    app2d.xtick_minor_check.setChecked(True)
    app2d.xtick_minor_interval_input.setText("0.5")
    app2d.plot_graph()

    assert isinstance(app2d.ax.xaxis.get_minor_locator(), mticker.MultipleLocator)

    app2d.xtick_minor_check.setChecked(False)
    app2d.xtick_minor_interval_input.setText("")
    app2d.plot_graph()


# ── clear_all Tests ──────────────────────────────────────────────────────────
def test_2d_clear_all_clears_stale_column_selectors(app2d):
    assert app2d.x_tab_widgets[0]["x_combo"].count() > 0
    assert app2d.errorbar_column_combo.count() > 0

    app2d.clear_all()

    assert app2d.x_tab_widgets[0]["x_combo"].count() == 0
    assert app2d.x_tab_widgets[0]["y1_listbox"].count() == 0
    assert app2d.errorbar_column_combo.count() == 0
    assert app2d.filter_column_combo.count() == 0


def test_3d_clear_all_clears_stale_column_selectors(app3d):
    assert app3d.x_axis_combo.count() > 0

    app3d.clear_all()

    assert app3d.x_axis_combo.count() == 0
    assert app3d.y_axis_combo.count() == 0
    assert app3d.z_listbox.count() == 0


# ── Reset & Clear All Tests ──────────────────────────────────────────────────
def test_2d_reset_and_clear(app2d):
    app2d.title_input.setText("Some Title")
    app2d.ylabel2_input.setText("Y2 label")
    app2d.spine_top_check.setChecked(False)
    app2d.x_log_check.setChecked(True)
    app2d.colormap_combo.setCurrentIndex(app2d.colormap_combo.findText("jet"))
    app2d.y1_series_styles["(Y1) Val1"] = {"color": "#ff0000"}

    app2d.reset_settings()

    assert app2d.title_input.text() == ""
    assert app2d.ylabel2_input.text() == ""
    assert app2d.spine_top_check.isChecked() is True
    assert app2d.x_log_check.isChecked() is False
    assert app2d.colormap_combo.currentText() == "viridis"
    assert app2d.y1_series_styles == {}
    assert app2d.df is None


def test_3d_reset_and_clear(app3d):
    app3d.elev_spin.setValue(45)
    app3d.title_input.setText("Some Title")
    app3d.colormap_combo.setCurrentIndex(app3d.colormap_combo.findText("jet"))

    app3d.reset_settings()

    assert app3d.elev_spin.value() == 30
    assert app3d.title_input.text() == ""
    assert app3d.colormap_combo.currentText() == "viridis"
    assert app3d.df is None


def test_3d_show_about_does_not_crash(app3d):
    # QMessageBox.about is mocked globally (conftest.mock_qt_dialogs); this
    # just guards that the 3D window has an About action wired up at all.
    app3d.show_about()
