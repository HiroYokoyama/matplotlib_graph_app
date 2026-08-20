# -*- coding: utf-8 -*-
"""
Tests for the series-picker controls (Select All / Clear / counter), the
scrollable settings tabs, and double-click-to-reset-zoom.
"""

import os

os.environ["QT_QPA_PLATFORM"] = "offscreen"

import pytest
from unittest.mock import MagicMock

from PyQt6.QtWidgets import QApplication, QLabel, QPushButton, QScrollArea
import matplotlib

matplotlib.use("Agg")

from hygrapher.main import GraphApp as GraphApp2D
from hygrapher.main_3d import GraphApp3D
from hygrapher.widgets import build_series_picker, wrap_in_scroll_area


@pytest.fixture(scope="session", autouse=True)
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    yield app
    app.processEvents()


@pytest.fixture(scope="module")
def sample_csv(tmp_path_factory):
    p = tmp_path_factory.mktemp("data") / "picker.csv"
    p.write_text("Time,Val1,Val2,Err\n1,10,100,0.5\n2,25,200,1.0\n3,15,300,1.5\n")
    return str(p)


@pytest.fixture
def app2d(sample_csv):
    app = GraphApp2D()
    app.load_data(file_path=sample_csv)
    yield app
    app.close()
    app.deleteLater()
    QApplication.processEvents()


def _buttons(box):
    return {b.text(): b for b in box.findChildren(QPushButton)}


def _counter_text(box):
    return [lbl.text() for lbl in box.findChildren(QLabel) if "selected" in lbl.text()]


def test_select_all_and_clear_on_y1_and_y2(app2d):
    tab = app2d.x_tab_widgets[0]
    for side in ("y1", "y2"):
        box, listbox = tab[f"{side}_box"], tab[f"{side}_listbox"]
        assert listbox.count() == 4

        _buttons(box)["Select All"].click()
        assert len(listbox.selectedItems()) == 4

        _buttons(box)["Clear"].click()
        assert listbox.selectedItems() == []


def test_select_all_feeds_the_plot(app2d):
    tab = app2d.x_tab_widgets[0]
    _buttons(tab["y1_box"])["Select All"].click()
    app2d.plot_graph()
    plotted = {line.get_label() for line in app2d.ax.get_lines()}
    assert {"Val1", "Val2", "Err"} <= plotted


def test_selection_counter_tracks_selection(app2d):
    tab = app2d.x_tab_widgets[0]
    box, listbox = tab["y1_box"], tab["y1_listbox"]
    _buttons(box)["Clear"].click()
    assert _counter_text(box) == ["0 selected"]
    listbox.item(1).setSelected(True)
    assert _counter_text(box) == ["1 selected"]
    _buttons(box)["Select All"].click()
    assert _counter_text(box) == ["4 selected"]


def test_series_lists_are_tall_enough_to_be_usable(app2d):
    tab = app2d.x_tab_widgets[0]
    for side in ("y1", "y2"):
        listbox = tab[f"{side}_listbox"]
        row_height = listbox.fontMetrics().height() + 8
        assert listbox.minimumHeight() >= row_height * 8


def test_every_settings_tab_scrolls(app2d):
    notebook = app2d.settings_notebook
    assert notebook.count() == 7
    for i in range(notebook.count()):
        page = notebook.widget(i)
        assert isinstance(page, QScrollArea)
        assert page.widgetResizable()


def test_scroll_wrapper_keeps_content_at_natural_height(app2d):
    notebook = app2d.settings_notebook
    spines = notebook.widget([notebook.tabText(i) for i in range(7)].index("Spines"))
    inner = spines.widget()
    # trailing stretch absorbs the slack instead of spreading the checkboxes
    assert inner.layout().count() == 2
    assert inner.layout().itemAt(1).widget() is None


def test_double_click_resets_zoom(app2d):
    tab = app2d.x_tab_widgets[0]
    tab["y1_listbox"].item(1).setSelected(True)
    app2d.plot_graph()
    home_xlim = app2d.ax.get_xlim()

    app2d.ax.set_xlim(1.2, 1.4)
    assert app2d.ax.get_xlim() != home_xlim

    event = MagicMock()
    event.dblclick = True
    app2d.on_canvas_click(event)
    assert app2d.ax.get_xlim() == pytest.approx(home_xlim)


def test_double_click_restores_configured_limits(app2d):
    tab = app2d.x_tab_widgets[0]
    tab["y1_listbox"].item(1).setSelected(True)
    app2d.ylim_min_input.setText("0")
    app2d.ylim_max_input.setText("99")
    app2d.plot_graph()
    app2d.ax.set_ylim(3, 4)

    event = MagicMock()
    event.dblclick = True
    app2d.on_canvas_click(event)
    assert app2d.ax.get_ylim() == pytest.approx((0.0, 99.0))


def test_single_click_leaves_the_view_alone(app2d):
    app2d.plot_graph()
    app2d.ax.set_xlim(1.2, 1.4)
    event = MagicMock()
    event.dblclick = False
    app2d.on_canvas_click(event)
    assert app2d.ax.get_xlim() == pytest.approx((1.2, 1.4))


def test_3d_z_picker_has_select_all_and_double_click_reset(sample_csv):
    app = GraphApp3D()
    try:
        app.load_data(file_path=sample_csv)
        box = app.z_listbox.parent()
        _buttons(box)["Select All"].click()
        assert len(app.z_listbox.selectedItems()) == app.z_listbox.count() > 0

        app.elev_spin.setValue(30)
        app.azim_spin.setValue(-60)
        app.plot_graph()
        app.ax.view_init(elev=5, azim=5)

        event = MagicMock()
        event.dblclick = True
        app.on_canvas_click(event)
        assert round(app.ax.elev) == 30
        assert round(app.ax.azim) == -60
    finally:
        app.close()
        app.deleteLater()
        QApplication.processEvents()


def test_build_series_picker_without_title():
    box, listbox = build_series_picker("", visible_rows=3)
    assert listbox.minimumHeight() < (listbox.fontMetrics().height() + 8) * 4
    assert set(_buttons(box)) == {"Select All", "Clear"}


def test_wrap_in_scroll_area_reparents_the_widget():
    box, _ = build_series_picker("Anything")
    area = wrap_in_scroll_area(box)
    assert area.widget() is not box
    assert box.parent() is area.widget()
