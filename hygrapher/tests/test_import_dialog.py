# -*- coding: utf-8 -*-
import os

os.environ["QT_QPA_PLATFORM"] = "offscreen"

import pytest
from PyQt6.QtWidgets import QApplication

from hygrapher.import_dialog import ImportPreviewDialog


@pytest.fixture(scope="module", autouse=True)
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    yield app
    app.processEvents()


def test_dialog_defaults_to_header_row_zero():
    rows = [["A", "B"], ["1", "2"], ["3", "4"]]
    dialog = ImportPreviewDialog("data.csv", rows)
    assert dialog.header_row == 0
    assert dialog.header_spin.value() == 0
    assert dialog.header_spin.maximum() == 2
    dialog.deleteLater()


def test_dialog_table_shows_all_preview_rows_and_columns():
    rows = [["Title Only"], ["A", "B", "C"], ["1", "2", "3"]]
    dialog = ImportPreviewDialog("data.csv", rows)
    assert dialog.table.rowCount() == 3
    assert dialog.table.columnCount() == 3
    assert dialog.table.item(0, 0).text() == "Title Only"
    assert dialog.table.item(1, 1).text() == "B"
    dialog.deleteLater()


def test_dialog_changing_header_row_updates_state():
    rows = [["Title"], ["A", "B"], ["1", "2"]]
    dialog = ImportPreviewDialog("data.csv", rows)
    dialog.header_spin.setValue(1)
    assert dialog.header_row == 1
    dialog.deleteLater()


def test_dialog_no_header_checkbox_disables_spin_and_clears_header_row():
    rows = [["1", "2"], ["3", "4"]]
    dialog = ImportPreviewDialog("data.csv", rows)
    dialog.no_header_check.setChecked(True)
    assert dialog.header_row is None
    assert dialog.header_spin.isEnabled() is False

    dialog.no_header_check.setChecked(False)
    assert dialog.header_row == dialog.header_spin.value()
    assert dialog.header_spin.isEnabled() is True
    dialog.deleteLater()


def test_dialog_handles_empty_preview():
    dialog = ImportPreviewDialog("data.csv", [])
    assert dialog.table.rowCount() == 0
    assert dialog.header_spin.maximum() == 0
    dialog.deleteLater()
