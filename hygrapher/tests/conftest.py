# -*- coding: utf-8 -*-
import os

os.environ["QT_QPA_PLATFORM"] = "offscreen"

import matplotlib

matplotlib.use("Agg")

import pytest
from unittest.mock import patch


@pytest.fixture(autouse=True)
def mock_qt_dialogs():
    with (
        patch("PyQt6.QtWidgets.QMessageBox.warning"),
        patch("PyQt6.QtWidgets.QMessageBox.critical"),
        patch("PyQt6.QtWidgets.QMessageBox.information"),
        patch("PyQt6.QtWidgets.QMessageBox.about"),
        patch("PyQt6.QtWidgets.QMessageBox.question"),
    ):
        yield
