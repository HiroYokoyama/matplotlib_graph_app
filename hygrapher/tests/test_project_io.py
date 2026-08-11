# -*- coding: utf-8 -*-
import json
import pandas as pd
from unittest.mock import MagicMock

from hygrapher.project_io import build_project_dict, save_project_file, load_project_file


class FakeLineEdit:
    def __init__(self, text=""):
        self._text = text

    def text(self):
        return self._text

    def setText(self, value):
        self._text = value


class FakeComboBox:
    def __init__(self, text=""):
        self._text = text

    def currentText(self):
        return self._text

    def findText(self, value):
        return 0 if value == self._text else -1

    def setCurrentIndex(self, idx):
        pass

    def setCurrentText(self, value):
        self._text = value


def _bare_app():
    """An app stand-in with none of the real widget attributes wired up.

    build_project_dict/load_project_file must degrade gracefully (defaults,
    no exceptions) rather than crash when fields are missing.
    """
    app = MagicMock(spec=["df", "get_x_tabs_data"])
    app.get_x_tabs_data = MagicMock(return_value=[{"x_axis": "X", "y1_cols": ["Y"], "y2_cols": []}])
    return app


def test_build_project_dict_defaults_when_widgets_missing():
    mock_app = _bare_app()
    mock_app.df = pd.DataFrame({"X": [1, 2], "Y": [10, 20]})

    project_dict = build_project_dict(mock_app, version_str="0.6.0", dimension="2D")
    assert project_dict["application"] == "HYGrapher"
    assert project_dict["application_version"] == "0.6.0"
    assert project_dict["dimension"] == "2D"
    assert project_dict["plot_type"] == "line"  # default, no plot_type_combo present
    assert project_dict["edited_data"]["columns"] == ["X", "Y"]
    assert len(project_dict["x_tabs"]) == 1


def test_build_project_dict_reads_real_widgets():
    mock_app = _bare_app()
    mock_app.df = pd.DataFrame({"X": [1, 2], "Y": [10, 20]})
    mock_app.title_input = FakeLineEdit("Test Title")
    mock_app.plot_type_combo = FakeComboBox("scatter")

    project_dict = build_project_dict(mock_app, version_str="0.6.0", dimension="2D")
    assert project_dict["title"] == "Test Title"
    assert project_dict["plot_type"] == "scatter"


def test_save_project_file(tmp_path):
    mock_app = _bare_app()
    mock_app.df = pd.DataFrame({"A": [1], "B": [2]})
    mock_app.get_x_tabs_data = MagicMock(return_value=[])

    target_path = tmp_path / "project.pmggrp"
    save_project_file(mock_app, str(target_path), version_str="0.6.0", dimension="2D")

    assert target_path.exists()
    content = json.loads(target_path.read_text(encoding="utf-8"))
    assert content["dimension"] == "2D"
    assert content["edited_data"]["columns"] == ["A", "B"]


def test_save_load_round_trip_preserves_settings(tmp_path):
    mock_app = _bare_app()
    mock_app.df = pd.DataFrame({"A": [1], "B": [2]})
    mock_app.get_x_tabs_data = MagicMock(return_value=[])
    mock_app.title_input = FakeLineEdit("My Plot")
    mock_app.plot_type_combo = FakeComboBox("bar")

    target_path = tmp_path / "project.pmggrp"
    save_project_file(mock_app, str(target_path), version_str="0.6.0", dimension="2D")

    restored_app = MagicMock()
    restored_app.df = None
    restored_app.title_input = FakeLineEdit("")
    restored_app.plot_type_combo = FakeComboBox("")

    load_project_file(restored_app, str(target_path))
    assert restored_app.title_input.text() == "My Plot"
    assert restored_app.plot_type_combo.currentText() == "bar"
