# -*- coding: utf-8 -*-
import json
import pandas as pd
from unittest.mock import MagicMock
from hygrapher.project_io import build_project_dict, save_project_file


def test_build_project_dict():
    mock_app = MagicMock()
    mock_app.df = pd.DataFrame({"X": [1, 2], "Y": [10, 20]})
    mock_app.get_data_from_sheet = MagicMock()
    mock_app.get_x_tabs_data = MagicMock(
        return_value=[{"tab_name": "X-Tab 1", "x_axis": "X"}]
    )

    # Setup mock variables
    mock_app.plot_type_var.get.return_value = "line"
    mock_app.x_axis_var.get.return_value = "X"
    mock_app.title_var.get.return_value = "Test Title"
    mock_app.xlabel_var.get.return_value = "X Label"
    mock_app.ylabel_var.get.return_value = "Y Label"

    project_dict = build_project_dict(mock_app, version_str="0.6.0", dimension="2D")
    assert project_dict["application"] == "HYGrapher"
    assert project_dict["application_version"] == "0.6.0"
    assert project_dict["dimension"] == "2D"
    assert project_dict["plot_type"] == "line"
    assert project_dict["edited_data"]["columns"] == ["X", "Y"]
    assert len(project_dict["x_tabs"]) == 1


def test_save_project_file(tmp_path):
    mock_app = MagicMock()
    mock_app.data_file_path = ""
    mock_app.df = pd.DataFrame({"A": [1], "B": [2]})
    mock_app.get_data_from_sheet = MagicMock()
    mock_app.get_x_tabs_data = MagicMock(return_value=[])

    target_path = tmp_path / "project.pmggrp"
    save_project_file(mock_app, str(target_path), version_str="0.6.0", dimension="2D")

    assert target_path.exists()
    content = json.loads(target_path.read_text(encoding="utf-8"))
    assert content["dimension"] == "2D"
    assert content["edited_data"]["columns"] == ["A", "B"]
