# -*- coding: utf-8 -*-
import pytest
import os
import json
import pandas as pd
import numpy as np
from unittest.mock import MagicMock

from hygrapher.utils import (
    bind_scroll_events,
    bind_mousewheel_recursive,
    get_font_list,
    apply_major_ticker,
    apply_minor_ticker,
)
from hygrapher.data_manager import DataManager
from hygrapher.project_io import build_project_dict, save_project_file, _safe_get


def test_utils_functions():
    fonts = get_font_list()
    assert isinstance(fonts, list)
    assert len(fonts) > 0

    mock_axis = MagicMock()
    apply_major_ticker(mock_axis, "2.5", False)
    apply_major_ticker(mock_axis, "invalid", False)
    apply_major_ticker(mock_axis, "5.0", True)

    apply_minor_ticker(mock_axis, True, "0.5", False)
    apply_minor_ticker(mock_axis, True, "", False)
    apply_minor_ticker(mock_axis, False, "0.5", False)
    apply_minor_ticker(mock_axis, True, "", True)


def test_data_manager_complete(tmp_path):
    dm = DataManager()
    assert dm.raw_df is None
    assert not dm.has_data()
    assert dm.get_columns() == []

    # Test CSV loading
    csv_p = tmp_path / "data.csv"
    csv_p.write_text("X,Y,Z\n1,10,100\n2,20,200\n3,30,300\n4,40,400\n")
    dm.load_file(str(csv_p))

    assert dm.has_data()
    assert dm.get_columns() == ["X", "Y", "Z"]

    # Test filtering min & max
    f1 = dm.get_filtered_df(filter_enabled=True, filter_column="Y", min_val_str="20", max_val_str="30")
    assert len(f1) == 2
    assert list(f1["Y"]) == ["20", "30"]

    # Test filter disabled
    f2 = dm.get_filtered_df(filter_enabled=False, filter_column="Y", min_val_str="20")
    assert len(f2) == 4

    # Test filter error fallback
    f3 = dm.get_filtered_df(filter_enabled=True, filter_column="Y", min_val_str="invalid_number")
    assert len(f3) == 4

    # Test updating from sheet
    dm.update_from_sheet_data([["10", "100", "1000"]], ["X", "Y", "Z"])
    assert len(dm.raw_df) == 1
    assert list(dm.raw_df["X"]) == ["10"]

    dm.clear()
    assert not dm.has_data()


def test_project_io_helpers(tmp_path):
    assert _safe_get(None, "def") == "def"

    class DummyVar:
        def get(self):
            return "val"

    assert _safe_get(DummyVar(), "def") == "val"

    class ErrVar:
        def get(self):
            raise ValueError("err")

    assert _safe_get(ErrVar(), "fallback") == "fallback"

    class MockApp:
        def __init__(self):
            self.df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
            self.data_file_path = "/path/data.csv"

    app = MockApp()
    p_dict = build_project_dict(app, version_str="0.6.0", dimension="2D")
    assert p_dict["application"] == "HYGrapher"
    assert p_dict["original_file_path"] == "/path/data.csv"
    assert p_dict["edited_data"]["columns"] == ["A", "B"]

    p_dict_3d = build_project_dict(app, version_str="0.6.0", dimension="3D")
    assert p_dict_3d["dimension"] == "3D"
    assert "view_elev" in p_dict_3d

    file_p = tmp_path / "test.pmggrp"
    save_project_file(app, str(file_p), "0.6.0", "2D")
    assert file_p.exists()
    saved = json.loads(file_p.read_text(encoding="utf-8"))
    assert saved["dimension"] == "2D"
