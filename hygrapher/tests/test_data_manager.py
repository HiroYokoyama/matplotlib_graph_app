# -*- coding: utf-8 -*-
from hygrapher.data_manager import DataManager


def test_data_manager_load_and_filter(tmp_path):
    dm = DataManager()
    assert not dm.has_data()

    # Create dummy CSV
    csv_file = tmp_path / "test_data.csv"
    csv_file.write_text("Time,Temp,Pressure\n1,10,100\n2,20,200\n3,30,300\n4,40,400\n")

    dm.load_file(str(csv_file))
    assert dm.has_data()
    assert dm.get_columns() == ["Time", "Temp", "Pressure"]
    assert len(dm.raw_df) == 4

    # Test non-destructive filtering
    filtered_df = dm.get_filtered_df(
        filter_enabled=True, filter_column="Temp", min_val_str="20", max_val_str="30"
    )
    assert len(filtered_df) == 2
    assert list(filtered_df["Temp"]) == ["20", "30"]

    # Verify raw_df was not mutated
    assert len(dm.raw_df) == 4


def test_data_manager_sheet_update():
    dm = DataManager()
    headers = ["A", "B"]
    data = [["1", "10"], ["2", "20"]]

    dm.update_from_sheet_data(data, headers)
    assert dm.has_data()
    assert dm.get_columns() == ["A", "B"]
    assert len(dm.raw_df) == 2

    dm.clear()
    assert not dm.has_data()
