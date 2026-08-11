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


def test_data_manager_load_tsv(tmp_path):
    dm = DataManager()
    tsv_file = tmp_path / "test_data.tsv"
    tsv_file.write_text("Time\tTemp\tPressure\n1\t10\t100\n2\t20\t200\n")

    dm.load_file(str(tsv_file))
    assert dm.has_data()
    assert dm.get_columns() == ["Time", "Temp", "Pressure"]
    assert list(dm.raw_df["Temp"]) == ["10", "20"]


def test_data_manager_load_txt_comma_delimited(tmp_path):
    dm = DataManager()
    txt_file = tmp_path / "test_data.txt"
    txt_file.write_text("Time,Temp,Pressure\n1,10,100\n2,20,200\n")

    dm.load_file(str(txt_file))
    assert dm.has_data()
    assert dm.get_columns() == ["Time", "Temp", "Pressure"]
    assert list(dm.raw_df["Temp"]) == ["10", "20"]


def test_data_manager_load_txt_tab_delimited(tmp_path):
    dm = DataManager()
    txt_file = tmp_path / "test_data.txt"
    txt_file.write_text("Time\tTemp\tPressure\n1\t10\t100\n2\t20\t200\n")

    dm.load_file(str(txt_file))
    assert dm.has_data()
    assert dm.get_columns() == ["Time", "Temp", "Pressure"]
    assert list(dm.raw_df["Temp"]) == ["10", "20"]


def test_data_manager_load_json(tmp_path):
    dm = DataManager()
    json_file = tmp_path / "test_data.json"
    json_file.write_text('[{"Time": 1, "Temp": 10}, {"Time": 2, "Temp": 20}]')

    dm.load_file(str(json_file))
    assert dm.has_data()
    assert dm.get_columns() == ["Time", "Temp"]
    assert list(dm.raw_df["Temp"]) == ["10", "20"]


def test_data_manager_load_xlsx(tmp_path):
    dm = DataManager()
    xlsx_file = tmp_path / "test_data.xlsx"
    import pandas as pd

    pd.DataFrame({"Time": [1, 2], "Temp": [10, 20]}).to_excel(
        xlsx_file, index=False
    )

    dm.load_file(str(xlsx_file))
    assert dm.has_data()
    assert dm.get_columns() == ["Time", "Temp"]
    assert list(dm.raw_df["Temp"]) == ["10", "20"]


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
