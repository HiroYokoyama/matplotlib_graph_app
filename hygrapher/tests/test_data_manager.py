# -*- coding: utf-8 -*-
import pandas as pd

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

    pd.DataFrame({"Time": [1, 2], "Temp": [10, 20]}).to_excel(
        xlsx_file, index=False
    )

    dm.load_file(str(xlsx_file))
    assert dm.has_data()
    assert dm.get_columns() == ["Time", "Temp"]
    assert list(dm.raw_df["Temp"]) == ["10", "20"]


def test_data_manager_load_file_with_title_row_above_header(tmp_path):
    """A common messy-file case: a title line, then the real header, then data."""
    messy_file = tmp_path / "messy.csv"
    messy_file.write_text("Experiment Run 42\nTime,Val1,Val2\n1,10,100\n2,20,200\n")

    dm = DataManager()
    dm.load_file(str(messy_file), header_row=1)
    assert dm.get_columns() == ["Time", "Val1", "Val2"]
    assert len(dm.raw_df) == 2
    assert list(dm.raw_df["Val1"]) == ["10", "20"]


def test_data_manager_load_file_no_header_auto_names_columns(tmp_path):
    headerless_file = tmp_path / "headerless.csv"
    headerless_file.write_text("1,10,100\n2,20,200\n")

    dm = DataManager()
    dm.load_file(str(headerless_file), header_row=None)
    assert dm.get_columns() == ["Column1", "Column2", "Column3"]
    assert len(dm.raw_df) == 2


def test_data_manager_read_preview_rows_csv(tmp_path):
    csv_file = tmp_path / "data.csv"
    csv_file.write_text("A,B\n1,2\n3,4\n")

    dm = DataManager()
    rows = dm.read_preview_rows(str(csv_file))
    assert rows == [["A", "B"], ["1", "2"], ["3", "4"]]


def test_data_manager_read_preview_rows_ragged_title_row(tmp_path):
    """Preview must tolerate a title row with a different field count than
    the data rows below it — pandas' own tokenizer rejects this outright."""
    messy_file = tmp_path / "messy.csv"
    messy_file.write_text("Experiment Run 42\nTime,Val1,Val2\n1,10,100\n2,20,200\n")

    dm = DataManager()
    rows = dm.read_preview_rows(str(messy_file))
    assert rows[0] == ["Experiment Run 42"]
    assert rows[1] == ["Time", "Val1", "Val2"]
    assert rows[2] == ["1", "10", "100"]


def test_data_manager_read_preview_rows_respects_max_rows(tmp_path):
    csv_file = tmp_path / "data.csv"
    csv_file.write_text("A\n" + "\n".join(str(i) for i in range(50)))

    dm = DataManager()
    rows = dm.read_preview_rows(str(csv_file), max_rows=5)
    assert len(rows) == 5


def test_data_manager_read_preview_rows_xlsx(tmp_path):
    xlsx_file = tmp_path / "data.xlsx"
    pd.DataFrame({"A": [1, 2]}).to_excel(xlsx_file, index=False)

    dm = DataManager()
    rows = dm.read_preview_rows(str(xlsx_file))
    assert rows[0] == ["A"]
    assert rows[1] == ["1"]


def test_data_manager_clear():
    dm = DataManager()
    dm.set_dataframe(pd.DataFrame({"A": ["1", "2"], "B": ["10", "20"]}))
    assert dm.has_data()

    dm.clear()
    assert not dm.has_data()
    assert dm.get_columns() == []

    dm.clear()
    assert not dm.has_data()
