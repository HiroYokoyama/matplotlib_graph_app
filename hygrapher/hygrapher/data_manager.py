# -*- coding: utf-8 -*-
"""
hygrapher.data_manager

Data handling module for HYGrapher.
Manages raw and filtered dataframes non-destructively, handles CSV/Excel reading,
and provides clean interfaces for sheet widgets and filtering.
"""

import csv
import os

import pandas as pd


class DataManager:
    """
    Manages application datasets non-destructively.
    `_raw_df` preserves original loaded data.
    `df` returns the active (optionally filtered) dataframe.
    """

    def __init__(self):
        self._raw_df = None
        self.file_path = ""

    @property
    def raw_df(self):
        return self._raw_df

    @raw_df.setter
    def raw_df(self, df):
        self._raw_df = df

    def has_data(self):
        return self._raw_df is not None and not self._raw_df.empty

    def set_dataframe(self, df):
        """
        Directly update the raw dataframe.
        """
        self._raw_df = df
        return self._raw_df

    def load_file(self, file_path, header_row=0):
        """
        Load data from a CSV, TSV, plain-text, JSON, or Excel file into a
        string DataFrame. `.txt` (and any unrecognized text extension) has
        its delimiter auto-detected, so comma-, tab-, or whitespace-
        separated files all work.

        `header_row` is the 0-based row (within the file) to treat as the
        column header; any rows above it (e.g. a title/metadata line) are
        skipped. Pass ``None`` for files with no header row at all — columns
        are then auto-named "Column1", "Column2", etc. Ignored for `.json`,
        whose keys are always the header.
        """
        ext = os.path.splitext(file_path)[1].lower()

        if ext == ".csv":
            df = pd.read_csv(file_path, dtype=str, header=header_row)
        elif ext == ".tsv":
            df = pd.read_csv(file_path, dtype=str, sep="\t", header=header_row)
        elif ext in (".xlsx", ".xls"):
            df = pd.read_excel(file_path, dtype=str, header=header_row)
        elif ext == ".json":
            df = pd.read_json(file_path)
            df = df.astype(str)
        else:
            # .txt and any other plain-text format: auto-detect the delimiter.
            # (pandas' own sep=None sniffing only looks at the very first
            # line, so it gets confused when header_row skips rows above it.)
            delimiter = self._detect_text_delimiter(file_path)
            df = pd.read_csv(file_path, dtype=str, sep=delimiter, header=header_row)

        if header_row is None and ext != ".json":
            df.columns = [f"Column{i + 1}" for i in range(len(df.columns))]

        df = df.fillna("")
        self._raw_df = df
        self.file_path = file_path
        return self._raw_df

    @staticmethod
    def _detect_text_delimiter(file_path):
        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
            sample = f.read(4096)
        try:
            return csv.Sniffer().sniff(sample, delimiters=",\t; ").delimiter
        except csv.Error:
            return ","

    def read_preview_rows(self, file_path, max_rows=20):
        """
        Return up to `max_rows` raw rows (list of lists of str) from a file
        with no header assumption, for an import-preview UI to let the user
        pick which row is the real header. Uses the stdlib `csv` reader
        (not pandas) because a title/metadata line above the real header
        commonly has a different field count than the data rows, which
        pandas' tokenizer rejects outright.
        """
        ext = os.path.splitext(file_path)[1].lower()

        if ext in (".xlsx", ".xls"):
            raw = pd.read_excel(file_path, header=None, dtype=str, nrows=max_rows)
            raw = raw.fillna("")
            return raw.astype(str).values.tolist()

        if ext == ".tsv":
            delimiter = "\t"
        elif ext == ".csv":
            delimiter = ","
        else:
            delimiter = self._detect_text_delimiter(file_path)

        rows = []
        with open(file_path, "r", encoding="utf-8", errors="replace", newline="") as f:
            for i, row in enumerate(csv.reader(f, delimiter=delimiter)):
                if i >= max_rows:
                    break
                rows.append(row)
        return rows

    def get_columns(self):
        if self.has_data():
            return self._raw_df.columns.tolist()
        return []

    def get_filtered_df(
        self, filter_enabled=False, filter_column="", min_val_str="", max_val_str=""
    ):
        """
        Return a copy of the dataframe, optionally filtered by range, without mutating `_raw_df`.
        """
        if not self.has_data():
            return None

        df_copy = self._raw_df.copy()
        if (
            not filter_enabled
            or not filter_column
            or filter_column not in df_copy.columns
        ):
            return df_copy

        try:
            raw_series = df_copy[filter_column].fillna("").astype(str)
            clean_series = raw_series.str.replace(r"[^\d.-]", "", regex=True)
            filter_series = pd.to_numeric(clean_series, errors="coerce")
            mask = pd.Series([True] * len(df_copy))

            if min_val_str:
                try:
                    min_val = float(min_val_str)
                    mask &= (filter_series >= min_val).fillna(False)
                except ValueError:
                    pass

            if max_val_str:
                try:
                    max_val = float(max_val_str)
                    mask &= (filter_series <= max_val).fillna(False)
                except ValueError:
                    pass

            filtered = df_copy[mask].reset_index(drop=True)
            return filtered
        except Exception as e:
            print(f"Data filter error: {e}")
            return df_copy

    def clear(self):
        self._raw_df = None
        self.file_path = ""
