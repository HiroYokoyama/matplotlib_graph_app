# -*- coding: utf-8 -*-
"""
hygrapher.data_manager

Data handling module for HYGrapher.
Manages raw and filtered dataframes non-destructively, handles CSV/Excel reading,
and provides clean interfaces for sheet widgets and filtering.
"""

import pandas as pd
import numpy as np


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

    def load_file(self, file_path):
        """
        Load data from a CSV or Excel file into a string DataFrame.
        """
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, dtype=str)
        else:
            df = pd.read_excel(file_path, dtype=str)

        df.fillna("", inplace=True)
        self._raw_df = df
        self.file_path = file_path
        return self._raw_df

    def get_columns(self):
        if self.has_data():
            return self._raw_df.columns.tolist()
        return []

    def get_filtered_df(self, filter_enabled=False, filter_column="", min_val_str="", max_val_str=""):
        """
        Return a copy of the dataframe, optionally filtered by range, without mutating `_raw_df`.
        """
        if not self.has_data():
            return None

        df_copy = self._raw_df.copy()
        if not filter_enabled or not filter_column or filter_column not in df_copy.columns:
            return df_copy

        try:
            raw_series = df_copy[filter_column].fillna("").astype(str)
            clean_series = raw_series.str.replace(r'[^\d.-]', '', regex=True)
            filter_series = pd.to_numeric(clean_series, errors='coerce')
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

    def update_from_sheet_data(self, data, headers):
        """
        Update `_raw_df` from tksheet data rows and headers.
        """
        if not data or not headers:
            return None

        header_len = len(headers)
        cleaned_data = [row[:header_len] if len(row) >= header_len else row + [""] * (header_len - len(row)) for row in data]

        temp_df = pd.DataFrame(cleaned_data, columns=headers).astype(str)
        temp_df.replace(r'^\s*$', pd.NA, regex=True, inplace=True)
        temp_df.dropna(how='all', inplace=True)
        self._raw_df = temp_df.fillna("")
        return self._raw_df

    def clear(self):
        self._raw_df = None
        self.file_path = ""
