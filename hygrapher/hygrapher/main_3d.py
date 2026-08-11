# -*- coding: utf-8 -*-
"""
HyGrapher 3D - 3D Matplotlib Plotting Desktop Application (PyQt6 Edition)
"""

import sys
import os
import pandas as pd
import numpy as np
from scipy.interpolate import griddata

from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QPushButton,
    QComboBox,
    QLineEdit,
    QSpinBox,
    QListWidget,
    QAbstractItemView,
    QFileDialog,
    QMessageBox,
    QFormLayout,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction, QDropEvent, QDragEnterEvent

import matplotlib

matplotlib.use("QtAgg")
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg, NavigationToolbar2QT
from matplotlib.figure import Figure

from hygrapher.data_manager import DataManager
from hygrapher.project_io import save_project_file, load_project_file

VERSION = "0.6.0"


class GraphApp3D(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Matplotlib 3D Graph App v{VERSION} (PyQt6)")
        self.resize(1200, 800)
        self.setAcceptDrops(True)

        self.data_mgr = DataManager()
        self.df = None
        self.current_project_path = None

        self.init_ui()

    def init_ui(self):
        # Menu Bar
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("File")

        open_action = QAction("Open Data File...", self)
        open_action.triggered.connect(self.load_data)
        file_menu.addAction(open_action)

        save_action = QAction("Save 3D Project", self)
        save_action.triggered.connect(self.overwrite_save)
        file_menu.addAction(save_action)

        export_action = QAction("Export Plot Image...", self)
        export_action.triggered.connect(self.export_graph)
        file_menu.addAction(export_action)

        # Central Layout with Splitter
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # Left Panel (Controls & Sheet)
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Buttons
        btn_layout = QHBoxLayout()
        self.open_btn = QPushButton("Open File")
        self.open_btn.clicked.connect(self.load_data)
        btn_layout.addWidget(self.open_btn)

        self.plot_btn = QPushButton("Plot 3D")
        self.plot_btn.setStyleSheet(
            "font-weight: bold; background-color: #2b5c8f; color: white;"
        )
        self.plot_btn.clicked.connect(self.plot_graph)
        btn_layout.addWidget(self.plot_btn)

        left_layout.addLayout(btn_layout)

        # Form Controls
        form = QFormLayout()

        self.plot_type_combo = QComboBox()
        self.plot_type_combo.addItems(
            ["surface", "wireframe", "contour3d", "scatter3d", "line3d"]
        )
        form.addRow("Plot Type:", self.plot_type_combo)

        self.x_axis_combo = QComboBox()
        form.addRow("X Axis:", self.x_axis_combo)

        self.y_axis_combo = QComboBox()
        form.addRow("Y Axis:", self.y_axis_combo)

        self.z_listbox = QListWidget()
        self.z_listbox.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        form.addRow("Z Column(s):", self.z_listbox)

        self.title_input = QLineEdit()
        form.addRow("Title:", self.title_input)

        self.elev_spin = QSpinBox()
        self.elev_spin.setRange(-180, 180)
        self.elev_spin.setValue(30)
        form.addRow("Elevation (elev):", self.elev_spin)

        self.azim_spin = QSpinBox()
        self.azim_spin.setRange(-180, 180)
        self.azim_spin.setValue(-60)
        form.addRow("Azimuth (azim):", self.azim_spin)

        self.resolution_spin = QSpinBox()
        self.resolution_spin.setRange(5, 200)
        self.resolution_spin.setValue(30)
        form.addRow("Mesh Resolution:", self.resolution_spin)

        self.colormap_combo = QComboBox()
        self.colormap_combo.addItems(
            [
                "viridis",
                "plasma",
                "inferno",
                "magma",
                "cividis",
                "coolwarm",
                "jet",
                "rainbow",
            ]
        )
        form.addRow("Colormap:", self.colormap_combo)

        left_layout.addLayout(form)

        # Data Sheet Table
        self.data_table = QTableWidget()
        left_layout.addWidget(self.data_table)

        splitter.addWidget(left_panel)

        # Right Panel (3D Canvas)
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)

        self.fig = Figure(figsize=(7, 5), dpi=100)
        self.ax = self.fig.add_subplot(111, projection="3d")

        self.canvas = FigureCanvasQTAgg(self.fig)
        self.toolbar = NavigationToolbar2QT(self.canvas, self)

        right_layout.addWidget(self.toolbar)
        right_layout.addWidget(self.canvas)

        splitter.addWidget(right_panel)
        splitter.setSizes([450, 750])

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".pmggrp":
                self.load_project_file(file_path)
            elif ext in [".csv", ".tsv", ".xlsx", ".xls", ".txt", ".json"]:
                self.load_data(file_path=file_path)

    def load_data(self, file_path=None):
        if not file_path:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Open Data File",
                "",
                "Supported Files (*.csv *.tsv *.xlsx *.xls *.txt *.json)",
            )
        if not file_path:
            return

        try:
            self.df = self.data_mgr.load_file(file_path)
            self.populate_data_table()
            self.update_combos()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load file:\n{e}")

    def populate_data_table(self):
        if self.df is None:
            return
        self.data_table.setRowCount(len(self.df))
        self.data_table.setColumnCount(len(self.df.columns))
        self.data_table.setHorizontalHeaderLabels(list(self.df.columns))
        for r in range(len(self.df)):
            for c in range(len(self.df.columns)):
                self.data_table.setItem(r, c, QTableWidgetItem(str(self.df.iat[r, c])))

    def update_combos(self):
        if self.df is None:
            return
        cols = self.data_mgr.get_columns()
        self.x_axis_combo.clear()
        self.x_axis_combo.addItems(cols)
        self.y_axis_combo.clear()
        self.y_axis_combo.addItems(cols)
        self.z_listbox.clear()
        for col in cols:
            self.z_listbox.addItem(col)
        if len(cols) >= 3:
            self.x_axis_combo.setCurrentIndex(0)
            self.y_axis_combo.setCurrentIndex(1)
            self.z_listbox.item(2).setSelected(True)

    def get_data_from_table(self):
        if self.df is None or self.data_table.rowCount() == 0:
            return
        cols = [
            self.data_table.horizontalHeaderItem(c).text()
            for c in range(self.data_table.columnCount())
        ]
        data = [
            [
                self.data_table.item(r, c).text() if self.data_table.item(r, c) else ""
                for c in range(self.data_table.columnCount())
            ]
            for r in range(self.data_table.rowCount())
        ]
        self.data_mgr.set_dataframe(pd.DataFrame(data, columns=cols))
        self.df = self.data_mgr.get_filtered_df()

    def plot_graph(self):
        try:
            self.fig.clear()
            self.ax = self.fig.add_subplot(111, projection="3d")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to clear 3D graph:\n{e}")
            return

        self.get_data_from_table()
        if not self.data_mgr.has_data():
            QMessageBox.information(self, "Info", "No data to plot.")
            return

        x_col = self.x_axis_combo.currentText()
        y_col = self.y_axis_combo.currentText()
        z_cols = [item.text() for item in self.z_listbox.selectedItems()]

        if not x_col or not y_col or not z_cols:
            QMessageBox.warning(
                self, "Warning", "Select X, Y, and at least one Z column."
            )
            return

        try:
            self.ax.view_init(elev=self.elev_spin.value(), azim=self.azim_spin.value())

            x_num = pd.to_numeric(
                self.df[x_col].astype(str).str.replace(r"[^\d.-]", "", regex=True),
                errors="coerce",
            )
            y_num = pd.to_numeric(
                self.df[y_col].astype(str).str.replace(r"[^\d.-]", "", regex=True),
                errors="coerce",
            )
            plot_type = self.plot_type_combo.currentText()

            for z_c in z_cols:
                z_num = pd.to_numeric(
                    self.df[z_c].astype(str).str.replace(r"[^\d.-]", "", regex=True),
                    errors="coerce",
                )
                valid_df = pd.DataFrame({"x": x_num, "y": y_num, "z": z_num}).dropna()
                if valid_df.empty:
                    continue

                if plot_type in ["surface", "wireframe", "contour3d"]:
                    res = self.resolution_spin.value()
                    xi = np.linspace(valid_df["x"].min(), valid_df["x"].max(), res)
                    yi = np.linspace(valid_df["y"].min(), valid_df["y"].max(), res)
                    Xi, Yi = np.meshgrid(xi, yi)
                    try:
                        Zi = griddata(
                            (valid_df["x"].values, valid_df["y"].values),
                            valid_df["z"].values,
                            (Xi, Yi),
                            method="linear",
                        )
                    except Exception:
                        Zi = griddata(
                            (valid_df["x"].values, valid_df["y"].values),
                            valid_df["z"].values,
                            (Xi, Yi),
                            method="nearest",
                        )
                    Zi = np.nan_to_num(Zi, nan=np.nanmean(valid_df["z"].values))

                    if plot_type == "surface":
                        self.ax.plot_surface(
                            Xi,
                            Yi,
                            Zi,
                            cmap=self.colormap_combo.currentText(),
                            alpha=0.8,
                        )
                    elif plot_type == "wireframe":
                        self.ax.plot_wireframe(Xi, Yi, Zi, rstride=5, cstride=5)
                    elif plot_type == "contour3d":
                        if np.nanmin(Zi) != np.nanmax(Zi):
                            self.ax.contour3D(
                                Xi, Yi, Zi, 20, cmap=self.colormap_combo.currentText()
                            )

                elif plot_type == "scatter3d":
                    self.ax.scatter(
                        valid_df["x"], valid_df["y"], valid_df["z"], label=z_c
                    )
                elif plot_type == "line3d":
                    self.ax.plot(valid_df["x"], valid_df["y"], valid_df["z"], label=z_c)

            self.ax.set_xlabel(x_col)
            self.ax.set_ylabel(y_col)
            self.ax.set_zlabel(", ".join(z_cols))
            if self.title_input.text():
                self.ax.set_title(self.title_input.text())

            self.canvas.draw()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to plot 3D:\n{e}")

    def save_settings(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save 3D Project", "", "Matplotlib Graph Project (*.pmggrp)"
        )
        if file_path:
            save_project_file(self, file_path, version_str=VERSION, dimension="3D")
            self.current_project_path = file_path

    def overwrite_save(self):
        if not self.current_project_path:
            self.save_settings()
        else:
            save_project_file(
                self, self.current_project_path, version_str=VERSION, dimension="3D"
            )

    def load_project_file(self, file_path):
        load_project_file(self, file_path)

    def export_graph(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export 3D Plot Image", "", "PNG Image (*.png);;PDF Document (*.pdf)"
        )
        if file_path:
            self.fig.savefig(file_path, dpi=300, bbox_inches="tight")
            QMessageBox.information(self, "Success", f"Plot saved to {file_path}")

    def clear_all(self):
        self.data_mgr.clear()
        self.df = None
        self.data_table.clear()
        self.data_table.setRowCount(0)
        self.data_table.setColumnCount(0)
        self.fig.clear()
        self.ax = self.fig.add_subplot(111, projection="3d")
        self.canvas.draw()

    def reset_settings(self):
        self.plot_type_combo.setCurrentIndex(0)
        self.elev_spin.setValue(30)
        self.azim_spin.setValue(-60)
        self.clear_all()


def main():
    app = QApplication(sys.argv)
    window = GraphApp3D()
    window.show()
    sys.exit(app.exec())


GraphApp = GraphApp3D

if __name__ == "__main__":
    main()
