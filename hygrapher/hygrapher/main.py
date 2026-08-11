# -*- coding: utf-8 -*-
"""
HyGrapher - Matplotlib Plotting Desktop Application (PyQt6 Edition)
"""

import sys
import os
import pathlib
import pandas as pd
import numpy as np
from scipy.interpolate import griddata

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QSplitter, QTableWidget, QTableWidgetItem, QLabel,
    QPushButton, QComboBox, QLineEdit, QCheckBox, QSpinBox, QDoubleSpinBox,
    QGroupBox, QListWidget, QAbstractItemView, QFileDialog, QMessageBox,
    QColorDialog, QScrollArea, QFormLayout, QHeaderView, QMenu, QMenuBar
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QAction, QColor, QFont, QDropEvent, QDragEnterEvent

import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg, NavigationToolbar2QT
from matplotlib.figure import Figure

from hygrapher.data_manager import DataManager
from hygrapher.project_io import save_project_file, load_project_file
from hygrapher.utils import (
    apply_major_ticker, apply_minor_ticker, get_font_list
)

VERSION = "0.6.0"


class GraphApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Matplotlib Graph App v{VERSION} (PyQt6)")
        self.resize(1300, 850)
        self.setAcceptDrops(True)

        self.data_mgr = DataManager()
        self.df = None
        self.current_project_path = None

        self.x_tab_widgets = []
        self.y1_series_styles = {}
        self.y2_series_styles = {}

        self.init_ui()

    def init_ui(self):
        # Menu Bar
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("File")

        open_action = QAction("Open Data File...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.load_data)
        file_menu.addAction(open_action)

        save_action = QAction("Save Project", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.overwrite_save)
        file_menu.addAction(save_action)

        save_as_action = QAction("Save Project As...", self)
        save_as_action.triggered.connect(self.save_settings)
        file_menu.addAction(save_as_action)

        load_proj_action = QAction("Load Project...", self)
        load_proj_action.setShortcut("Ctrl+L")
        load_proj_action.triggered.connect(self.load_settings)
        file_menu.addAction(load_proj_action)

        file_menu.addSeparator()
        export_graph_action = QAction("Export Plot Image...", self)
        export_graph_action.triggered.connect(self.export_graph)
        file_menu.addAction(export_graph_action)

        export_data_action = QAction("Export Filtered Data...", self)
        export_data_action.triggered.connect(self.export_filtered_data)
        file_menu.addAction(export_data_action)

        mode_menu = menu_bar.addMenu("Mode")
        mode_3d_action = QAction("Switch to 3D Plotter Mode", self)
        mode_3d_action.triggered.connect(self.open_in_3d_mode)
        mode_menu.addAction(mode_3d_action)

        help_menu = menu_bar.addMenu("Help")
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        # Central Layout with Splitter
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # Left Panel (Settings & Data)
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Top Control Buttons
        top_btn_layout = QHBoxLayout()
        self.open_file_btn = QPushButton("Open File")
        self.open_file_btn.clicked.connect(self.load_data)
        top_btn_layout.addWidget(self.open_file_btn)

        self.plot_button = QPushButton("Plot Graph")
        self.plot_button.setStyleSheet("font-weight: bold; background-color: #2b5c8f; color: white;")
        self.plot_button.clicked.connect(self.plot_graph)
        top_btn_layout.addWidget(self.plot_button)

        self.export_button = QPushButton("Export Plot")
        self.export_button.clicked.connect(self.export_graph)
        top_btn_layout.addWidget(self.export_button)

        self.reset_btn = QPushButton("Reset All")
        self.reset_btn.clicked.connect(self.reset_settings)
        top_btn_layout.addWidget(self.reset_btn)

        left_layout.addLayout(top_btn_layout)

        # Settings Tabs
        self.settings_notebook = QTabWidget()
        left_layout.addWidget(self.settings_notebook, stretch=3)

        self.create_basic_settings_tab()
        self.create_style_settings_tab()
        self.create_axis_ticks_tab()
        self.create_font_size_tab()
        self.create_spines_tab()
        self.create_legend_tab()
        self.create_advanced_tab()

        # Data Table View
        self.data_group = QGroupBox("Data Sheet View")
        data_layout = QVBoxLayout(self.data_group)
        self.data_table = QTableWidget()
        data_layout.addWidget(self.data_table)
        left_layout.addWidget(self.data_group, stretch=2)

        splitter.addWidget(left_panel)

        # Right Panel (Plot Area)
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)

        self.fig = Figure(figsize=(7, 5), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax2 = None

        self.canvas = FigureCanvasQTAgg(self.fig)
        self.toolbar = NavigationToolbar2QT(self.canvas, self)

        right_layout.addWidget(self.toolbar)
        
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(self.canvas)
        right_layout.addWidget(scroll_area)

        splitter.addWidget(right_panel)
        splitter.setSizes([550, 750])

    def create_basic_settings_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Plot type
        pt_layout = QHBoxLayout()
        pt_layout.addWidget(QLabel("Plot Type:"))
        self.plot_type_combo = QComboBox()
        self.plot_type_combo.addItems([
            "line", "scatter", "bar", "step", "stem", "area",
            "pie", "box", "violin", "heatmap", "contour", "polar"
        ])
        pt_layout.addWidget(self.plot_type_combo)
        layout.addLayout(pt_layout)

        # X-Tabs Notebook
        x_tabs_header = QHBoxLayout()
        x_tabs_header.addWidget(QLabel("Multi-X Axis Tabs:"))
        self.add_x_tab_btn = QPushButton("+ Add X-Tab")
        self.add_x_tab_btn.clicked.connect(lambda: self.add_x_tab())
        x_tabs_header.addWidget(self.add_x_tab_btn)
        layout.addLayout(x_tabs_header)

        self.x_notebook = QTabWidget()
        layout.addWidget(self.x_notebook)
        self.add_x_tab(is_initial=True)

        # Titles & Labels
        form = QFormLayout()
        self.title_input = QLineEdit()
        self.xlabel_input = QLineEdit()
        self.ylabel_input = QLineEdit()
        self.ylabel2_input = QLineEdit()

        form.addRow("Title:", self.title_input)
        form.addRow("X Label:", self.xlabel_input)
        form.addRow("Y1 Label:", self.ylabel_input)
        form.addRow("Y2 Label:", self.ylabel2_input)
        layout.addLayout(form)

        # Axis Options
        opt_group = QGroupBox("Axis Options")
        opt_layout = QVBoxLayout(opt_group)
        
        self.x_log_check = QCheckBox("X-Axis Log Scale")
        self.y1_log_check = QCheckBox("Y1-Axis Log Scale")
        self.y2_log_check = QCheckBox("Y2-Axis Log Scale")
        self.y1_invert_check = QCheckBox("Invert Y1 Axis")
        self.y2_invert_check = QCheckBox("Invert Y2 Axis")
        self.grid_check = QCheckBox("Show Grid")
        self.grid_check.setChecked(True)
        self.subplot_mode_check = QCheckBox("Subplot Mode (Y1/Y2 separated)")

        opt_layout.addWidget(self.x_log_check)
        opt_layout.addWidget(self.y1_log_check)
        opt_layout.addWidget(self.y2_log_check)
        opt_layout.addWidget(self.y1_invert_check)
        opt_layout.addWidget(self.y2_invert_check)
        opt_layout.addWidget(self.grid_check)
        opt_layout.addWidget(self.subplot_mode_check)
        layout.addWidget(opt_group)

        self.settings_notebook.addTab(tab, "Basic")

    def add_x_tab(self, is_initial=False):
        tab_index = len(self.x_tab_widgets) + 1
        tab_title = f"X-Tab {tab_index}"

        tab_widget = QWidget()
        layout = QVBoxLayout(tab_widget)

        # X-Axis Selector
        x_layout = QHBoxLayout()
        x_layout.addWidget(QLabel("X Axis:"))
        x_combo = QComboBox()
        x_layout.addWidget(x_combo)
        layout.addLayout(x_layout)

        # Y1 and Y2 Column Pickers
        y_layout = QHBoxLayout()
        
        y1_box = QGroupBox("Y1 Series (Left)")
        y1_box_layout = QVBoxLayout(y1_box)
        y1_list = QListWidget()
        y1_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        y1_box_layout.addWidget(y1_list)
        y_layout.addWidget(y1_box)

        y2_box = QGroupBox("Y2 Series (Right)")
        y2_box_layout = QVBoxLayout(y2_box)
        y2_list = QListWidget()
        y2_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        y2_box_layout.addWidget(y2_list)
        y_layout.addWidget(y2_box)

        layout.addLayout(y_layout)

        if not is_initial:
            remove_btn = QPushButton("Delete This Tab")
            remove_btn.clicked.connect(lambda: self.remove_x_tab(tab_widget))
            layout.addWidget(remove_btn)

        self.x_notebook.addTab(tab_widget, tab_title)

        tab_info = {
            "tab_widget": tab_widget,
            "x_combo": x_combo,
            "y1_listbox": y1_list,
            "y2_listbox": y2_list
        }
        self.x_tab_widgets.append(tab_info)

        if self.df is not None:
            self.update_plot_options()

    def remove_x_tab(self, tab_widget):
        if len(self.x_tab_widgets) <= 1:
            QMessageBox.warning(self, "Warning", "Cannot remove the last X-Tab.")
            return

        idx = self.x_notebook.indexOf(tab_widget)
        if idx != -1:
            self.x_notebook.removeTab(idx)
            self.x_tab_widgets = [t for t in self.x_tab_widgets if t["tab_widget"] != tab_widget]

    def create_style_settings_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        target_layout = QHBoxLayout()
        target_layout.addWidget(QLabel("Select Series:"))
        self.combined_style_target_combo = QComboBox()
        self.combined_style_target_combo.currentIndexChanged.connect(self.on_combined_series_select)
        target_layout.addWidget(self.combined_style_target_combo)
        layout.addLayout(target_layout)

        form = QFormLayout()

        # Color picker
        color_layout = QHBoxLayout()
        self.style_color_input = QLineEdit("Auto")
        color_btn = QPushButton("Choose Color...")
        color_btn.clicked.connect(self.on_style_editor_color_pick)
        auto_color_btn = QPushButton("Auto")
        auto_color_btn.clicked.connect(self.on_style_editor_color_auto)
        color_layout.addWidget(self.style_color_input)
        color_layout.addWidget(color_btn)
        color_layout.addWidget(auto_color_btn)
        form.addRow("Color:", color_layout)

        # Line style
        self.style_linestyle_combo = QComboBox()
        self.style_linestyle_combo.addItems(['-', '--', '-.', ':', 'None'])
        form.addRow("Line Style:", self.style_linestyle_combo)

        # Line width
        self.style_linewidth_spin = QDoubleSpinBox()
        self.style_linewidth_spin.setRange(0.1, 10.0)
        self.style_linewidth_spin.setSingleStep(0.5)
        self.style_linewidth_spin.setValue(1.5)
        form.addRow("Line Width:", self.style_linewidth_spin)

        # Marker style
        self.style_marker_combo = QComboBox()
        self.style_marker_combo.addItems(['None', 'o', 's', '^', 'v', 'D', 'x', '+', '*'])
        form.addRow("Marker:", self.style_marker_combo)

        # Marker size
        self.style_markersize_spin = QDoubleSpinBox()
        self.style_markersize_spin.setRange(1.0, 20.0)
        self.style_markersize_spin.setValue(6.0)
        form.addRow("Marker Size:", self.style_markersize_spin)

        # Alpha
        self.style_alpha_spin = QDoubleSpinBox()
        self.style_alpha_spin.setRange(0.0, 1.0)
        self.style_alpha_spin.setSingleStep(0.1)
        self.style_alpha_spin.setValue(1.0)
        form.addRow("Alpha:", self.style_alpha_spin)

        # Display Label
        self.style_label_input = QLineEdit()
        form.addRow("Legend Label:", self.style_label_input)

        layout.addLayout(form)

        apply_btn = QPushButton("Apply Style to Selected Series")
        apply_btn.clicked.connect(self.on_style_editor_change)
        layout.addWidget(apply_btn)

        self.settings_notebook.addTab(tab, "Styles")

    def create_axis_ticks_tab(self):
        tab = QWidget()
        layout = QFormLayout(tab)

        self.xlim_min_input = QLineEdit()
        self.xlim_max_input = QLineEdit()
        self.ylim_min_input = QLineEdit()
        self.ylim_max_input = QLineEdit()
        self.ylim2_min_input = QLineEdit()
        self.ylim2_max_input = QLineEdit()

        layout.addRow("X Min:", self.xlim_min_input)
        layout.addRow("X Max:", self.xlim_max_input)
        layout.addRow("Y1 Min:", self.ylim_min_input)
        layout.addRow("Y1 Max:", self.ylim_max_input)
        layout.addRow("Y2 Min:", self.ylim2_min_input)
        layout.addRow("Y2 Max:", self.ylim2_max_input)

        self.rotate_labels_check = QCheckBox("Rotate X-Tick Labels")
        self.rotation_angle_spin = QSpinBox()
        self.rotation_angle_spin.setRange(0, 360)
        self.rotation_angle_spin.setValue(45)

        layout.addRow(self.rotate_labels_check, self.rotation_angle_spin)

        self.xaxis_plain_check = QCheckBox("X-Axis Plain Format (No Sci notation)")
        self.yaxis1_plain_check = QCheckBox("Y1-Axis Plain Format")
        self.yaxis2_plain_check = QCheckBox("Y2-Axis Plain Format")

        layout.addRow(self.xaxis_plain_check)
        layout.addRow(self.yaxis1_plain_check)
        layout.addRow(self.yaxis2_plain_check)

        self.settings_notebook.addTab(tab, "Axis Limits & Ticks")

    def create_font_size_tab(self):
        tab = QWidget()
        layout = QFormLayout(tab)

        self.font_family_combo = QComboBox()
        self.font_family_combo.addItems(get_font_list())
        layout.addRow("Font Family:", self.font_family_combo)

        self.title_fontsize_spin = QSpinBox()
        self.title_fontsize_spin.setValue(14)
        layout.addRow("Title Font Size:", self.title_fontsize_spin)

        self.xlabel_fontsize_spin = QSpinBox()
        self.xlabel_fontsize_spin.setValue(12)
        layout.addRow("X Label Font Size:", self.xlabel_fontsize_spin)

        self.ylabel_fontsize_spin = QSpinBox()
        self.ylabel_fontsize_spin.setValue(12)
        layout.addRow("Y1 Label Font Size:", self.ylabel_fontsize_spin)

        self.ylabel2_fontsize_spin = QSpinBox()
        self.ylabel2_fontsize_spin.setValue(12)
        layout.addRow("Y2 Label Font Size:", self.ylabel2_fontsize_spin)

        self.tick_fontsize_spin = QSpinBox()
        self.tick_fontsize_spin.setValue(10)
        layout.addRow("Tick Font Size:", self.tick_fontsize_spin)

        self.legend_fontsize_spin = QSpinBox()
        self.legend_fontsize_spin.setValue(10)
        layout.addRow("Legend Font Size:", self.legend_fontsize_spin)

        self.settings_notebook.addTab(tab, "Fonts & Sizes")

    def create_spines_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        self.spine_top_check = QCheckBox("Show Top Spine")
        self.spine_top_check.setChecked(True)
        self.spine_bottom_check = QCheckBox("Show Bottom Spine")
        self.spine_bottom_check.setChecked(True)
        self.spine_left_check = QCheckBox("Show Left Spine")
        self.spine_left_check.setChecked(True)
        self.spine_right_check = QCheckBox("Show Right Spine")
        self.spine_right_check.setChecked(True)

        layout.addWidget(self.spine_top_check)
        layout.addWidget(self.spine_bottom_check)
        layout.addWidget(self.spine_left_check)
        layout.addWidget(self.spine_right_check)

        self.settings_notebook.addTab(tab, "Spines")

    def create_legend_tab(self):
        tab = QWidget()
        layout = QFormLayout(tab)

        self.legend_show_check = QCheckBox("Show Legend")
        self.legend_show_check.setChecked(True)
        layout.addRow(self.legend_show_check)

        self.legend_loc_combo = QComboBox()
        self.legend_loc_combo.addItems([
            'best', 'upper right', 'upper left', 'lower left', 'lower right',
            'right', 'center left', 'center right', 'lower center', 'upper center', 'center'
        ])
        layout.addRow("Legend Position:", self.legend_loc_combo)

        self.settings_notebook.addTab(tab, "Legend")

    def create_advanced_tab(self):
        tab = QWidget()
        layout = QFormLayout(tab)

        # Smoothing
        self.enable_smoothing_check = QCheckBox("Enable Line Smoothing")
        self.smoothing_window_spin = QSpinBox()
        self.smoothing_window_spin.setRange(2, 50)
        self.smoothing_window_spin.setValue(3)
        layout.addRow(self.enable_smoothing_check, self.smoothing_window_spin)

        # Error Bars
        self.enable_errorbar_check = QCheckBox("Enable Error Bars")
        self.errorbar_column_combo = QComboBox()
        layout.addRow(self.enable_errorbar_check, self.errorbar_column_combo)

        # Annotations
        self.enable_annotation_check = QCheckBox("Show Data Point Annotations")
        layout.addRow(self.enable_annotation_check)

        # Filtering
        self.data_filter_check = QCheckBox("Enable Non-Destructive Filter")
        self.filter_column_combo = QComboBox()
        self.filter_min_input = QLineEdit()
        self.filter_max_input = QLineEdit()
        layout.addRow(self.data_filter_check, self.filter_column_combo)
        layout.addRow("Filter Min:", self.filter_min_input)
        layout.addRow("Filter Max:", self.filter_max_input)

        # Colormap & Export DPI
        self.colormap_combo = QComboBox()
        self.colormap_combo.addItems(['viridis', 'plasma', 'inferno', 'magma', 'cividis', 'coolwarm', 'jet', 'rainbow', 'turbo', 'gray'])
        layout.addRow("Colormap:", self.colormap_combo)

        self.export_dpi_spin = QSpinBox()
        self.export_dpi_spin.setRange(72, 1200)
        self.export_dpi_spin.setValue(300)
        layout.addRow("Export DPI:", self.export_dpi_spin)

        self.settings_notebook.addTab(tab, "Advanced")

    # ── Drag and Drop Event Handlers ──────────────────────────────────────────
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.pmggrp':
                self.load_project_file(file_path)
            elif ext in ['.csv', '.tsv', '.xlsx', '.xls', '.txt', '.json']:
                self.load_data(file_path=file_path)

    # ── File & Data Management ───────────────────────────────────────────────
    def load_data(self, file_path=None):
        if not file_path:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Open Data File", "",
                "Supported Files (*.csv *.tsv *.xlsx *.xls *.txt *.json);;CSV Files (*.csv);;All Files (*)"
            )
        if not file_path:
            return

        try:
            self.df = self.data_mgr.load_file(file_path)
            self.populate_data_table()
            self.update_plot_options()
            self.update_style_editor_targets()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load file:\n{e}")

    def populate_data_table(self):
        if self.df is None:
            return

        self.data_table.setRowCount(len(self.df))
        self.data_table.setColumnCount(len(self.df.columns))
        self.data_table.setHorizontalHeaderLabels(list(self.df.columns))

        for row in range(len(self.df)):
            for col in range(len(self.df.columns)):
                val = str(self.df.iat[row, col])
                item = QTableWidgetItem(val)
                self.data_table.setItem(row, col, item)

    def get_data_from_table(self):
        if self.df is None or self.data_table.rowCount() == 0:
            return

        cols = [self.data_table.horizontalHeaderItem(c).text() for c in range(self.data_table.columnCount())]
        data = []
        for r in range(self.data_table.rowCount()):
            row_vals = []
            for c in range(self.data_table.columnCount()):
                item = self.data_table.item(r, c)
                row_vals.append(item.text() if item else "")
            data.append(row_vals)

        new_df = pd.DataFrame(data, columns=cols)
        self.data_mgr.set_dataframe(new_df)
        self.df = self.data_mgr.get_filtered_df()

    def update_plot_options(self):
        if self.df is None:
            return

        columns = self.data_mgr.get_columns()

        for info in self.x_tab_widgets:
            info["x_combo"].clear()
            info["x_combo"].addItems(columns)
            info["y1_listbox"].clear()
            info["y2_listbox"].clear()
            for col in columns:
                info["y1_listbox"].addItem(col)
                info["y2_listbox"].addItem(col)

        if columns and self.x_tab_widgets:
            init_tab = self.x_tab_widgets[0]
            self.xlabel_input.setText(columns[0])
            if len(columns) > 1:
                init_tab["y1_listbox"].item(1).setSelected(True)
            else:
                init_tab["y1_listbox"].item(0).setSelected(True)

        self.errorbar_column_combo.clear()
        self.errorbar_column_combo.addItems([''] + columns)
        self.filter_column_combo.clear()
        self.filter_column_combo.addItems(columns)

    # ── Multi X-Tab Helper ───────────────────────────────────────────────────
    def get_x_tabs_data(self):
        tabs_info = []
        for info in self.x_tab_widgets:
            x_col = info["x_combo"].currentText()
            y1_cols = [item.text() for item in info["y1_listbox"].selectedItems()]
            y2_cols = [item.text() for item in info["y2_listbox"].selectedItems()]
            if x_col:
                tabs_info.append({"x_axis": x_col, "y1_cols": y1_cols, "y2_cols": y2_cols})
        return tabs_info

    # ── Style Editor Interactivity ───────────────────────────────────────────
    def update_style_editor_targets(self):
        self.combined_style_target_combo.clear()
        tabs = self.get_x_tabs_data()
        for idx, tab in enumerate(tabs):
            prefix = f"T{idx+1}-" if len(tabs) > 1 else ""
            for col in tab["y1_cols"]:
                self.combined_style_target_combo.addItem(f"({prefix}Y1) {col}")
            for col in tab["y2_cols"]:
                self.combined_style_target_combo.addItem(f"({prefix}Y2) {col}")

    def on_combined_series_select(self):
        target = self.combined_style_target_combo.currentText()
        if not target:
            return

        if target in self.y1_series_styles:
            style = self.y1_series_styles[target]
        elif target in self.y2_series_styles:
            style = self.y2_series_styles[target]
        else:
            style = {
                'color': 'Auto', 'linestyle': '-', 'linewidth': 1.5,
                'marker': 'None', 'markersize': 6.0, 'alpha': 1.0,
                'label': target.split(') ', 1)[-1] if ') ' in target else target
            }

        self.style_color_input.setText(style.get('color', 'Auto'))
        idx_line = self.style_linestyle_combo.findText(style.get('linestyle', '-'))
        if idx_line >= 0: self.style_linestyle_combo.setCurrentIndex(idx_line)
        self.style_linewidth_spin.setValue(float(style.get('linewidth', 1.5)))
        idx_marker = self.style_marker_combo.findText(style.get('marker', 'None'))
        if idx_marker >= 0: self.style_marker_combo.setCurrentIndex(idx_marker)
        self.style_markersize_spin.setValue(float(style.get('markersize', 6.0)))
        self.style_alpha_spin.setValue(float(style.get('alpha', 1.0)))
        self.style_label_input.setText(style.get('label', ''))

    def on_style_editor_color_pick(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.style_color_input.setText(color.name())
            self.on_style_editor_change()

    def on_style_editor_color_auto(self):
        self.style_color_input.setText("Auto")
        self.on_style_editor_change()

    def on_style_editor_change(self):
        target = self.combined_style_target_combo.currentText()
        if not target:
            return

        style = {
            'color': self.style_color_input.text(),
            'linestyle': self.style_linestyle_combo.currentText(),
            'linewidth': self.style_linewidth_spin.value(),
            'marker': self.style_marker_combo.currentText(),
            'markersize': self.style_markersize_spin.value(),
            'alpha': self.style_alpha_spin.value(),
            'label': self.style_label_input.text()
        }

        if "Y1" in target:
            self.y1_series_styles[target] = style
        else:
            self.y2_series_styles[target] = style

    # ── Core Plotting Engine ─────────────────────────────────────────────────
    def plot_graph(self):
        try:
            self.fig.clear()

            if self.subplot_mode_check.isChecked():
                self.ax = self.fig.add_subplot(211)
                self.ax2 = self.fig.add_subplot(212, sharex=self.ax)
            else:
                self.ax = self.fig.add_subplot(111)
                self.ax2 = None

        except Exception as e:
            QMessageBox.critical(self, "Plot Error", f"Failed to reset canvas:\n{e}")
            return

        self.get_data_from_table()
        if not self.data_mgr.has_data():
            QMessageBox.information(self, "Info", "No data loaded.")
            return

        # Data filtering
        plot_df = self.data_mgr.get_filtered_df(
            filter_enabled=self.data_filter_check.isChecked(),
            filter_column=self.filter_column_combo.currentText(),
            min_val_str=self.filter_min_input.text(),
            max_val_str=self.filter_max_input.text()
        )

        x_tabs_info = self.get_x_tabs_data()
        plot_type = self.plot_type_combo.currentText()

        def plot_series(ax, x_col, y_col, is_twin_ax=False, label_prefix=""):
            series_key = f"({label_prefix}Y2) {y_col}" if is_twin_ax else f"({label_prefix}Y1) {y_col}"
            style_dict = self.y2_series_styles.get(series_key, {}) if is_twin_ax else self.y1_series_styles.get(series_key, {})

            color = style_dict.get('color', 'Auto')
            color = None if color in ['Auto', 'None', ''] else color
            linestyle = style_dict.get('linestyle', '-')
            linewidth = float(style_dict.get('linewidth', 1.5))
            markerstyle = style_dict.get('marker', 'None')
            markersize = float(style_dict.get('markersize', 6.0))
            alpha = float(style_dict.get('alpha', 1.0))
            display_label = style_dict.get('label', f"{y_col} (Y2)" if is_twin_ax else y_col)

            x_data_raw = plot_df[x_col]
            y_data_raw = plot_df[y_col]
            y_cleaned = y_data_raw.astype(str).str.replace(r'[^\d.-]', '', regex=True)
            y_data_numeric = pd.to_numeric(y_cleaned, errors='coerce')

            if plot_type == "bar":
                x_data = x_data_raw.astype(str)
                valid_mask = ~y_data_numeric.isnull()
                ax.bar(x_data[valid_mask], y_data_numeric[valid_mask], alpha=alpha, label=display_label, color=color)
            else:
                x_cleaned = x_data_raw.astype(str).str.replace(r'[^\d.-]', '', regex=True)
                x_numeric = pd.to_numeric(x_cleaned, errors='coerce')
                valid_df = pd.DataFrame({'x': x_numeric, 'y': y_data_numeric}).dropna()
                if valid_df.empty:
                    return

                plot_x, plot_y = valid_df['x'], valid_df['y']

                if plot_type == "line" and self.enable_smoothing_check.isChecked() and len(plot_y) >= self.smoothing_window_spin.value():
                    plot_y = plot_y.rolling(window=self.smoothing_window_spin.value(), center=True).mean().fillna(plot_y)

                errorbar_vals = None
                if (self.enable_errorbar_check.isChecked() and not is_twin_ax
                        and self.errorbar_column_combo.currentText() in plot_df.columns):
                    err_cleaned = plot_df[self.errorbar_column_combo.currentText()].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                    err_numeric = pd.to_numeric(err_cleaned, errors='coerce')
                    errorbar_vals = err_numeric.reindex(valid_df.index).values

                if plot_type == "line":
                    line_kw = {'linestyle': linestyle, 'linewidth': linewidth, 'marker': markerstyle, 'markersize': markersize, 'alpha': alpha, 'label': display_label}
                    if color: line_kw['color'] = color
                    if errorbar_vals is not None:
                        line_kw['yerr'] = errorbar_vals; line_kw['capsize'] = 3
                        ax.errorbar(plot_x, plot_y, **line_kw)
                    else:
                        ax.plot(plot_x, plot_y, **line_kw)
                elif plot_type == "scatter":
                    scatter_kw = {'alpha': alpha, 'label': display_label, 's': markersize**2}
                    if color: scatter_kw['color'] = color
                    if markerstyle != 'None': scatter_kw['marker'] = markerstyle
                    if errorbar_vals is not None:
                        scatter_kw['yerr'] = errorbar_vals; scatter_kw['capsize'] = 3
                        fmt = markerstyle if markerstyle != 'None' else 'o'
                        ax.errorbar(plot_x, plot_y, fmt=fmt, **{k: v for k, v in scatter_kw.items() if k not in ['marker', 's']})
                    else:
                        ax.scatter(plot_x, plot_y, **scatter_kw)
                elif plot_type == "step":
                    step_kw = {'linestyle': linestyle, 'linewidth': linewidth, 'alpha': alpha, 'label': display_label}
                    if color: step_kw['color'] = color
                    ax.step(plot_x, plot_y, where='mid', **step_kw)
                elif plot_type == "area":
                    area_kw = {'linestyle': linestyle, 'linewidth': linewidth, 'alpha': alpha, 'label': display_label}
                    if color: area_kw['color'] = color
                    ax.fill_between(plot_x, 0, plot_y, **area_kw)
                elif plot_type == "stem":
                    stem_kw = {'label': display_label}
                    if color: stem_kw['linefmt'] = color; stem_kw['markerfmt'] = color + 'o'
                    ml, sl, _ = ax.stem(plot_x, plot_y, **stem_kw)
                    ml.set_alpha(alpha); sl.set_alpha(alpha)

                if self.enable_annotation_check.isChecked() and not is_twin_ax:
                    for i, (xi, yi) in enumerate(zip(plot_x, plot_y)):
                        if i % max(1, len(plot_x) // 10) == 0:
                            ax.annotate(f'{yi:.2f}', (xi, yi), textcoords='offset points', xytext=(0, 5), ha='center', fontsize=8)

        # Handle Special Plot Types (Pie, Box, Violin, Heatmap, Contour, Polar)
        _first_tab = x_tabs_info[0] if x_tabs_info else {}
        _x_col_0 = _first_tab.get("x_axis", "")
        _y1_cols_0 = _first_tab.get("y1_cols", [])
        _x0_raw = plot_df[_x_col_0] if _x_col_0 and _x_col_0 in plot_df.columns else None

        def _to_numeric_series(col):
            return pd.to_numeric(plot_df[col].astype(str).str.replace(r'[^\d.-]', '', regex=True), errors='coerce')

        def _finish_special():
            self.fig.tight_layout()
            self.canvas.draw()

        if plot_type == "pie":
            if not _y1_cols_0:
                QMessageBox.warning(self, "Warning", "Pie chart requires at least one Y column.")
                return
            valid = pd.DataFrame({'x': _x0_raw, 'y': _to_numeric_series(_y1_cols_0[0])}).dropna()
            self.ax.pie(valid['y'], labels=valid['x'].astype(str), autopct='%1.1f%%', startangle=90)
            self.ax.axis('equal')
            if self.title_input.text(): self.ax.set_title(self.title_input.text())
            _finish_special()
            return

        elif plot_type == "box":
            box_data = [_to_numeric_series(yc).dropna() for yc in _y1_cols_0 if len(_to_numeric_series(yc).dropna()) > 0]
            if box_data:
                self.ax.boxplot(box_data, labels=_y1_cols_0)
                if self.title_input.text(): self.ax.set_title(self.title_input.text())
                if self.grid_check.isChecked(): self.ax.grid(True)
            _finish_special()
            return

        elif plot_type == "violin":
            v_data = [_to_numeric_series(yc).dropna() for yc in _y1_cols_0 if len(_to_numeric_series(yc).dropna()) > 0]
            if v_data:
                self.ax.violinplot(v_data, showmeans=True, showmedians=True)
                self.ax.set_xticks(range(1, len(_y1_cols_0) + 1))
                self.ax.set_xticklabels(_y1_cols_0)
                if self.title_input.text(): self.ax.set_title(self.title_input.text())
                if self.grid_check.isChecked(): self.ax.grid(True)
            _finish_special()
            return

        elif plot_type == "heatmap":
            if not _y1_cols_0: return
            hmap = [_to_numeric_series(yc).fillna(0).values for yc in _y1_cols_0]
            im = self.ax.imshow(np.array(hmap), aspect='auto', cmap=self.colormap_combo.currentText())
            self.ax.set_yticks(range(len(_y1_cols_0)))
            self.ax.set_yticklabels(_y1_cols_0)
            self.fig.colorbar(im, ax=self.ax)
        elif plot_type == "contour":
            if len(_y1_cols_0) < 2 or _x0_raw is None:
                QMessageBox.warning(self, "Warning", "Contour plot requires at least 2 Y columns (Y-coord, Z-value).")
                return
            x_num = pd.to_numeric(_x0_raw.astype(str).str.replace(r'[^\d.-]', '', regex=True), errors='coerce')
            y_num = _to_numeric_series(_y1_cols_0[0])
            z_num = _to_numeric_series(_y1_cols_0[1])
            valid = pd.DataFrame({'x': x_num, 'y': y_num, 'z': z_num}).dropna()
            if len(valid) >= 4 and valid['x'].nunique() > 1 and valid['y'].nunique() > 1:
                xi = np.linspace(valid['x'].min(), valid['x'].max(), 50)
                yi = np.linspace(valid['y'].min(), valid['y'].max(), 50)
                Xi, Yi = np.meshgrid(xi, yi)
                try:
                    Zi = griddata((valid['x'].values, valid['y'].values), valid['z'].values, (Xi, Yi), method='linear')
                except Exception:
                    Zi = griddata((valid['x'].values, valid['y'].values), valid['z'].values, (Xi, Yi), method='nearest')
                cf = self.ax.contourf(Xi, Yi, Zi, levels=15, cmap=self.colormap_combo.currentText())
                self.fig.colorbar(cf, ax=self.ax)
            _finish_special()
            return

        elif plot_type == "polar":
            self.fig.clear()
            self.ax = self.fig.add_subplot(111, projection='polar')
            x_num = pd.to_numeric(_x0_raw.astype(str).str.replace(r'[^\d.-]', '', regex=True), errors='coerce') if _x0_raw is not None else pd.Series()
            for yc in _y1_cols_0:
                valid = pd.DataFrame({'x': x_num, 'y': _to_numeric_series(yc)}).dropna()
                if not valid.empty:
                    self.ax.plot(np.radians(valid['x'].values), valid['y'].values, label=yc, marker='o')
            if self.legend_show_check.isChecked(): self.ax.legend()
            _finish_special()
            return

        # Standard Multi-Tab Plot Loop
        has_twin_y2 = any(len(tab["y2_cols"]) > 0 for tab in x_tabs_info)
        if has_twin_y2 and not self.subplot_mode_check.isChecked():
            self.ax2 = self.ax.twinx()

        for tab_idx, tab_data in enumerate(x_tabs_info):
            x_col = tab_data["x_axis"]
            if not x_col or x_col not in plot_df.columns:
                continue
            prefix = f"T{tab_idx+1}-" if len(x_tabs_info) > 1 else ""

            for y_col in tab_data["y1_cols"]:
                if y_col in plot_df.columns:
                    plot_series(self.ax, x_col, y_col, is_twin_ax=False, label_prefix=prefix)

            target_ax = self.ax2 if self.ax2 else self.ax
            for y_col in tab_data["y2_cols"]:
                if y_col in plot_df.columns:
                    plot_series(target_ax, x_col, y_col, is_twin_ax=True, label_prefix=prefix)

        # Titles and Formatting
        font_family = self.font_family_combo.currentText()
        first_x_col = x_tabs_info[0]["x_axis"] if x_tabs_info else ""
        y1_names = [c for tab in x_tabs_info for c in tab["y1_cols"]]
        y2_names = [c for tab in x_tabs_info for c in tab["y2_cols"]]

        self.ax.set_title(self.title_input.text(), fontsize=self.title_fontsize_spin.value(), fontfamily=font_family)
        self.ax.set_xlabel(self.xlabel_input.text() if self.xlabel_input.text() else first_x_col, fontsize=self.xlabel_fontsize_spin.value(), fontfamily=font_family)
        self.ax.set_ylabel(self.ylabel_input.text() if self.ylabel_input.text() else ", ".join(y1_names), fontsize=self.ylabel_fontsize_spin.value(), fontfamily=font_family)

        self.set_axis_limits(self.ax, 'x', self.xlim_min_input.text(), self.xlim_max_input.text())
        self.set_axis_limits(self.ax, 'y', self.ylim_min_input.text(), self.ylim_max_input.text())

        if self.grid_check.isChecked():
            self.ax.grid(True)

        if self.rotate_labels_check.isChecked():
            self.ax.tick_params(axis='x', rotation=self.rotation_angle_spin.value())

        if self.legend_show_check.isChecked():
            h1, l1 = self.ax.get_legend_handles_labels()
            h2, l2 = self.ax2.get_legend_handles_labels() if self.ax2 else ([], [])
            self.ax.legend(h1 + h2, l1 + l2, loc=self.legend_loc_combo.currentText())

        self.fig.tight_layout()
        self.canvas.draw()

    def set_axis_limits(self, ax, axis, min_val_str, max_val_str):
        try:
            val_min = float(min_val_str) if min_val_str else None
            val_max = float(max_val_str) if max_val_str else None
            if axis == 'x':
                if val_min is not None and val_max is not None: ax.set_xlim(val_min, val_max)
                elif val_min is not None: ax.set_xlim(left=val_min)
                elif val_max is not None: ax.set_xlim(right=val_max)
            elif axis == 'y':
                if val_min is not None and val_max is not None: ax.set_ylim(val_min, val_max)
                elif val_min is not None: ax.set_ylim(bottom=val_min)
                elif val_max is not None: ax.set_ylim(top=val_max)
        except ValueError:
            pass

    # ── Import / Export & Project Files ──────────────────────────────────────
    def save_settings(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Project", "", "Matplotlib Graph Project (*.pmggrp)")
        if file_path:
            save_project_file(self, file_path, version_str=VERSION, dimension="2D")
            self.current_project_path = file_path

    def overwrite_save(self):
        if not self.current_project_path:
            self.save_settings()
        else:
            save_project_file(self, self.current_project_path, version_str=VERSION, dimension="2D")

    def load_settings(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Project", "", "Matplotlib Graph Project (*.pmggrp)")
        if file_path:
            self.load_project_file(file_path)

    def load_project_file(self, file_path):
        load_project_file(self, file_path)
        self.current_project_path = file_path

    def export_graph(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Export Plot Image", "", "PNG Image (*.png);;PDF Document (*.pdf);;SVG Image (*.svg)")
        if file_path:
            self.fig.savefig(file_path, dpi=self.export_dpi_spin.value(), bbox_inches='tight')
            QMessageBox.information(self, "Success", f"Plot saved to {file_path}")

    def export_filtered_data(self):
        if self.df is None:
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "Export Data", "", "CSV File (*.csv);;Excel File (*.xlsx)")
        if file_path:
            df_export = self.data_mgr.get_filtered_df(
                filter_enabled=self.data_filter_check.isChecked(),
                filter_column=self.filter_column_combo.currentText(),
                min_val_str=self.filter_min_input.text(),
                max_val_str=self.filter_max_input.text()
            )
            if file_path.endswith('.xlsx'):
                df_export.to_excel(file_path, index=False)
            else:
                df_export.to_csv(file_path, index=False)
            QMessageBox.information(self, "Success", f"Data exported to {file_path}")

    def clear_all(self):
        self.data_mgr.clear()
        self.df = None
        self.data_table.clear()
        self.data_table.setRowCount(0)
        self.data_table.setColumnCount(0)
        self.fig.clear()
        self.ax = self.fig.add_subplot(111)
        self.canvas.draw()

    def reset_settings(self):
        self.plot_type_combo.setCurrentIndex(0)
        self.title_input.clear()
        self.xlabel_input.clear()
        self.ylabel_input.clear()
        self.ylabel2_input.clear()
        self.grid_check.setChecked(True)
        self.clear_all()

    def open_in_3d_mode(self):
        try:
            from hygrapher.main_3d import GraphApp3D
            self.win_3d = GraphApp3D()
            self.win_3d.show()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open 3D Mode:\n{e}")

    def show_about(self):
        QMessageBox.about(self, "About HyGrapher", f"HyGrapher v{VERSION}\nA cross-platform Matplotlib GUI application built with PyQt6.")


def main():
    app = QApplication(sys.argv)
    window = GraphApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
