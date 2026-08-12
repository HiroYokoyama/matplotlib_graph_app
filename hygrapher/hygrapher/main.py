# -*- coding: utf-8 -*-
"""
HyGrapher - Matplotlib Plotting Desktop Application (PyQt6 Edition)
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
    QTabWidget,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QLabel,
    QPushButton,
    QComboBox,
    QLineEdit,
    QCheckBox,
    QSpinBox,
    QDoubleSpinBox,
    QGroupBox,
    QListWidget,
    QAbstractItemView,
    QFileDialog,
    QMessageBox,
    QColorDialog,
    QFormLayout,
    QSizePolicy,
    QDialog,
    QMenu,
    QInputDialog,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import (
    QAction,
    QColor,
    QDropEvent,
    QDragEnterEvent,
    QKeySequence,
    QUndoStack,
)

import matplotlib

matplotlib.use("QtAgg")
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg, NavigationToolbar2QT
from matplotlib.figure import Figure

from hygrapher.data_manager import DataManager
from hygrapher.project_io import save_project_file, load_project_file, reset_to_defaults
from hygrapher.utils import (
    apply_major_ticker,
    apply_minor_ticker,
    get_app_version,
    get_font_list,
    resolve_cli_file,
)
from hygrapher.import_dialog import ImportPreviewDialog
from hygrapher.table_undo import (
    DeleteColumnCommand,
    DeleteRowCommand,
    InsertColumnCommand,
    InsertRowCommand,
    handle_table_item_changed,
    snapshot_table,
)

VERSION = get_app_version()


class GraphApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(1400, 900)
        self.setMinimumSize(1000, 650)
        self.setAcceptDrops(True)

        self.data_mgr = DataManager()
        self.df = None
        self.current_project_path = None

        self.x_tab_widgets = []
        self.y1_series_styles = {}
        self.y2_series_styles = {}

        self.undo_stack = QUndoStack(self)
        self.undo_stack.cleanChanged.connect(self.update_window_title)
        self._table_snapshot = {}

        self.init_ui()
        self.update_window_title()

    def init_ui(self):
        # Menu Bar
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("File")

        open_action = QAction("Open Data File...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.import_data_interactive)
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

        edit_menu = menu_bar.addMenu("Edit")
        undo_action = self.undo_stack.createUndoAction(self, "Undo")
        undo_action.setShortcut(QKeySequence.StandardKey.Undo)
        edit_menu.addAction(undo_action)

        redo_action = self.undo_stack.createRedoAction(self, "Redo")
        redo_action.setShortcut(QKeySequence.StandardKey.Redo)
        edit_menu.addAction(redo_action)

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
        left_panel.setMinimumWidth(360)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Top Control Buttons
        top_btn_layout = QHBoxLayout()
        self.open_file_btn = QPushButton("Open File")
        self.open_file_btn.clicked.connect(self.import_data_interactive)
        top_btn_layout.addWidget(self.open_file_btn)

        self.plot_button = QPushButton("Plot Graph")
        self.plot_button.setStyleSheet(
            "font-weight: bold; background-color: #2b5c8f; color: white;"
        )
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
        self.data_table.itemChanged.connect(self.on_table_item_changed)
        self.data_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.data_table.customContextMenuRequested.connect(
            self.show_table_context_menu
        )
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
        self.canvas.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self.canvas.setMinimumSize(300, 250)
        self.toolbar = NavigationToolbar2QT(self.canvas, self)

        right_layout.addWidget(self.toolbar)
        right_layout.addWidget(self.canvas, stretch=1)

        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([480, 920])

    def create_basic_settings_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Plot type
        pt_layout = QHBoxLayout()
        pt_layout.addWidget(QLabel("Plot Type:"))
        self.plot_type_combo = QComboBox()
        self.plot_type_combo.addItems(
            [
                "line",
                "scatter",
                "bar",
                "step",
                "stem",
                "area",
                "pie",
                "box",
                "violin",
                "heatmap",
                "contour",
                "polar",
            ]
        )
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
            "y2_listbox": y2_list,
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
            self.x_tab_widgets = [
                t for t in self.x_tab_widgets if t["tab_widget"] != tab_widget
            ]

    def create_style_settings_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        target_layout = QHBoxLayout()
        target_layout.addWidget(QLabel("Select Series:"))
        self.combined_style_target_combo = QComboBox()
        self.combined_style_target_combo.currentIndexChanged.connect(
            self.on_combined_series_select
        )
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
        self.style_linestyle_combo.addItems(["-", "--", "-.", ":", "None"])
        form.addRow("Line Style:", self.style_linestyle_combo)

        # Line width
        self.style_linewidth_spin = QDoubleSpinBox()
        self.style_linewidth_spin.setRange(0.1, 10.0)
        self.style_linewidth_spin.setSingleStep(0.5)
        self.style_linewidth_spin.setValue(1.5)
        form.addRow("Line Width:", self.style_linewidth_spin)

        # Marker style
        self.style_marker_combo = QComboBox()
        self.style_marker_combo.addItems(
            ["None", "o", "s", "^", "v", "D", "x", "+", "*"]
        )
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

        self.xtick_major_interval_input = QLineEdit()
        self.ytick_major_interval_input = QLineEdit()
        self.ytick2_major_interval_input = QLineEdit()

        layout.addRow("X Major Tick Interval:", self.xtick_major_interval_input)
        layout.addRow("Y1 Major Tick Interval:", self.ytick_major_interval_input)
        layout.addRow("Y2 Major Tick Interval:", self.ytick2_major_interval_input)

        self.xtick_minor_check = QCheckBox("Show X Minor Ticks")
        self.xtick_minor_interval_input = QLineEdit()
        layout.addRow(self.xtick_minor_check, self.xtick_minor_interval_input)

        self.ytick_minor_check = QCheckBox("Show Y1 Minor Ticks")
        self.ytick_minor_interval_input = QLineEdit()
        layout.addRow(self.ytick_minor_check, self.ytick_minor_interval_input)

        self.ytick2_minor_check = QCheckBox("Show Y2 Minor Ticks")
        self.ytick2_minor_interval_input = QLineEdit()
        layout.addRow(self.ytick2_minor_check, self.ytick2_minor_interval_input)

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

        self.tick2_fontsize_spin = QSpinBox()
        self.tick2_fontsize_spin.setValue(10)
        layout.addRow("Y2 Tick Font Size:", self.tick2_fontsize_spin)

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
        self.legend_loc_combo.addItems(
            [
                "best",
                "upper right",
                "upper left",
                "lower left",
                "lower right",
                "right",
                "center left",
                "center right",
                "lower center",
                "upper center",
                "center",
            ]
        )
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
                "turbo",
                "gray",
            ]
        )
        layout.addRow("Colormap:", self.colormap_combo)

        self.export_dpi_spin = QSpinBox()
        self.export_dpi_spin.setRange(72, 1200)
        self.export_dpi_spin.setValue(300)
        self.grid_alpha_spin = QDoubleSpinBox()
        self.grid_alpha_spin.setRange(0.0, 1.0)
        self.grid_alpha_spin.setSingleStep(0.1)
        self.grid_alpha_spin.setValue(0.3)
        layout.addRow("Grid Alpha:", self.grid_alpha_spin)

        self.grid_linestyle_combo = QComboBox()
        self.grid_linestyle_combo.addItems(["--", "-", "-.", ":"])
        layout.addRow("Grid Line Style:", self.grid_linestyle_combo)

        self.grid_linewidth_spin = QDoubleSpinBox()
        self.grid_linewidth_spin.setRange(0.1, 5.0)
        self.grid_linewidth_spin.setSingleStep(0.1)
        self.grid_linewidth_spin.setValue(0.5)
        layout.addRow("Grid Line Width:", self.grid_linewidth_spin)

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
            if ext == ".pmggrp":
                self.load_project_file(file_path)
            elif ext in [".csv", ".tsv", ".xlsx", ".xls", ".txt", ".json"]:
                self.import_data_interactive(file_path=file_path)

    # ── File & Data Management ───────────────────────────────────────────────
    def import_data_interactive(self, file_path=None):
        """
        Interactive entry point for Open File / menu / drag-drop: shows the
        file picker if needed, then (for row-oriented formats) a preview
        dialog to pick which row is the header before actually loading.
        """
        if not file_path:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Open Data File",
                "",
                "Supported Files (*.csv *.tsv *.xlsx *.xls *.txt *.json);;CSV Files (*.csv);;All Files (*)",
            )
        if not file_path:
            return

        header_row = 0
        ext = os.path.splitext(file_path)[1].lower()
        if ext != ".json":
            try:
                preview_rows = self.data_mgr.read_preview_rows(file_path)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to preview file:\n{e}")
                return
            if len(preview_rows) > 1:
                dialog = ImportPreviewDialog(file_path, preview_rows, parent=self)
                if dialog.exec() != QDialog.DialogCode.Accepted:
                    return
                header_row = dialog.header_row

        self.load_data(file_path=file_path, header_row=header_row)

    def load_data(self, file_path=None, header_row=0):
        if not file_path:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Open Data File",
                "",
                "Supported Files (*.csv *.tsv *.xlsx *.xls *.txt *.json);;CSV Files (*.csv);;All Files (*)",
            )
        if not file_path:
            return

        try:
            self.df = self.data_mgr.load_file(file_path, header_row=header_row)
            self.populate_data_table()
            self.update_plot_options()
            self.update_style_editor_targets()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load file:\n{e}")

    def populate_data_table(self):
        if self.df is None:
            return

        self.data_table.blockSignals(True)
        self.data_table.setRowCount(len(self.df))
        self.data_table.setColumnCount(len(self.df.columns))
        self.data_table.setHorizontalHeaderLabels(list(self.df.columns))

        for row in range(len(self.df)):
            for col in range(len(self.df.columns)):
                val = str(self.df.iat[row, col])
                item = QTableWidgetItem(val)
                self.data_table.setItem(row, col, item)
        self.data_table.blockSignals(False)

        # A freshly (re)populated table is a new document as far as undo
        # history and cell-edit tracking are concerned.
        self._table_snapshot = snapshot_table(self.data_table)
        self.undo_stack.clear()

    def on_table_item_changed(self, item):
        handle_table_item_changed(
            item, self.data_table, self._table_snapshot, self.undo_stack
        )

    # ── Row/Column Insert & Delete (undoable) ────────────────────────────────
    def show_table_context_menu(self, pos):
        if self.data_table.columnCount() == 0:
            return
        menu = QMenu(self)
        menu.addAction("Insert Row Above", self.insert_row_above)
        menu.addAction("Insert Row Below", self.insert_row_below)
        menu.addAction("Delete Row", self.delete_selected_row)
        menu.addSeparator()
        menu.addAction("Insert Column Left", self.insert_column_left)
        menu.addAction("Insert Column Right", self.insert_column_right)
        menu.addAction("Delete Column", self.delete_selected_column)
        menu.exec(self.data_table.viewport().mapToGlobal(pos))

    def insert_row_above(self):
        row = self.data_table.currentRow()
        self._insert_row(row if row >= 0 else 0)

    def insert_row_below(self):
        row = self.data_table.currentRow()
        self._insert_row(row + 1 if row >= 0 else self.data_table.rowCount())

    def _insert_row(self, row):
        if self.data_table.columnCount() == 0:
            return
        cmd = InsertRowCommand(
            self.data_table,
            self._table_snapshot,
            row,
            self.data_table.columnCount(),
            on_changed=self.get_data_from_table,
        )
        self.undo_stack.push(cmd)

    def delete_selected_row(self):
        row = self.data_table.currentRow()
        if row < 0:
            return
        if self.data_table.rowCount() <= 1:
            QMessageBox.warning(self, "Warning", "Cannot delete the last row.")
            return
        row_data = [
            self.data_table.item(row, c).text() if self.data_table.item(row, c) else ""
            for c in range(self.data_table.columnCount())
        ]
        cmd = DeleteRowCommand(
            self.data_table,
            self._table_snapshot,
            row,
            row_data,
            on_changed=self.get_data_from_table,
        )
        self.undo_stack.push(cmd)

    def insert_column_left(self):
        col = self.data_table.currentColumn()
        self._insert_column(col if col >= 0 else 0)

    def insert_column_right(self):
        col = self.data_table.currentColumn()
        self._insert_column(col + 1 if col >= 0 else self.data_table.columnCount())

    def _insert_column(self, col):
        if self.data_table.rowCount() == 0:
            return
        default_name = f"Column{self.data_table.columnCount() + 1}"
        name, ok = QInputDialog.getText(
            self, "Insert Column", "Column name:", text=default_name
        )
        if not ok:
            return
        cmd = InsertColumnCommand(
            self.data_table,
            self._table_snapshot,
            col,
            name.strip() or default_name,
            self.data_table.rowCount(),
            on_changed=self._sync_after_column_structure_change,
        )
        self.undo_stack.push(cmd)

    def delete_selected_column(self):
        col = self.data_table.currentColumn()
        if col < 0:
            return
        if self.data_table.columnCount() <= 1:
            QMessageBox.warning(self, "Warning", "Cannot delete the last column.")
            return
        header_item = self.data_table.horizontalHeaderItem(col)
        header_text = header_item.text() if header_item else f"Column{col + 1}"
        col_data = [
            self.data_table.item(r, col).text() if self.data_table.item(r, col) else ""
            for r in range(self.data_table.rowCount())
        ]
        cmd = DeleteColumnCommand(
            self.data_table,
            self._table_snapshot,
            col,
            header_text,
            col_data,
            on_changed=self._sync_after_column_structure_change,
        )
        self.undo_stack.push(cmd)

    def _sync_after_column_structure_change(self):
        self.get_data_from_table()
        self.update_plot_options()
        self.update_style_editor_targets()

    def get_data_from_table(self):
        if self.df is None or self.data_table.rowCount() == 0:
            return

        cols = [
            self.data_table.horizontalHeaderItem(c).text()
            for c in range(self.data_table.columnCount())
        ]
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
        self.errorbar_column_combo.addItems([""] + columns)
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
                tabs_info.append(
                    {"x_axis": x_col, "y1_cols": y1_cols, "y2_cols": y2_cols}
                )
        return tabs_info

    # ── Style Editor Interactivity ───────────────────────────────────────────
    def update_style_editor_targets(self):
        self.combined_style_target_combo.clear()
        tabs = self.get_x_tabs_data()
        for idx, tab in enumerate(tabs):
            prefix = f"T{idx + 1}-" if len(tabs) > 1 else ""
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
                "color": "Auto",
                "linestyle": "-",
                "linewidth": 1.5,
                "marker": "None",
                "markersize": 6.0,
                "alpha": 1.0,
                "label": target.split(") ", 1)[-1] if ") " in target else target,
            }

        self.style_color_input.setText(style.get("color", "Auto"))
        idx_line = self.style_linestyle_combo.findText(style.get("linestyle", "-"))
        if idx_line >= 0:
            self.style_linestyle_combo.setCurrentIndex(idx_line)
        self.style_linewidth_spin.setValue(float(style.get("linewidth", 1.5)))
        idx_marker = self.style_marker_combo.findText(style.get("marker", "None"))
        if idx_marker >= 0:
            self.style_marker_combo.setCurrentIndex(idx_marker)
        self.style_markersize_spin.setValue(float(style.get("markersize", 6.0)))
        self.style_alpha_spin.setValue(float(style.get("alpha", 1.0)))
        self.style_label_input.setText(style.get("label", ""))

    def on_style_editor_color_pick(self):
        current = self.style_color_input.text()
        initial = QColor(current) if current and current != "Auto" else QColor("white")
        color = QColorDialog.getColor(initial, self, "Choose Series Color")
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
            "color": self.style_color_input.text(),
            "linestyle": self.style_linestyle_combo.currentText(),
            "linewidth": self.style_linewidth_spin.value(),
            "marker": self.style_marker_combo.currentText(),
            "markersize": self.style_markersize_spin.value(),
            "alpha": self.style_alpha_spin.value(),
            "label": self.style_label_input.text(),
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
            max_val_str=self.filter_max_input.text(),
        )

        x_tabs_info = self.get_x_tabs_data()
        plot_type = self.plot_type_combo.currentText()

        def plot_series(ax, x_col, y_col, is_twin_ax=False, label_prefix=""):
            series_key = (
                f"({label_prefix}Y2) {y_col}"
                if is_twin_ax
                else f"({label_prefix}Y1) {y_col}"
            )
            style_dict = (
                self.y2_series_styles.get(series_key, {})
                if is_twin_ax
                else self.y1_series_styles.get(series_key, {})
            )

            color = style_dict.get("color", "Auto")
            color = None if color in ["Auto", "None", ""] else color
            linestyle = style_dict.get("linestyle", "-")
            linewidth = float(style_dict.get("linewidth", 1.5))
            markerstyle = style_dict.get("marker", "None")
            markersize = float(style_dict.get("markersize", 6.0))
            alpha = float(style_dict.get("alpha", 1.0))
            display_label = style_dict.get(
                "label", f"{y_col} (Y2)" if is_twin_ax else y_col
            )

            x_data_raw = plot_df[x_col]
            y_data_raw = plot_df[y_col]
            y_cleaned = y_data_raw.astype(str).str.replace(r"[^\d.-]", "", regex=True)
            y_data_numeric = pd.to_numeric(y_cleaned, errors="coerce")

            if plot_type == "bar":
                x_data = x_data_raw.astype(str)
                valid_mask = ~y_data_numeric.isnull()
                ax.bar(
                    x_data[valid_mask],
                    y_data_numeric[valid_mask],
                    alpha=alpha,
                    label=display_label,
                    color=color,
                )
            else:
                x_cleaned = x_data_raw.astype(str).str.replace(
                    r"[^\d.-]", "", regex=True
                )
                x_numeric = pd.to_numeric(x_cleaned, errors="coerce")
                valid_df = pd.DataFrame({"x": x_numeric, "y": y_data_numeric}).dropna()
                if valid_df.empty:
                    return

                plot_x, plot_y = valid_df["x"], valid_df["y"]

                if (
                    plot_type == "line"
                    and self.enable_smoothing_check.isChecked()
                    and len(plot_y) >= self.smoothing_window_spin.value()
                ):
                    plot_y = (
                        plot_y.rolling(
                            window=self.smoothing_window_spin.value(), center=True
                        )
                        .mean()
                        .fillna(plot_y)
                    )

                errorbar_vals = None
                if (
                    self.enable_errorbar_check.isChecked()
                    and not is_twin_ax
                    and self.errorbar_column_combo.currentText() in plot_df.columns
                ):
                    err_cleaned = (
                        plot_df[self.errorbar_column_combo.currentText()]
                        .astype(str)
                        .str.replace(r"[^\d.-]", "", regex=True)
                    )
                    err_numeric = pd.to_numeric(err_cleaned, errors="coerce")
                    errorbar_vals = err_numeric.reindex(valid_df.index).values

                if plot_type == "line":
                    line_kw = {
                        "linestyle": linestyle,
                        "linewidth": linewidth,
                        "marker": markerstyle,
                        "markersize": markersize,
                        "alpha": alpha,
                        "label": display_label,
                    }
                    if color:
                        line_kw["color"] = color
                    if errorbar_vals is not None:
                        line_kw["yerr"] = errorbar_vals
                        line_kw["capsize"] = 3
                        ax.errorbar(plot_x, plot_y, **line_kw)
                    else:
                        ax.plot(plot_x, plot_y, **line_kw)
                elif plot_type == "scatter":
                    scatter_kw = {
                        "alpha": alpha,
                        "label": display_label,
                        "s": markersize**2,
                    }
                    if color:
                        scatter_kw["color"] = color
                    if markerstyle != "None":
                        scatter_kw["marker"] = markerstyle
                    if errorbar_vals is not None:
                        scatter_kw["yerr"] = errorbar_vals
                        scatter_kw["capsize"] = 3
                        fmt = markerstyle if markerstyle != "None" else "o"
                        ax.errorbar(
                            plot_x,
                            plot_y,
                            fmt=fmt,
                            **{
                                k: v
                                for k, v in scatter_kw.items()
                                if k not in ["marker", "s"]
                            },
                        )
                    else:
                        ax.scatter(plot_x, plot_y, **scatter_kw)
                elif plot_type == "step":
                    step_kw = {
                        "linestyle": linestyle,
                        "linewidth": linewidth,
                        "alpha": alpha,
                        "label": display_label,
                    }
                    if color:
                        step_kw["color"] = color
                    ax.step(plot_x, plot_y, where="mid", **step_kw)
                elif plot_type == "area":
                    area_kw = {
                        "linestyle": linestyle,
                        "linewidth": linewidth,
                        "alpha": alpha,
                        "label": display_label,
                    }
                    if color:
                        area_kw["color"] = color
                    ax.fill_between(plot_x, 0, plot_y, **area_kw)
                elif plot_type == "stem":
                    stem_kw = {"label": display_label}
                    if color:
                        stem_kw["linefmt"] = color
                        stem_kw["markerfmt"] = color + "o"
                    ml, sl, _ = ax.stem(plot_x, plot_y, **stem_kw)
                    ml.set_alpha(alpha)
                    sl.set_alpha(alpha)

                if self.enable_annotation_check.isChecked() and not is_twin_ax:
                    for i, (xi, yi) in enumerate(zip(plot_x, plot_y)):
                        if i % max(1, len(plot_x) // 10) == 0:
                            ax.annotate(
                                f"{yi:.2f}",
                                (xi, yi),
                                textcoords="offset points",
                                xytext=(0, 5),
                                ha="center",
                                fontsize=8,
                            )

        # Handle Special Plot Types (Pie, Box, Violin, Heatmap, Contour, Polar)
        _first_tab = x_tabs_info[0] if x_tabs_info else {}
        _x_col_0 = _first_tab.get("x_axis", "")
        _y1_cols_0 = _first_tab.get("y1_cols", [])
        _x0_raw = (
            plot_df[_x_col_0] if _x_col_0 and _x_col_0 in plot_df.columns else None
        )

        def _to_numeric_series(col):
            return pd.to_numeric(
                plot_df[col].astype(str).str.replace(r"[^\d.-]", "", regex=True),
                errors="coerce",
            )

        def _finish_special():
            self.fig.tight_layout()
            self.canvas.draw()

        if plot_type == "pie":
            if not _y1_cols_0:
                QMessageBox.warning(
                    self, "Warning", "Pie chart requires at least one Y column."
                )
                return
            valid = pd.DataFrame(
                {"x": _x0_raw, "y": _to_numeric_series(_y1_cols_0[0])}
            ).dropna()
            self.ax.pie(
                valid["y"],
                labels=valid["x"].astype(str),
                autopct="%1.1f%%",
                startangle=90,
            )
            self.ax.axis("equal")
            if self.title_input.text():
                self.ax.set_title(self.title_input.text())
            _finish_special()
            return

        elif plot_type == "box":
            box_data = [
                _to_numeric_series(yc).dropna()
                for yc in _y1_cols_0
                if len(_to_numeric_series(yc).dropna()) > 0
            ]
            if box_data:
                try:
                    self.ax.boxplot(box_data, tick_labels=_y1_cols_0)
                except TypeError:
                    self.ax.boxplot(box_data, labels=_y1_cols_0)
                if self.title_input.text():
                    self.ax.set_title(self.title_input.text())
                if self.grid_check.isChecked():
                    self.ax.grid(True)
            _finish_special()
            return

        elif plot_type == "violin":
            v_data = [
                _to_numeric_series(yc).dropna()
                for yc in _y1_cols_0
                if len(_to_numeric_series(yc).dropna()) > 0
            ]
            if v_data:
                self.ax.violinplot(v_data, showmeans=True, showmedians=True)
                self.ax.set_xticks(range(1, len(_y1_cols_0) + 1))
                self.ax.set_xticklabels(_y1_cols_0)
                if self.title_input.text():
                    self.ax.set_title(self.title_input.text())
                if self.grid_check.isChecked():
                    self.ax.grid(True)
            _finish_special()
            return

        elif plot_type == "heatmap":
            if not _y1_cols_0:
                return
            hmap = [_to_numeric_series(yc).fillna(0).values for yc in _y1_cols_0]
            im = self.ax.imshow(
                np.array(hmap), aspect="auto", cmap=self.colormap_combo.currentText()
            )
            self.ax.set_yticks(range(len(_y1_cols_0)))
            self.ax.set_yticklabels(_y1_cols_0)
            self.fig.colorbar(im, ax=self.ax)
        elif plot_type == "contour":
            if len(_y1_cols_0) < 2 or _x0_raw is None:
                QMessageBox.warning(
                    self,
                    "Warning",
                    "Contour plot requires at least 2 Y columns (Y-coord, Z-value).",
                )
                return
            x_num = pd.to_numeric(
                _x0_raw.astype(str).str.replace(r"[^\d.-]", "", regex=True),
                errors="coerce",
            )
            y_num = _to_numeric_series(_y1_cols_0[0])
            z_num = _to_numeric_series(_y1_cols_0[1])
            valid = pd.DataFrame({"x": x_num, "y": y_num, "z": z_num}).dropna()
            if (
                len(valid) >= 4
                and valid["x"].nunique() > 1
                and valid["y"].nunique() > 1
            ):
                xi = np.linspace(valid["x"].min(), valid["x"].max(), 50)
                yi = np.linspace(valid["y"].min(), valid["y"].max(), 50)
                Xi, Yi = np.meshgrid(xi, yi)
                try:
                    Zi = griddata(
                        (valid["x"].values, valid["y"].values),
                        valid["z"].values,
                        (Xi, Yi),
                        method="linear",
                    )
                except Exception:
                    Zi = griddata(
                        (valid["x"].values, valid["y"].values),
                        valid["z"].values,
                        (Xi, Yi),
                        method="nearest",
                    )
                cf = self.ax.contourf(
                    Xi, Yi, Zi, levels=15, cmap=self.colormap_combo.currentText()
                )
                self.fig.colorbar(cf, ax=self.ax)
            _finish_special()
            return

        elif plot_type == "polar":
            self.fig.clear()
            self.ax = self.fig.add_subplot(111, projection="polar")
            x_num = (
                pd.to_numeric(
                    _x0_raw.astype(str).str.replace(r"[^\d.-]", "", regex=True),
                    errors="coerce",
                )
                if _x0_raw is not None
                else pd.Series()
            )
            for yc in _y1_cols_0:
                valid = pd.DataFrame({"x": x_num, "y": _to_numeric_series(yc)}).dropna()
                if not valid.empty:
                    self.ax.plot(
                        np.radians(valid["x"].values),
                        valid["y"].values,
                        label=yc,
                        marker="o",
                    )
            if self.legend_show_check.isChecked():
                self.ax.legend()
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
            prefix = f"T{tab_idx + 1}-" if len(x_tabs_info) > 1 else ""

            for y_col in tab_data["y1_cols"]:
                if y_col in plot_df.columns:
                    plot_series(
                        self.ax, x_col, y_col, is_twin_ax=False, label_prefix=prefix
                    )

            target_ax = self.ax2 if self.ax2 else self.ax
            for y_col in tab_data["y2_cols"]:
                if y_col in plot_df.columns:
                    plot_series(
                        target_ax, x_col, y_col, is_twin_ax=True, label_prefix=prefix
                    )

        # Titles and Formatting
        font_family = self.font_family_combo.currentText()
        first_x_col = x_tabs_info[0]["x_axis"] if x_tabs_info else ""
        y1_names = [c for tab in x_tabs_info for c in tab["y1_cols"]]

        self.ax.set_title(
            self.title_input.text(),
            fontsize=self.title_fontsize_spin.value(),
            fontfamily=font_family,
        )
        self.ax.set_xlabel(
            self.xlabel_input.text() if self.xlabel_input.text() else first_x_col,
            fontsize=self.xlabel_fontsize_spin.value(),
            fontfamily=font_family,
        )
        self.ax.set_ylabel(
            self.ylabel_input.text()
            if self.ylabel_input.text()
            else ", ".join(y1_names),
            fontsize=self.ylabel_fontsize_spin.value(),
            fontfamily=font_family,
        )
        if self.ax2:
            y2_names = [c for tab in x_tabs_info for c in tab["y2_cols"]]
            self.ax2.set_ylabel(
                self.ylabel2_input.text()
                if self.ylabel2_input.text()
                else ", ".join(y2_names),
                fontsize=self.ylabel2_fontsize_spin.value(),
                fontfamily=font_family,
            )

        self.set_axis_limits(
            self.ax, "x", self.xlim_min_input.text(), self.xlim_max_input.text()
        )
        self.set_axis_limits(
            self.ax, "y", self.ylim_min_input.text(), self.ylim_max_input.text()
        )
        if self.ax2:
            self.set_axis_limits(
                self.ax2, "y", self.ylim2_min_input.text(), self.ylim2_max_input.text()
            )

        # Log scale / axis inversion (categorical axes, e.g. bar charts,
        # can't be log-scaled — skip rather than crash the plot)
        try:
            if self.x_log_check.isChecked():
                self.ax.set_xscale("log")
            if self.y1_log_check.isChecked():
                self.ax.set_yscale("log")
        except ValueError:
            pass
        if self.y1_invert_check.isChecked():
            self.ax.invert_yaxis()
        if self.ax2:
            try:
                if self.y2_log_check.isChecked():
                    self.ax2.set_yscale("log")
            except ValueError:
                pass
            if self.y2_invert_check.isChecked():
                self.ax2.invert_yaxis()

        # Plain (non-scientific) number formatting. ticklabel_format only
        # supports the default ScalarFormatter, so log-scale/categorical
        # axes (which use a different formatter) are skipped silently.
        if self.xaxis_plain_check.isChecked():
            try:
                self.ax.ticklabel_format(axis="x", style="plain", useOffset=False)
            except (ValueError, AttributeError):
                pass
        if self.yaxis1_plain_check.isChecked():
            try:
                self.ax.ticklabel_format(axis="y", style="plain", useOffset=False)
            except (ValueError, AttributeError):
                pass
        if self.ax2 and self.yaxis2_plain_check.isChecked():
            try:
                self.ax2.ticklabel_format(axis="y", style="plain", useOffset=False)
            except (ValueError, AttributeError):
                pass

        # Spine visibility
        self.ax.spines["top"].set_visible(self.spine_top_check.isChecked())
        self.ax.spines["bottom"].set_visible(self.spine_bottom_check.isChecked())
        self.ax.spines["left"].set_visible(self.spine_left_check.isChecked())
        self.ax.spines["right"].set_visible(self.spine_right_check.isChecked())

        self.ax.tick_params(labelsize=self.tick_fontsize_spin.value())
        if self.ax2:
            self.ax2.tick_params(labelsize=self.tick2_fontsize_spin.value())

        apply_major_ticker(
            self.ax.xaxis,
            self.xtick_major_interval_input.text(),
            self.x_log_check.isChecked(),
        )
        apply_major_ticker(
            self.ax.yaxis,
            self.ytick_major_interval_input.text(),
            self.y1_log_check.isChecked(),
        )
        if self.ax2:
            apply_major_ticker(
                self.ax2.yaxis,
                self.ytick2_major_interval_input.text(),
                self.y2_log_check.isChecked(),
            )

        apply_minor_ticker(
            self.ax.xaxis,
            self.xtick_minor_check.isChecked(),
            self.xtick_minor_interval_input.text(),
            self.x_log_check.isChecked(),
        )
        apply_minor_ticker(
            self.ax.yaxis,
            self.ytick_minor_check.isChecked(),
            self.ytick_minor_interval_input.text(),
            self.y1_log_check.isChecked(),
        )
        if self.ax2:
            apply_minor_ticker(
                self.ax2.yaxis,
                self.ytick2_minor_check.isChecked(),
                self.ytick2_minor_interval_input.text(),
                self.y2_log_check.isChecked(),
            )

        if self.grid_check.isChecked():
            self.ax.grid(
                True,
                alpha=self.grid_alpha_spin.value(),
                linestyle=self.grid_linestyle_combo.currentText(),
                linewidth=self.grid_linewidth_spin.value(),
            )

        if self.rotate_labels_check.isChecked():
            self.ax.tick_params(axis="x", rotation=self.rotation_angle_spin.value())

        if self.legend_show_check.isChecked():
            h1, l1 = self.ax.get_legend_handles_labels()
            h2, l2 = self.ax2.get_legend_handles_labels() if self.ax2 else ([], [])
            self.ax.legend(
                h1 + h2,
                l1 + l2,
                loc=self.legend_loc_combo.currentText(),
                fontsize=self.legend_fontsize_spin.value(),
            )

        self.fig.tight_layout()
        self.canvas.draw()

    def set_axis_limits(self, ax, axis, min_val_str, max_val_str):
        try:
            val_min = float(min_val_str) if min_val_str else None
            val_max = float(max_val_str) if max_val_str else None
            if axis == "x":
                if val_min is not None and val_max is not None:
                    ax.set_xlim(val_min, val_max)
                elif val_min is not None:
                    ax.set_xlim(left=val_min)
                elif val_max is not None:
                    ax.set_xlim(right=val_max)
            elif axis == "y":
                if val_min is not None and val_max is not None:
                    ax.set_ylim(val_min, val_max)
                elif val_min is not None:
                    ax.set_ylim(bottom=val_min)
                elif val_max is not None:
                    ax.set_ylim(top=val_max)
        except ValueError:
            pass

    # ── Import / Export & Project Files ──────────────────────────────────────
    def save_settings(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Project", "", "Matplotlib Graph Project (*.pmggrp)"
        )
        if file_path:
            save_project_file(self, file_path, version_str=VERSION, dimension="2D")
            self.current_project_path = file_path
            self.undo_stack.setClean()
            self.update_window_title()

    def overwrite_save(self):
        if not self.current_project_path:
            self.save_settings()
        else:
            save_project_file(
                self, self.current_project_path, version_str=VERSION, dimension="2D"
            )
            self.undo_stack.setClean()
            self.update_window_title()

    def load_settings(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Project", "", "Matplotlib Graph Project (*.pmggrp)"
        )
        if file_path:
            self.load_project_file(file_path)

    def load_project_file(self, file_path):
        load_project_file(self, file_path)
        self.current_project_path = file_path
        self.undo_stack.setClean()
        self.update_window_title()

    def export_graph(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Plot Image",
            "",
            "PNG Image (*.png);;PDF Document (*.pdf);;SVG Image (*.svg)",
        )
        if file_path:
            self.fig.savefig(
                file_path, dpi=self.export_dpi_spin.value(), bbox_inches="tight"
            )
            QMessageBox.information(self, "Success", f"Plot saved to {file_path}")

    def export_filtered_data(self):
        if self.df is None:
            return
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export Data", "", "CSV File (*.csv);;Excel File (*.xlsx)"
        )
        if file_path:
            df_export = self.data_mgr.get_filtered_df(
                filter_enabled=self.data_filter_check.isChecked(),
                filter_column=self.filter_column_combo.currentText(),
                min_val_str=self.filter_min_input.text(),
                max_val_str=self.filter_max_input.text(),
            )
            if file_path.endswith(".xlsx"):
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
        self._table_snapshot = {}
        self.undo_stack.clear()

        for info in self.x_tab_widgets:
            info["x_combo"].clear()
            info["y1_listbox"].clear()
            info["y2_listbox"].clear()
        self.combined_style_target_combo.clear()
        self.errorbar_column_combo.clear()
        self.filter_column_combo.clear()

        self.fig.clear()
        self.ax = self.fig.add_subplot(111)
        self.canvas.draw()

    def reset_settings(self):
        reset_to_defaults(self, dimension="2D")
        self.y1_series_styles = {}
        self.y2_series_styles = {}
        self.clear_all()

    def open_in_3d_mode(self):
        try:
            from hygrapher.main_3d import GraphApp3D

            self.win_3d = GraphApp3D()
            if self.df is not None:
                self.win_3d.data_mgr.set_dataframe(self.df.copy())
                self.win_3d.df = self.win_3d.data_mgr.get_filtered_df()
                self.win_3d.populate_data_table()
                self.win_3d.update_combos()
            self.win_3d.show()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open 3D Mode:\n{e}")

    def closeEvent(self, event):
        if not self.undo_stack.isClean():
            reply = QMessageBox.question(
                self,
                "Unsaved Changes",
                "You have unsaved changes to the data table. Save before closing?",
                QMessageBox.StandardButton.Save
                | QMessageBox.StandardButton.Discard
                | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Save,
            )
            if reply == QMessageBox.StandardButton.Cancel:
                event.ignore()
                return
            if reply == QMessageBox.StandardButton.Save:
                self.overwrite_save()
                if not self.undo_stack.isClean():
                    # Save As was cancelled by the user — don't close.
                    event.ignore()
                    return
        event.accept()

    def update_window_title(self):
        base = f"HyGrapher v{VERSION}"
        if self.current_project_path:
            name = os.path.basename(self.current_project_path)
            base = f"{name} - {base}"
        if not self.undo_stack.isClean():
            base = f"*{base}"
        self.setWindowTitle(base)

    def show_about(self):
        QMessageBox.about(
            self,
            "About HyGrapher",
            f"HyGrapher v{VERSION}\nA cross-platform Matplotlib GUI application built with PyQt6.",
        )


def main():
    app = QApplication(sys.argv)
    window = GraphApp()
    window.show()

    cli_file = resolve_cli_file(sys.argv[1:])
    if cli_file:
        window.load_data(cli_file)
        if window.df is not None:
            window.plot_graph()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
