#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
hygrapher.main

Author: Hiromichi Yokoyama
License: Apache-2.0 license
Repo: https://github.com/HiroYokoyama/matplotlib_graph_app

Main 2D application module with cross-platform scrolling, Multi X-Axis Tabs,
extended graph settings, non-destructive data filtering, and headless support.
"""

VERSION = "0.6.0"

import os
import sys
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import pandas as pd
import numpy as np
import matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.ticker as ticker
from scipy.interpolate import griddata
from tksheet import Sheet

# Import modular components
from .utils import (
    bind_scroll_events,
    bind_mousewheel_recursive,
    get_font_list,
    apply_major_ticker,
    apply_minor_ticker,
)
from .data_manager import DataManager
from .project_io import build_project_dict, save_project_file

# Try to import tkinterdnd2 for drag and drop support
try:
    from tkinterdnd2 import TkinterDnD
    BASE_CLASS = TkinterDnD.Tk
    DND_AVAILABLE = True
except ImportError:
    BASE_CLASS = tk.Tk
    DND_AVAILABLE = False


class GraphApp(BASE_CLASS):
    def __init__(self):
        super().__init__()
        self.title(f"HYGrapher ver. {VERSION}")
        self.geometry("1600x900")

        self.data_mgr = DataManager()
        self.sheet = None
        self.current_project_path = ""

        # Enable drag and drop
        self.setup_drag_and_drop()

        # Get system font list
        self.font_list = get_font_list()

        # Create all tk.Variables
        self.create_all_tk_variables()

        # X-Tabs metadata container
        self.x_tab_widgets = []  # List of dicts with tab widget references

        # Figure & Canvas setup
        self.fig = Figure(figsize=(self.fig_width_var.get(), self.fig_height_var.get()), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax2 = None

        # Menu Bar
        self.setup_menu_bar()

        # Main Layout
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Top Frame (File Operations)
        self.setup_top_frame(main_frame)

        # Content Split PanedWindow
        content_frame = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Left Panel (Settings & Data Editor)
        left_container = ttk.Frame(content_frame)
        content_frame.add(left_container, weight=1)

        left_paned = ttk.PanedWindow(left_container, orient=tk.VERTICAL)
        left_paned.pack(fill=tk.BOTH, expand=True)

        settings_container = ttk.Frame(left_paned)
        left_paned.add(settings_container, weight=1)

        # Plot Button
        self.plot_button = ttk.Button(settings_container, text="Plot / Update Graph", command=self.plot_graph)
        self.plot_button.pack(side=tk.BOTTOM, fill=tk.X, pady=5, padx=5)
        self.plot_button['state'] = 'disabled'

        # Notebook for Settings Tabs
        self.settings_notebook = ttk.Notebook(settings_container)
        self.settings_notebook.pack(fill=tk.BOTH, expand=True, pady=2)

        # Build Notebook Tabs
        self.create_scrollable_tab(self.settings_notebook, "Basic Settings", self.create_basic_settings_tab)
        self.create_scrollable_tab(self.settings_notebook, "Style", self.create_style_settings_tab)
        self.create_scrollable_tab(self.settings_notebook, "Font", self.create_font_size_tab)
        self.create_scrollable_tab(self.settings_notebook, "Axis/Ticks", self.create_axis_ticks_tab)
        self.create_scrollable_tab(self.settings_notebook, "Spines/BG", self.create_spines_tab)
        self.create_scrollable_tab(self.settings_notebook, "Legend", self.create_legend_tab)
        self.create_scrollable_tab(self.settings_notebook, "Advanced", self.create_advanced_tab)

        # Bottom section: Data Editor
        data_frame = ttk.LabelFrame(left_paned, text="Data Editor")
        left_paned.add(data_frame, weight=3)

        self.sheet_frame = ttk.Frame(data_frame)
        self.sheet_frame.pack(fill=tk.BOTH, expand=True)

        # Right Panel (Graph Preview)
        right_panel = ttk.LabelFrame(content_frame, text="Graph Preview")
        content_frame.add(right_panel, weight=1)

        def force_sash_position(event=None):
            try:
                width = content_frame.winfo_width()
                content_frame.sashpos(0, width // 2)
                content_frame.unbind("<Configure>")
            except Exception:
                pass

        content_frame.bind("<Configure>", force_sash_position)

        # Graph Canvas & Scrollbars
        self.scrollable_canvas = tk.Canvas(right_panel, borderwidth=0, highlightthickness=0)
        v_scroll = ttk.Scrollbar(right_panel, orient=tk.VERTICAL, command=self.scrollable_canvas.yview)
        h_scroll = ttk.Scrollbar(right_panel, orient=tk.HORIZONTAL, command=self.scrollable_canvas.xview)
        self.scrollable_canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.scrollable_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.graph_frame = ttk.Frame(self.scrollable_canvas)
        self.scrollable_canvas.create_window((0, 0), window=self.graph_frame, anchor="nw")

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_frame)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.graph_frame)
        self.toolbar.update()
        self.toolbar.pack(side=tk.TOP, fill=tk.X)
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.NONE, expand=False, padx=5, pady=5)

        self.graph_frame.bind("<Configure>", self.on_graph_frame_configure)

    @property
    def df(self):
        return self.data_mgr.raw_df

    @df.setter
    def df(self, dataframe):
        self.data_mgr.raw_df = dataframe

    @property
    def data_file_path(self):
        return self.data_mgr.file_path

    @data_file_path.setter
    def data_file_path(self, path):
        self.data_mgr.file_path = path

    def on_graph_frame_configure(self, event=None):
        self.scrollable_canvas.configure(scrollregion=self.scrollable_canvas.bbox("all"))

    def create_all_tk_variables(self):
        # Basic
        self.plot_type_var = tk.StringVar(value="line")
        self.x_axis_var = tk.StringVar()
        self.title_var = tk.StringVar()
        self.xlabel_var = tk.StringVar()
        self.ylabel_var = tk.StringVar()
        self.ylabel2_var = tk.StringVar()

        # Series styles dicts
        self.y1_series_styles = {}
        self.y2_series_styles = {}

        self.combined_style_target_var = tk.StringVar()

        # Transient Style Editor variables
        self.current_style_color_var = tk.StringVar(value="#000000")
        self.current_style_linestyle_var = tk.StringVar(value="-")
        self.current_style_marker_var = tk.StringVar(value="o")
        self.current_style_linewidth_var = tk.DoubleVar(value=1.5)
        self.current_style_markersize_var = tk.DoubleVar(value=6.0)
        self.current_style_alpha_var = tk.DoubleVar(value=1.0)

        self.grid_var = tk.BooleanVar(value=False)
        self.marker_var = tk.BooleanVar(value=True)

        # Font & Size
        self.font_family_var = tk.StringVar(value=self.font_list[0] if self.font_list else 'sans-serif')
        self.title_fontsize_var = tk.DoubleVar(value=16.0)
        self.xlabel_fontsize_var = tk.DoubleVar(value=14.0)
        self.ylabel_fontsize_var = tk.DoubleVar(value=14.0)
        self.ylabel2_fontsize_var = tk.DoubleVar(value=14.0)
        self.tick_fontsize_var = tk.DoubleVar(value=14.0)
        self.tick2_fontsize_var = tk.DoubleVar(value=14.0)
        self.legend_fontsize_var = tk.DoubleVar(value=12.0)
        self.fig_width_var = tk.DoubleVar(value=7.0)
        self.fig_height_var = tk.DoubleVar(value=6.0)

        # Axis & Ticks
        self.xlim_min_var = tk.StringVar()
        self.xlim_max_var = tk.StringVar()
        self.ylim_min_var = tk.StringVar()
        self.ylim_max_var = tk.StringVar()
        self.ylim2_min_var = tk.StringVar()
        self.ylim2_max_var = tk.StringVar()

        self.xtick_show_var = tk.BooleanVar(value=True)
        self.xtick_label_show_var = tk.BooleanVar(value=True)
        self.xtick_direction_var = tk.StringVar(value='out')
        self.ytick_show_var = tk.BooleanVar(value=True)
        self.ytick_label_show_var = tk.BooleanVar(value=True)
        self.ytick_direction_var = tk.StringVar(value='out')
        self.ytick2_show_var = tk.BooleanVar(value=True)
        self.ytick2_label_show_var = tk.BooleanVar(value=True)
        self.ytick2_direction_var = tk.StringVar(value='out')

        # Minor Ticks
        self.xtick_minor_show_var = tk.BooleanVar(value=False)
        self.ytick_minor_show_var = tk.BooleanVar(value=False)
        self.ytick2_minor_show_var = tk.BooleanVar(value=False)
        self.xtick_minor_interval_var = tk.StringVar()
        self.ytick_minor_interval_var = tk.StringVar()
        self.ytick2_minor_interval_var = tk.StringVar()

        self.xaxis_plain_format_var = tk.BooleanVar(value=False)
        self.yaxis1_plain_format_var = tk.BooleanVar(value=False)
        self.yaxis2_plain_format_var = tk.BooleanVar(value=False)

        self.xtick_major_interval_var = tk.StringVar()
        self.ytick_major_interval_var = tk.StringVar()
        self.ytick2_major_interval_var = tk.StringVar()

        # Spines / BG
        self.spine_top_var = tk.BooleanVar(value=True)
        self.spine_bottom_var = tk.BooleanVar(value=True)
        self.spine_left_var = tk.BooleanVar(value=True)
        self.spine_right_var = tk.BooleanVar(value=True)
        self.face_color_var = tk.StringVar(value='#FFFFFF')
        self.fig_color_var = tk.StringVar(value='#FFFFFF')

        # Legend
        self.legend_show_var = tk.BooleanVar(value=False)
        self.legend_loc_var = tk.StringVar(value='best')

        # Log Scale & Invert
        self.x_log_scale_var = tk.BooleanVar(value=False)
        self.y1_log_scale_var = tk.BooleanVar(value=False)
        self.y2_log_scale_var = tk.BooleanVar(value=False)

        self.x_invert_var = tk.BooleanVar(value=False)
        self.y1_invert_var = tk.BooleanVar(value=False)
        self.y2_invert_var = tk.BooleanVar(value=False)

        # Advanced
        self.enable_smoothing_var = tk.BooleanVar(value=False)
        self.smoothing_window_var = tk.IntVar(value=5)
        self.enable_errorbar_var = tk.BooleanVar(value=False)
        self.errorbar_column_var = tk.StringVar()
        self.enable_annotation_var = tk.BooleanVar(value=False)

        self.data_filter_enabled_var = tk.BooleanVar(value=False)
        self.filter_min_var = tk.StringVar()
        self.filter_max_var = tk.StringVar()
        self.filter_column_var = tk.StringVar()
        self.grid_alpha_var = tk.DoubleVar(value=0.3)
        self.grid_linestyle_var = tk.StringVar(value='--')
        self.grid_linewidth_var = tk.DoubleVar(value=0.5)
        self.subplot_mode_var = tk.BooleanVar(value=False)
        self.rotate_labels_var = tk.BooleanVar(value=False)
        self.rotation_angle_var = tk.IntVar(value=45)

        # Extended settings
        self.export_dpi_var = tk.IntVar(value=150)
        self.colormap_var = tk.StringVar(value="viridis")
        self.tight_layout_pad_var = tk.DoubleVar(value=1.0)

    def setup_menu_bar(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Load Data (CSV/Excel)...", command=self.load_data)
        file_menu.add_separator()
        file_menu.add_command(label="Open Project...", command=self.load_settings, accelerator="Ctrl+O")
        file_menu.add_command(label="Save Project", command=self.overwrite_save, accelerator="Ctrl+S")
        file_menu.add_command(label="Save Project As...", command=self.save_settings)
        file_menu.add_separator()
        file_menu.add_command(label="Export Graph...", command=self.export_graph)
        file_menu.add_command(label="Export Data (CSV)...", command=self.export_filtered_data)
        file_menu.add_separator()
        file_menu.add_command(label="Open in 3D Mode...", command=self.open_in_3d_mode)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)

        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Clear All", command=self.clear_all)
        edit_menu.add_command(label="Reset Settings", command=self.reset_settings)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)

    def setup_top_frame(self, parent):
        top_frame = ttk.Frame(parent)
        top_frame.pack(fill=tk.X, pady=(0, 5))

        data_ops_frame = ttk.LabelFrame(top_frame, text="Data")
        data_ops_frame.pack(side=tk.LEFT, padx=(0, 5))
        self.load_button = ttk.Button(data_ops_frame, text="Load Data", command=self.load_data, width=12)
        self.load_button.pack(side=tk.LEFT, padx=2, pady=2)
        self.clear_button = ttk.Button(data_ops_frame, text="Clear All", command=self.clear_all, width=12)
        self.clear_button.pack(side=tk.LEFT, padx=2, pady=2)

        project_ops_frame = ttk.LabelFrame(top_frame, text="Project")
        project_ops_frame.pack(side=tk.LEFT, padx=5)
        self.load_settings_button = ttk.Button(project_ops_frame, text="Open", command=self.load_settings, width=10)
        self.load_settings_button.pack(side=tk.LEFT, padx=2, pady=2)
        self.overwrite_save_button = ttk.Button(project_ops_frame, text="Save", command=self.overwrite_save, width=10)
        self.overwrite_save_button.pack(side=tk.LEFT, padx=2, pady=2)
        self.overwrite_save_button['state'] = 'disabled'
        self.save_settings_button = ttk.Button(project_ops_frame, text="Save As...", command=self.save_settings, width=10)
        self.save_settings_button.pack(side=tk.LEFT, padx=2, pady=2)
        self.reset_settings_button = ttk.Button(project_ops_frame, text="Reset", command=self.reset_settings, width=10)
        self.reset_settings_button.pack(side=tk.LEFT, padx=2, pady=2)

        export_ops_frame = ttk.LabelFrame(top_frame, text="Export")
        export_ops_frame.pack(side=tk.LEFT, padx=(5, 0))
        self.export_button = ttk.Button(export_ops_frame, text="Graph", command=self.export_graph, width=10)
        self.export_button.pack(side=tk.LEFT, padx=2, pady=2)
        self.export_button['state'] = 'disabled'
        self.export_data_button = ttk.Button(export_ops_frame, text="Data", command=self.export_filtered_data, width=10)
        self.export_data_button.pack(side=tk.LEFT, padx=2, pady=2)
        self.export_data_button['state'] = 'disabled'

        self.bind('<Control-s>', lambda e: self.overwrite_save())
        self.bind('<Control-o>', lambda e: self.load_settings())

    def create_scrollable_tab(self, parent_notebook, tab_name, create_content_func):
        tab_container = ttk.Frame(parent_notebook)
        parent_notebook.add(tab_container, text=tab_name)

        tab_canvas = tk.Canvas(tab_container, borderwidth=0, highlightthickness=0)
        tab_scrollbar = ttk.Scrollbar(tab_container, orient=tk.VERTICAL, command=tab_canvas.yview)
        tab_canvas.configure(yscrollcommand=tab_scrollbar.set)

        tab_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tab_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        content_frame = ttk.Frame(tab_canvas, padding=5)
        canvas_window = tab_canvas.create_window((0, 0), window=content_frame, anchor="nw")

        def on_content_configure(event):
            tab_canvas.configure(scrollregion=tab_canvas.bbox("all"))

        content_frame.bind("<Configure>", on_content_configure)

        def on_canvas_configure(event):
            tab_canvas.itemconfig(canvas_window, width=event.width)

        tab_canvas.bind("<Configure>", on_canvas_configure)

        # Cross-platform scroll handler
        def _scroll_cb(scroll_units, event):
            tab_canvas.yview_scroll(scroll_units, "units")

        bind_scroll_events(tab_canvas, _scroll_cb)

        create_content_func(content_frame)

        bind_mousewheel_recursive(content_frame, _scroll_cb)
        return content_frame

    # --- Tab 1: Basic Settings (Multi X-Axis Support) ---
    def create_basic_settings_tab(self, frame):
        top_settings_frame = ttk.Frame(frame)
        top_settings_frame.pack(fill=tk.X, pady=2)

        ttk.Label(top_settings_frame, text="Graph Title:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.title_entry = ttk.Entry(top_settings_frame, textvariable=self.title_var, width=25)
        self.title_entry.grid(row=0, column=1, columnspan=3, padx=3, pady=2, sticky=tk.EW)

        ttk.Label(top_settings_frame, text="Plot Type:").grid(row=1, column=0, padx=3, pady=2, sticky=tk.W)
        self.plot_type_combo = ttk.Combobox(
            top_settings_frame,
            textvariable=self.plot_type_var,
            values=["line", "scatter", "bar", "step", "stem", "area", "pie", "box", "violin", "heatmap", "contour", "polar"],
            state='readonly', width=24
        )
        self.plot_type_combo.grid(row=1, column=1, columnspan=3, padx=3, pady=2, sticky=tk.EW)

        top_settings_frame.columnconfigure(1, weight=1)

        # Multi X-Axis Tabs Container
        x_tabs_header = ttk.Frame(frame)
        x_tabs_header.pack(fill=tk.X, pady=(5, 2))

        ttk.Label(x_tabs_header, text="Data Series Configuration", font=("-weight bold")).pack(side=tk.LEFT, padx=3)
        add_tab_btn = ttk.Button(x_tabs_header, text="+ Add X-Tab", command=self.add_x_tab, width=12)
        add_tab_btn.pack(side=tk.RIGHT, padx=2)

        self.x_notebook = ttk.Notebook(frame)
        self.x_notebook.pack(fill=tk.BOTH, expand=True, pady=2)

        # Add initial primary X-Tab
        self.add_x_tab(is_initial=True)

        # Figure Size
        fig_size_frame = ttk.Frame(frame)
        fig_size_frame.pack(fill=tk.X, pady=(5, 0))

        ttk.Label(fig_size_frame, text="Figure Width (inch):").grid(row=0, column=0, padx=2, pady=1, sticky=tk.W)
        self.fig_width_spin = ttk.Spinbox(fig_size_frame, from_=3, to=20, increment=0.5, textvariable=self.fig_width_var, width=8)
        self.fig_width_spin.grid(row=0, column=1, padx=2, pady=1, sticky=tk.W)

        ttk.Label(fig_size_frame, text="Figure Height (inch):").grid(row=0, column=2, padx=5, pady=1, sticky=tk.W)
        self.fig_height_spin = ttk.Spinbox(fig_size_frame, from_=3, to=20, increment=0.5, textvariable=self.fig_height_var, width=8)
        self.fig_height_spin.grid(row=0, column=3, padx=2, pady=1, sticky=tk.W)

    def add_x_tab(self, is_initial=False):
        tab_index = len(self.x_tab_widgets) + 1
        tab_title = f"X-Tab {tab_index}"

        tab_frame = ttk.Frame(self.x_notebook, padding=3)
        self.x_notebook.add(tab_frame, text=tab_title)

        # Controls header inside tab
        ctrl_bar = ttk.Frame(tab_frame)
        ctrl_bar.pack(fill=tk.X, pady=2)

        if not is_initial:
            remove_btn = ttk.Button(ctrl_bar, text="✕ Remove Tab", command=lambda tf=tab_frame: self.remove_x_tab(tf), width=12)
            remove_btn.pack(side=tk.RIGHT, padx=2)

        # X-Axis Frame
        x_axis_frame = ttk.LabelFrame(tab_frame, text="X-Axis Settings")
        x_axis_frame.pack(fill=tk.X, padx=2, pady=2)

        ttk.Label(x_axis_frame, text="Label:").grid(row=0, column=0, padx=2, pady=1, sticky=tk.W)
        xlabel_entry = ttk.Entry(x_axis_frame, width=22)
        xlabel_entry.grid(row=0, column=1, columnspan=2, padx=2, pady=1, sticky=tk.EW)
        if is_initial:
            xlabel_entry.config(textvariable=self.xlabel_var)

        ttk.Label(x_axis_frame, text="Data:").grid(row=1, column=0, padx=2, pady=1, sticky=tk.W)
        x_combo = ttk.Combobox(x_axis_frame, state='disabled', width=20)
        x_combo.grid(row=1, column=1, columnspan=2, padx=2, pady=1, sticky=tk.EW)
        if is_initial:
            x_combo.config(textvariable=self.x_axis_var)

        x_axis_frame.columnconfigure(1, weight=1)

        # Y-Axes split
        y_paned = ttk.PanedWindow(tab_frame, orient=tk.HORIZONTAL)
        y_paned.pack(fill=tk.X, expand=False, padx=2, pady=2)

        # Y1
        y1_frame = ttk.LabelFrame(y_paned, text="Y-Axis (Left)")
        y_paned.add(y1_frame, weight=1)

        ttk.Label(y1_frame, text="Label:").grid(row=0, column=0, padx=2, pady=1, sticky=tk.W)
        y1_label_entry = ttk.Entry(y1_frame, width=18)
        y1_label_entry.grid(row=0, column=1, padx=2, pady=1, sticky=tk.EW)
        if is_initial:
            y1_label_entry.config(textvariable=self.ylabel_var)

        ttk.Label(y1_frame, text="Data (Multi-select):").grid(row=1, column=0, columnspan=2, padx=2, pady=1, sticky=tk.W)
        y1_list_frame = ttk.Frame(y1_frame, height=75)
        y1_scroll = ttk.Scrollbar(y1_list_frame, orient=tk.VERTICAL)
        y1_listbox = tk.Listbox(y1_list_frame, selectmode=tk.MULTIPLE, yscrollcommand=y1_scroll.set, exportselection=False, height=4)
        y1_scroll.config(command=y1_listbox.yview)
        y1_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        y1_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y1_list_frame.grid(row=2, column=0, columnspan=2, padx=2, pady=1, sticky=tk.EW)
        y1_list_frame.pack_propagate(False)
        y1_frame.columnconfigure(1, weight=1)

        # Y2
        y2_frame = ttk.LabelFrame(y_paned, text="Y-Axis (Right)")
        y_paned.add(y2_frame, weight=1)

        ttk.Label(y2_frame, text="Label:").grid(row=0, column=0, padx=2, pady=1, sticky=tk.W)
        y2_label_entry = ttk.Entry(y2_frame, width=18)
        y2_label_entry.grid(row=0, column=1, padx=2, pady=1, sticky=tk.EW)
        if is_initial:
            y2_label_entry.config(textvariable=self.ylabel2_var)

        ttk.Label(y2_frame, text="Data (Multi-select):").grid(row=1, column=0, columnspan=2, padx=2, pady=1, sticky=tk.W)
        y2_list_frame = ttk.Frame(y2_frame, height=75)
        y2_scroll = ttk.Scrollbar(y2_list_frame, orient=tk.VERTICAL)
        y2_listbox = tk.Listbox(y2_list_frame, selectmode=tk.MULTIPLE, yscrollcommand=y2_scroll.set, exportselection=False, height=4)
        y2_scroll.config(command=y2_listbox.yview)
        y2_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        y2_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y2_list_frame.grid(row=2, column=0, columnspan=2, padx=2, pady=1, sticky=tk.EW)
        y2_list_frame.pack_propagate(False)
        y2_frame.columnconfigure(1, weight=1)

        tab_info = {
            "tab_frame": tab_frame,
            "x_combo": x_combo,
            "xlabel_entry": xlabel_entry,
            "y1_listbox": y1_listbox,
            "y1_label_entry": y1_label_entry,
            "y2_listbox": y2_listbox,
            "y2_label_entry": y2_label_entry,
        }

        if is_initial:
            self.y_listbox = y1_listbox
            self.y2_listbox = y2_listbox
            self.x_axis_combo = x_combo

        self.x_tab_widgets.append(tab_info)

        # Refresh combobox values if dataset is already loaded
        if self.df is not None:
            cols = self.df.columns.tolist()
            x_combo['values'] = cols
            x_combo['state'] = 'readonly'
            y1_listbox.delete(0, tk.END)
            y2_listbox.delete(0, tk.END)
            for c in cols:
                y1_listbox.insert(tk.END, c)
                y2_listbox.insert(tk.END, c)

    def remove_x_tab(self, tab_frame):
        for i, info in enumerate(self.x_tab_widgets):
            if info["tab_frame"] == tab_frame:
                self.x_notebook.forget(tab_frame)
                tab_frame.destroy()
                self.x_tab_widgets.pop(i)
                break

    def get_x_tabs_data(self):
        result = []
        for i, info in enumerate(self.x_tab_widgets):
            x_col = info["x_combo"].get()
            xlabel = info["xlabel_entry"].get()
            y1_cols = [info["y1_listbox"].get(idx) for idx in info["y1_listbox"].curselection()]
            y1_label = info["y1_label_entry"].get()
            y2_cols = [info["y2_listbox"].get(idx) for idx in info["y2_listbox"].curselection()]
            y2_label = info["y2_label_entry"].get()

            result.append({
                "tab_name": f"X-Tab {i+1}",
                "x_axis": x_col,
                "xlabel": xlabel,
                "y1_cols": y1_cols,
                "ylabel": y1_label,
                "y2_cols": y2_cols,
                "ylabel2": y2_label,
            })
        return result

    # --- Tab 2: Style Settings ---
    def create_style_settings_tab(self, frame):
        common_frame = ttk.LabelFrame(frame, text="Common Style Settings")
        common_frame.pack(fill=tk.X, padx=2, pady=2)

        self.grid_check = ttk.Checkbutton(common_frame, text="Show Grid", variable=self.grid_var)
        self.grid_check.grid(row=0, column=0, padx=2, pady=1, sticky=tk.W)

        self.marker_check = ttk.Checkbutton(common_frame, text="Show Markers", variable=self.marker_var)
        self.marker_check.grid(row=0, column=1, padx=5, pady=1, sticky=tk.W)

        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(5, 5))

        selector_frame = ttk.Frame(frame)
        selector_frame.pack(fill=tk.X, pady=2)

        ttk.Label(selector_frame, text="Select Series (All Tabs):").pack(side=tk.LEFT, padx=2, pady=1)
        self.style_combo = ttk.Combobox(selector_frame, textvariable=self.combined_style_target_var, state='readonly', width=25)
        self.style_combo.pack(side=tk.LEFT, padx=2, pady=1, fill=tk.X, expand=True)
        self.style_combo.bind("<<ComboboxSelected>>", self.on_combined_series_select)

        editor_frame = ttk.LabelFrame(frame, text="Style Editor (Selected Series)")
        editor_frame.pack(fill=tk.X, padx=2, pady=2)

        ttk.Label(editor_frame, text="Color:").grid(row=0, column=0, padx=2, pady=1, sticky=tk.W)
        self.style_editor_color_btn = ttk.Button(editor_frame, text="Select", command=self.on_style_editor_color_pick, width=8)
        self.style_editor_color_btn.grid(row=0, column=1, padx=1, pady=1)
        self.style_editor_color_auto_btn = ttk.Button(editor_frame, text="Auto", command=self.on_style_editor_color_auto, width=8)
        self.style_editor_color_auto_btn.grid(row=0, column=2, padx=1, pady=1)

        self.style_editor_color_label = ttk.Label(editor_frame, text="#000000", background="#000000", width=10, anchor=tk.CENTER)
        self.style_editor_color_label.grid(row=0, column=3, padx=2, pady=1)

        ttk.Label(editor_frame, text="Line Style:").grid(row=0, column=4, padx=5, pady=1, sticky=tk.W)
        self.style_editor_linestyle_combo = ttk.Combobox(
            editor_frame, textvariable=self.current_style_linestyle_var,
            values=['-', '--', ':', '-.', 'None'], state='readonly', width=6
        )
        self.style_editor_linestyle_combo.grid(row=0, column=5, padx=2, pady=1, sticky=tk.W)
        self.style_editor_linestyle_combo.bind("<<ComboboxSelected>>", self.on_style_editor_change)

        ttk.Label(editor_frame, text="Marker Style:").grid(row=1, column=0, padx=2, pady=1, sticky=tk.W)
        self.style_editor_marker_combo = ttk.Combobox(
            editor_frame, textvariable=self.current_style_marker_var,
            values=['o', '.', ',', 's', 'p', '*', '^', '<', '>', 'D', 'H', 'None'], state='readonly', width=6
        )
        self.style_editor_marker_combo.grid(row=1, column=1, columnspan=3, padx=2, pady=1, sticky=tk.W)
        self.style_editor_marker_combo.bind("<<ComboboxSelected>>", self.on_style_editor_change)

        ttk.Label(editor_frame, text="Line Width:").grid(row=1, column=4, padx=5, pady=1, sticky=tk.W)
        self.style_editor_linewidth_spin = ttk.Spinbox(
            editor_frame, from_=0.5, to=10.0, increment=0.5, textvariable=self.current_style_linewidth_var, width=6,
            command=self.on_style_editor_change
        )
        self.style_editor_linewidth_spin.grid(row=1, column=5, padx=2, pady=1, sticky=tk.W)
        self.style_editor_linewidth_spin.bind("<Return>", self.on_style_editor_change)

        ttk.Label(editor_frame, text="Marker Size:").grid(row=2, column=0, padx=2, pady=1, sticky=tk.W)
        self.style_editor_markersize_spin = ttk.Spinbox(
            editor_frame, from_=1.0, to=30.0, increment=1.0, textvariable=self.current_style_markersize_var, width=6,
            command=self.on_style_editor_change
        )
        self.style_editor_markersize_spin.grid(row=2, column=1, columnspan=3, padx=2, pady=1, sticky=tk.W)
        self.style_editor_markersize_spin.bind("<Return>", self.on_style_editor_change)

        ttk.Label(editor_frame, text="Alpha (Opacity):").grid(row=2, column=4, padx=5, pady=1, sticky=tk.W)
        self.style_editor_alpha_spin = ttk.Spinbox(
            editor_frame, from_=0.0, to=1.0, increment=0.1, textvariable=self.current_style_alpha_var, width=6,
            command=self.on_style_editor_change
        )
        self.style_editor_alpha_spin.grid(row=2, column=5, padx=2, pady=1, sticky=tk.W)
        self.style_editor_alpha_spin.bind("<Return>", self.on_style_editor_change)

    # --- Tab 3: Font Settings ---
    def create_font_size_tab(self, frame):
        ttk.Label(frame, text="Font Family:").grid(row=0, column=0, padx=2, pady=1, sticky=tk.W)
        self.font_family_combo = ttk.Combobox(frame, textvariable=self.font_family_var, values=self.font_list, state='readonly', width=18)
        self.font_family_combo.grid(row=0, column=1, columnspan=2, padx=2, pady=1, sticky=tk.W)

        ttk.Label(frame, text="Title Size:").grid(row=1, column=0, padx=2, pady=1, sticky=tk.W)
        self.title_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.title_fontsize_var, width=6)
        self.title_fontsize_spin.grid(row=1, column=1, padx=2, pady=1, sticky=tk.W)

        ttk.Label(frame, text="X-Label Size:").grid(row=2, column=0, padx=2, pady=1, sticky=tk.W)
        self.xlabel_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.xlabel_fontsize_var, width=6)
        self.xlabel_fontsize_spin.grid(row=2, column=1, padx=2, pady=1, sticky=tk.W)

        ttk.Label(frame, text="Y-Left Label Size:").grid(row=3, column=0, padx=2, pady=1, sticky=tk.W)
        self.ylabel_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.ylabel_fontsize_var, width=6)
        self.ylabel_fontsize_spin.grid(row=3, column=1, padx=2, pady=1, sticky=tk.W)

        ttk.Label(frame, text="Y-Right Label Size:").grid(row=4, column=0, padx=2, pady=1, sticky=tk.W)
        self.ylabel2_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.ylabel2_fontsize_var, width=6)
        self.ylabel2_fontsize_spin.grid(row=4, column=1, padx=2, pady=1, sticky=tk.W)

        ttk.Label(frame, text="Ticks (Left/X) Size:").grid(row=1, column=2, padx=5, pady=1, sticky=tk.W)
        self.tick_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.tick_fontsize_var, width=6)
        self.tick_fontsize_spin.grid(row=1, column=3, padx=2, pady=1, sticky=tk.W)

        ttk.Label(frame, text="Ticks (Right) Size:").grid(row=2, column=2, padx=5, pady=1, sticky=tk.W)
        self.tick2_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.tick2_fontsize_var, width=6)
        self.tick2_fontsize_spin.grid(row=2, column=3, padx=2, pady=1, sticky=tk.W)

        ttk.Label(frame, text="Legend Font Size:").grid(row=3, column=2, padx=5, pady=1, sticky=tk.W)
        self.legend_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.legend_fontsize_var, width=6)
        self.legend_fontsize_spin.grid(row=3, column=3, padx=2, pady=1, sticky=tk.W)

    # --- Tab 4: Axis & Ticks Settings ---
    def create_axis_ticks_tab(self, frame):
        # X-Axis
        ttk.Label(frame, text="X-Axis", font=("-weight bold")).grid(row=0, column=0, sticky=tk.W, pady=3)
        ttk.Label(frame, text="Range Min:").grid(row=1, column=0, padx=2, pady=2, sticky=tk.W)
        self.xlim_min_entry = ttk.Entry(frame, textvariable=self.xlim_min_var, width=10)
        self.xlim_min_entry.grid(row=1, column=1, padx=2, pady=2)

        ttk.Label(frame, text="Range Max:").grid(row=1, column=2, padx=2, pady=2, sticky=tk.W)
        self.xlim_max_entry = ttk.Entry(frame, textvariable=self.xlim_max_var, width=10)
        self.xlim_max_entry.grid(row=1, column=3, padx=2, pady=2)

        ttk.Label(frame, text="Major Interval:").grid(row=2, column=0, padx=2, pady=2, sticky=tk.W)
        self.xtick_major_interval_entry = ttk.Entry(frame, textvariable=self.xtick_major_interval_var, width=10)
        self.xtick_major_interval_entry.grid(row=2, column=1, padx=2, pady=2)

        ttk.Label(frame, text="Direction:").grid(row=2, column=2, padx=2, pady=2, sticky=tk.W)
        self.xtick_direction_combo = ttk.Combobox(frame, textvariable=self.xtick_direction_var, values=['out', 'in', 'inout'], state='readonly', width=8)
        self.xtick_direction_combo.grid(row=2, column=3, padx=2, pady=2, sticky=tk.W)

        self.xtick_show_check = ttk.Checkbutton(frame, text="Show Major Ticks", variable=self.xtick_show_var)
        self.xtick_show_check.grid(row=3, column=0, padx=2, pady=2, sticky=tk.W)
        self.xtick_label_show_check = ttk.Checkbutton(frame, text="Show Tick Labels", variable=self.xtick_label_show_var)
        self.xtick_label_show_check.grid(row=3, column=1, padx=2, pady=2, sticky=tk.W)

        self.xtick_minor_show_check = ttk.Checkbutton(frame, text="Minor Ticks", variable=self.xtick_minor_show_var)
        self.xtick_minor_show_check.grid(row=3, column=2, padx=2, pady=2, sticky=tk.W)
        self.xtick_minor_interval_entry = ttk.Entry(frame, textvariable=self.xtick_minor_interval_var, width=8)
        self.xtick_minor_interval_entry.grid(row=3, column=3, padx=2, pady=2, sticky=tk.W)

        self.x_log_scale_check = ttk.Checkbutton(frame, text="Log Scale", variable=self.x_log_scale_var)
        self.x_log_scale_check.grid(row=4, column=0, padx=2, pady=2, sticky=tk.W)
        self.x_invert_check = ttk.Checkbutton(frame, text="Invert Axis", variable=self.x_invert_var)
        self.x_invert_check.grid(row=4, column=1, padx=2, pady=2, sticky=tk.W)
        self.xaxis_plain_format_check = ttk.Checkbutton(frame, text="Disable Sci Notation", variable=self.xaxis_plain_format_var)
        self.xaxis_plain_format_check.grid(row=4, column=2, columnspan=2, padx=2, pady=2, sticky=tk.W)

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=5, column=0, columnspan=4, sticky="ew", pady=5)

        # Y-Axis (Left)
        ttk.Label(frame, text="Y-Axis (Left)", font=("-weight bold")).grid(row=6, column=0, sticky=tk.W, pady=3)
        ttk.Label(frame, text="Range Min:").grid(row=7, column=0, padx=2, pady=2, sticky=tk.W)
        self.ylim_min_entry = ttk.Entry(frame, textvariable=self.ylim_min_var, width=10)
        self.ylim_min_entry.grid(row=7, column=1, padx=2, pady=2)

        ttk.Label(frame, text="Range Max:").grid(row=7, column=2, padx=2, pady=2, sticky=tk.W)
        self.ylim_max_entry = ttk.Entry(frame, textvariable=self.ylim_max_var, width=10)
        self.ylim_max_entry.grid(row=7, column=3, padx=2, pady=2)

        ttk.Label(frame, text="Major Interval:").grid(row=8, column=0, padx=2, pady=2, sticky=tk.W)
        self.ytick_major_interval_entry = ttk.Entry(frame, textvariable=self.ytick_major_interval_var, width=10)
        self.ytick_major_interval_entry.grid(row=8, column=1, padx=2, pady=2)

        ttk.Label(frame, text="Direction:").grid(row=8, column=2, padx=2, pady=2, sticky=tk.W)
        self.ytick_direction_combo = ttk.Combobox(frame, textvariable=self.ytick_direction_var, values=['out', 'in', 'inout'], state='readonly', width=8)
        self.ytick_direction_combo.grid(row=8, column=3, padx=2, pady=2, sticky=tk.W)

        self.ytick_show_check = ttk.Checkbutton(frame, text="Show Major Ticks", variable=self.ytick_show_var)
        self.ytick_show_check.grid(row=9, column=0, padx=2, pady=2, sticky=tk.W)
        self.ytick_label_show_check = ttk.Checkbutton(frame, text="Show Tick Labels", variable=self.ytick_label_show_var)
        self.ytick_label_show_check.grid(row=9, column=1, padx=2, pady=2, sticky=tk.W)

        self.ytick_minor_show_check = ttk.Checkbutton(frame, text="Minor Ticks", variable=self.ytick_minor_show_var)
        self.ytick_minor_show_check.grid(row=9, column=2, padx=2, pady=2, sticky=tk.W)
        self.ytick_minor_interval_entry = ttk.Entry(frame, textvariable=self.ytick_minor_interval_var, width=8)
        self.ytick_minor_interval_entry.grid(row=9, column=3, padx=2, pady=2, sticky=tk.W)

        self.y1_log_scale_check = ttk.Checkbutton(frame, text="Log Scale", variable=self.y1_log_scale_var)
        self.y1_log_scale_check.grid(row=10, column=0, padx=2, pady=2, sticky=tk.W)
        self.y1_invert_check = ttk.Checkbutton(frame, text="Invert Axis", variable=self.y1_invert_var)
        self.y1_invert_check.grid(row=10, column=1, padx=2, pady=2, sticky=tk.W)
        self.yaxis_plain_format_check = ttk.Checkbutton(frame, text="Disable Sci Notation", variable=self.yaxis1_plain_format_var)
        self.yaxis_plain_format_check.grid(row=10, column=2, columnspan=2, padx=2, pady=2, sticky=tk.W)

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=11, column=0, columnspan=4, sticky="ew", pady=5)

        # Y-Axis (Right)
        ttk.Label(frame, text="Y-Axis (Right)", font=("-weight bold")).grid(row=12, column=0, sticky=tk.W, pady=3)
        ttk.Label(frame, text="Range Min:").grid(row=13, column=0, padx=2, pady=2, sticky=tk.W)
        self.ylim2_min_entry = ttk.Entry(frame, textvariable=self.ylim2_min_var, width=10)
        self.ylim2_min_entry.grid(row=13, column=1, padx=2, pady=2)

        ttk.Label(frame, text="Range Max:").grid(row=13, column=2, padx=2, pady=2, sticky=tk.W)
        self.ylim2_max_entry = ttk.Entry(frame, textvariable=self.ylim2_max_var, width=10)
        self.ylim2_max_entry.grid(row=13, column=3, padx=2, pady=2)

        ttk.Label(frame, text="Major Interval:").grid(row=14, column=0, padx=2, pady=2, sticky=tk.W)
        self.ytick2_major_interval_entry = ttk.Entry(frame, textvariable=self.ytick2_major_interval_var, width=10)
        self.ytick2_major_interval_entry.grid(row=14, column=1, padx=2, pady=2)

        ttk.Label(frame, text="Direction:").grid(row=14, column=2, padx=2, pady=2, sticky=tk.W)
        self.ytick2_direction_combo = ttk.Combobox(frame, textvariable=self.ytick2_direction_var, values=['out', 'in', 'inout'], state='readonly', width=8)
        self.ytick2_direction_combo.grid(row=14, column=3, padx=2, pady=2, sticky=tk.W)

        self.ytick2_show_check = ttk.Checkbutton(frame, text="Show Major Ticks", variable=self.ytick2_show_var)
        self.ytick2_show_check.grid(row=15, column=0, padx=2, pady=2, sticky=tk.W)
        self.ytick2_label_show_check = ttk.Checkbutton(frame, text="Show Tick Labels", variable=self.ytick2_label_show_var)
        self.ytick2_label_show_check.grid(row=15, column=1, padx=2, pady=2, sticky=tk.W)

        self.ytick2_minor_show_check = ttk.Checkbutton(frame, text="Minor Ticks", variable=self.ytick2_minor_show_var)
        self.ytick2_minor_show_check.grid(row=15, column=2, padx=2, pady=2, sticky=tk.W)
        self.ytick2_minor_interval_entry = ttk.Entry(frame, textvariable=self.ytick2_minor_interval_var, width=8)
        self.ytick2_minor_interval_entry.grid(row=15, column=3, padx=2, pady=2, sticky=tk.W)

        self.y2_log_scale_check = ttk.Checkbutton(frame, text="Log Scale", variable=self.y2_log_scale_var)
        self.y2_log_scale_check.grid(row=16, column=0, padx=2, pady=2, sticky=tk.W)
        self.y2_invert_check = ttk.Checkbutton(frame, text="Invert Axis", variable=self.y2_invert_var)
        self.y2_invert_check.grid(row=16, column=1, padx=2, pady=2, sticky=tk.W)
        self.yaxis2_plain_format_check = ttk.Checkbutton(frame, text="Disable Sci Notation", variable=self.yaxis2_plain_format_var)
        self.yaxis2_plain_format_check.grid(row=16, column=2, columnspan=2, padx=2, pady=2, sticky=tk.W)

    # --- Tab 5: Spines & Background ---
    def create_spines_tab(self, frame):
        ttk.Label(frame, text="Show Graph Spines:").grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=3)

        self.spine_top_check = ttk.Checkbutton(frame, text="Top", variable=self.spine_top_var)
        self.spine_top_check.grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)

        self.spine_bottom_check = ttk.Checkbutton(frame, text="Bottom", variable=self.spine_bottom_var)
        self.spine_bottom_check.grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)

        self.spine_left_check = ttk.Checkbutton(frame, text="Left", variable=self.spine_left_var)
        self.spine_left_check.grid(row=1, column=2, padx=5, pady=2, sticky=tk.W)

        self.spine_right_check = ttk.Checkbutton(frame, text="Right", variable=self.spine_right_var)
        self.spine_right_check.grid(row=1, column=3, padx=5, pady=2, sticky=tk.W)

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=2, column=0, columnspan=4, sticky="ew", pady=10)

        ttk.Label(frame, text="Axes Background Color:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.face_color_btn = ttk.Button(frame, text="Select", command=lambda: self.choose_color(self.face_color_var, self.face_color_label))
        self.face_color_btn.grid(row=3, column=1, padx=5, pady=5)
        self.face_color_label = ttk.Label(frame, text=self.face_color_var.get(), background=self.face_color_var.get(), width=10)
        self.face_color_label.grid(row=3, column=2, padx=5, pady=5)

        ttk.Label(frame, text="Figure Background Color:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.fig_color_btn = ttk.Button(frame, text="Select", command=lambda: self.choose_color(self.fig_color_var, self.fig_color_label))
        self.fig_color_btn.grid(row=4, column=1, padx=5, pady=5)
        self.fig_color_label = ttk.Label(frame, text=self.fig_color_var.get(), background=self.fig_color_var.get(), width=10)
        self.fig_color_label.grid(row=4, column=2, padx=5, pady=5)

    # --- Tab 6: Legend Settings ---
    def create_legend_tab(self, frame):
        self.legend_show_check = ttk.Checkbutton(frame, text="Show Legend (Auto-combined)", variable=self.legend_show_var)
        self.legend_show_check.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)

        ttk.Label(frame, text="Legend Location:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.legend_loc_combo = ttk.Combobox(
            frame, textvariable=self.legend_loc_var,
            values=['best', 'upper right', 'upper left', 'lower left', 'lower right', 'right', 'center left', 'center right', 'lower center', 'upper center', 'center'],
            state='readonly', width=25
        )
        self.legend_loc_combo.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

    # --- Tab 7: Advanced Settings ---
    def create_advanced_tab(self, frame):
        # Data Smoothing
        smoothing_frame = ttk.LabelFrame(frame, text="Data Smoothing (Moving Average)")
        smoothing_frame.pack(fill=tk.X, padx=5, pady=3)

        self.enable_smoothing_check = ttk.Checkbutton(smoothing_frame, text="Enable Smoothing", variable=self.enable_smoothing_var)
        self.enable_smoothing_check.grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)

        ttk.Label(smoothing_frame, text="Window Size:").grid(row=0, column=1, padx=5, pady=3, sticky=tk.W)
        self.smoothing_window_spin = ttk.Spinbox(smoothing_frame, from_=2, to=50, increment=1, textvariable=self.smoothing_window_var, width=8)
        self.smoothing_window_spin.grid(row=0, column=2, padx=5, pady=3, sticky=tk.W)

        # Error Bars
        errorbar_frame = ttk.LabelFrame(frame, text="Error Bars")
        errorbar_frame.pack(fill=tk.X, padx=5, pady=3)

        self.enable_errorbar_check = ttk.Checkbutton(errorbar_frame, text="Enable Error Bars", variable=self.enable_errorbar_var)
        self.enable_errorbar_check.grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)

        ttk.Label(errorbar_frame, text="Error Column:").grid(row=0, column=1, padx=5, pady=3, sticky=tk.W)
        self.errorbar_column_combo = ttk.Combobox(errorbar_frame, textvariable=self.errorbar_column_var, state='readonly', width=20)
        self.errorbar_column_combo.grid(row=0, column=2, padx=5, pady=3, sticky=tk.W)

        # Annotations
        annotation_frame = ttk.LabelFrame(frame, text="Data Point Annotations")
        annotation_frame.pack(fill=tk.X, padx=5, pady=3)
        self.enable_annotation_check = ttk.Checkbutton(annotation_frame, text="Show Data Values", variable=self.enable_annotation_var)
        self.enable_annotation_check.grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)

        # Data Filtering
        filter_frame = ttk.LabelFrame(frame, text="Data Filtering (Non-Destructive)")
        filter_frame.pack(fill=tk.X, padx=5, pady=3)

        self.data_filter_check = ttk.Checkbutton(filter_frame, text="Enable Filter", variable=self.data_filter_enabled_var)
        self.data_filter_check.grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)

        ttk.Label(filter_frame, text="Filter Column:").grid(row=0, column=1, padx=5, pady=3, sticky=tk.W)
        self.filter_column_combo = ttk.Combobox(filter_frame, textvariable=self.filter_column_var, state='readonly', width=18)
        self.filter_column_combo.grid(row=0, column=2, padx=5, pady=3, sticky=tk.W)

        ttk.Label(filter_frame, text="Min:").grid(row=1, column=0, padx=5, pady=3, sticky=tk.W)
        self.filter_min_entry = ttk.Entry(filter_frame, textvariable=self.filter_min_var, width=10)
        self.filter_min_entry.grid(row=1, column=1, padx=5, pady=3, sticky=tk.W)

        ttk.Label(filter_frame, text="Max:").grid(row=1, column=2, padx=5, pady=3, sticky=tk.W)
        self.filter_max_entry = ttk.Entry(filter_frame, textvariable=self.filter_max_var, width=10)
        self.filter_max_entry.grid(row=1, column=3, padx=5, pady=3, sticky=tk.W)

        # Grid Customization & Layout Padding
        grid_style_frame = ttk.LabelFrame(frame, text="Grid & Layout Settings")
        grid_style_frame.pack(fill=tk.X, padx=5, pady=3)

        ttk.Label(grid_style_frame, text="Grid Alpha:").grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)
        self.grid_alpha_spin = ttk.Spinbox(grid_style_frame, from_=0.0, to=1.0, increment=0.1, textvariable=self.grid_alpha_var, width=8)
        self.grid_alpha_spin.grid(row=0, column=1, padx=5, pady=3, sticky=tk.W)

        ttk.Label(grid_style_frame, text="Grid Style:").grid(row=0, column=2, padx=5, pady=3, sticky=tk.W)
        self.grid_linestyle_combo = ttk.Combobox(grid_style_frame, textvariable=self.grid_linestyle_var, values=['-', '--', '-.', ':'], state='readonly', width=8)
        self.grid_linestyle_combo.grid(row=0, column=3, padx=5, pady=3, sticky=tk.W)

        ttk.Label(grid_style_frame, text="Grid Width:").grid(row=1, column=0, padx=5, pady=3, sticky=tk.W)
        self.grid_linewidth_spin = ttk.Spinbox(grid_style_frame, from_=0.1, to=3.0, increment=0.1, textvariable=self.grid_linewidth_var, width=8)
        self.grid_linewidth_spin.grid(row=1, column=1, padx=5, pady=3, sticky=tk.W)

        ttk.Label(grid_style_frame, text="Layout Pad:").grid(row=1, column=2, padx=5, pady=3, sticky=tk.W)
        self.tight_layout_pad_spin = ttk.Spinbox(grid_style_frame, from_=0.1, to=5.0, increment=0.5, textvariable=self.tight_layout_pad_var, width=8)
        self.tight_layout_pad_spin.grid(row=1, column=3, padx=5, pady=3, sticky=tk.W)

        # Colormap & Export DPI
        extended_frame = ttk.LabelFrame(frame, text="Extended Plot & Export Options")
        extended_frame.pack(fill=tk.X, padx=5, pady=3)

        ttk.Label(extended_frame, text="Colormap (Heatmap/Contour):").grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)
        self.colormap_combo = ttk.Combobox(
            extended_frame, textvariable=self.colormap_var,
            values=['viridis', 'plasma', 'inferno', 'magma', 'cividis', 'coolwarm', 'jet', 'rainbow', 'turbo', 'gray'],
            state='readonly', width=12
        )
        self.colormap_combo.grid(row=0, column=1, padx=5, pady=3, sticky=tk.W)

        ttk.Label(extended_frame, text="Export DPI:").grid(row=0, column=2, padx=5, pady=3, sticky=tk.W)
        self.export_dpi_combo = ttk.Combobox(
            extended_frame, textvariable=self.export_dpi_var,
            values=[72, 100, 150, 300, 600], state='readonly', width=8
        )
        self.export_dpi_combo.grid(row=0, column=3, padx=5, pady=3, sticky=tk.W)

        # Label Rotation & Subplot mode
        misc_frame = ttk.LabelFrame(frame, text="Layout & Rotation")
        misc_frame.pack(fill=tk.X, padx=5, pady=3)

        self.rotate_labels_check = ttk.Checkbutton(misc_frame, text="Rotate X Labels", variable=self.rotate_labels_var)
        self.rotate_labels_check.grid(row=0, column=0, padx=5, pady=3, sticky=tk.W)

        ttk.Label(misc_frame, text="Angle:").grid(row=0, column=1, padx=5, pady=3, sticky=tk.W)
        self.rotation_angle_spin = ttk.Spinbox(misc_frame, from_=0, to=90, increment=15, textvariable=self.rotation_angle_var, width=8)
        self.rotation_angle_spin.grid(row=0, column=2, padx=5, pady=3, sticky=tk.W)

        self.subplot_mode_check = ttk.Checkbutton(misc_frame, text="Split Y1 & Y2 Subplots", variable=self.subplot_mode_var)
        self.subplot_mode_check.grid(row=0, column=3, padx=5, pady=3, sticky=tk.W)

    # --- Callbacks for Style Editor ---
    def on_combined_series_select(self, event=None):
        selected_item = self.combined_style_target_var.get()
        if not selected_item:
            return

        is_y1 = "(Y1)" in selected_item
        series_name = selected_item.split("] ", 1)[-1] if "] " in selected_item else selected_item.replace("(Y1) ", "").replace("(Y2) ", "")

        if series_name:
            self.load_style_to_editor(series_name, is_y1=is_y1)

    def load_style_to_editor(self, series_name, is_y1):
        if series_name is None:
            self.current_style_color_var.set("#000000")
            self.current_style_linestyle_var.set("-")
            self.current_style_marker_var.set("o")
            self.current_style_linewidth_var.set(1.5)
            self.current_style_markersize_var.set(6.0)
            self.current_style_alpha_var.set(1.0)
            self.update_color_label(self.style_editor_color_label, "#000000")
            return

        styles_dict = self.y1_series_styles if is_y1 else self.y2_series_styles
        series_style = self.get_or_create_default_style(series_name, styles_dict)

        self.current_style_color_var.set(series_style.get('color', 'None'))
        self.current_style_linestyle_var.set(series_style.get('linestyle', '-'))
        self.current_style_marker_var.set(series_style.get('marker', 'o'))
        self.current_style_linewidth_var.set(series_style.get('linewidth', 1.5))
        self.current_style_markersize_var.set(series_style.get('markersize', 6.0))
        self.current_style_alpha_var.set(series_style.get('alpha', 1.0))

        self.update_color_label(self.style_editor_color_label, self.current_style_color_var.get())

    def get_or_create_default_style(self, series_name, styles_dict):
        if series_name not in styles_dict:
            styles_dict[series_name] = {
                'color': None,
                'linestyle': '-',
                'marker': 'o',
                'linewidth': 1.5,
                'markersize': 6.0,
                'alpha': 1.0
            }
        return styles_dict[series_name]

    def on_style_editor_change(self, event=None):
        selected_item = self.combined_style_target_var.get()
        if not selected_item:
            return

        is_y1 = "(Y1)" in selected_item
        series_name = selected_item.split("] ", 1)[-1] if "] " in selected_item else selected_item.replace("(Y1) ", "").replace("(Y2) ", "")
        styles_dict = self.y1_series_styles if is_y1 else self.y2_series_styles

        if series_name not in styles_dict:
            styles_dict[series_name] = {}

        try:
            styles_dict[series_name]['color'] = self.current_style_color_var.get()
            styles_dict[series_name]['linestyle'] = self.current_style_linestyle_var.get()
            styles_dict[series_name]['marker'] = self.current_style_marker_var.get()
            styles_dict[series_name]['linewidth'] = self.current_style_linewidth_var.get()
            styles_dict[series_name]['markersize'] = self.current_style_markersize_var.get()
            styles_dict[series_name]['alpha'] = self.current_style_alpha_var.get()
        except tk.TclError as e:
            print(f"Style update error: {e}")

    def on_style_editor_color_pick(self):
        initial_color = self.current_style_color_var.get()
        if not initial_color or initial_color == 'None':
            initial_color = '#000000'

        color_code = colorchooser.askcolor(title="Choose Color", initialcolor=initial_color)[1]
        if color_code:
            self.current_style_color_var.set(color_code)
            self.update_color_label(self.style_editor_color_label, color_code)
            self.on_style_editor_change()

    def on_style_editor_color_auto(self):
        self.current_style_color_var.set('None')
        self.update_color_label(self.style_editor_color_label, 'None')
        self.on_style_editor_change()

    def choose_color(self, color_var, color_label):
        color_code = colorchooser.askcolor(title="Choose Color", initialcolor=color_var.get())[1]
        if color_code:
            color_var.set(color_code)
            color_label.config(background=color_code, text=color_code)

    def load_data(self, file_path=None):
        if file_path is None:
            file_path = filedialog.askopenfilename(
                title="Select Data File",
                filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
        if not file_path:
            return

        try:
            self.data_mgr.load_file(file_path)
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to read file:\n{e}")
            return

        if self.sheet:
            self.sheet.destroy()

        self.sheet = Sheet(
            self.sheet_frame,
            data=self.df.values.tolist(),
            headers=self.df.columns.tolist(),
            show_toolbar=True,
            show_top_left=True
        )
        self.sheet.enable_bindings()
        self.sheet.pack(fill=tk.BOTH, expand=True)

        self.update_plot_options()

    def update_plot_options(self):
        if self.df is None:
            return

        columns = self.data_mgr.get_columns()

        for info in self.x_tab_widgets:
            info["x_combo"]['values'] = columns
            info["x_combo"]['state'] = 'readonly'
            info["y1_listbox"].delete(0, tk.END)
            info["y2_listbox"].delete(0, tk.END)
            for col in columns:
                info["y1_listbox"].insert(tk.END, col)
                info["y2_listbox"].insert(tk.END, col)

        self.plot_button['state'] = 'normal'
        self.export_button['state'] = 'normal'
        self.export_data_button['state'] = 'normal'

        if columns and self.x_tab_widgets:
            init_tab = self.x_tab_widgets[0]
            self.x_axis_var.set(columns[0])
            self.xlabel_var.set(columns[0])
            if len(columns) > 1:
                init_tab["y1_listbox"].select_set(1)
            else:
                init_tab["y1_listbox"].select_set(0)

        if hasattr(self, 'errorbar_column_combo'):
            self.errorbar_column_combo['values'] = [''] + columns
        if hasattr(self, 'filter_column_combo'):
            self.filter_column_combo['values'] = columns

    def get_data_from_sheet(self):
        if not self.sheet or self.df is None:
            return

        try:
            data = None
            if hasattr(self.sheet, 'get_sheet_data') and callable(self.sheet.get_sheet_data):
                data = self.sheet.get_sheet_data()
            elif hasattr(self.sheet, 'data') and isinstance(self.sheet.data, list):
                data = self.sheet.data
            else:
                return

            headers = None
            if hasattr(self.sheet, 'get_headers') and callable(self.sheet.get_headers):
                headers = self.sheet.get_headers()
            elif hasattr(self.sheet, 'headers'):
                headers = self.sheet.headers() if callable(self.sheet.headers) else self.sheet.headers
            else:
                return

            self.data_mgr.update_from_sheet_data(data, headers)
        except Exception as e:
            print(f"Data Retrieval Error: {e}")

    def save_settings(self):
        file_path = filedialog.asksaveasfilename(
            title="Save Project",
            filetypes=[("Matplotlib Graph Project", "*.pmggrp")],
            defaultextension=".pmggrp"
        )
        if not file_path:
            return

        try:
            save_project_file(self, file_path, version_str=VERSION, dimension="2D")
            self.current_project_path = file_path
            self.overwrite_save_button['state'] = 'normal'
            messagebox.showinfo("Success", f"Project saved to {file_path}.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save project:\n{e}")

    def overwrite_save(self):
        if not self.current_project_path:
            self.save_settings()
            return
        try:
            save_project_file(self, self.current_project_path, version_str=VERSION, dimension="2D")
            messagebox.showinfo("Success", f"Project saved to {self.current_project_path}.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save project:\n{e}")

    def load_settings(self):
        file_path = filedialog.askopenfilename(
            title="Select Project File",
            filetypes=[("Matplotlib Graph Project", "*.pmggrp")]
        )
        if not file_path:
            return
        self.load_project_file(file_path)

    def set_variable_from_dict(self, var, settings_dict, key, fallback_key=None):
        value_to_set = settings_dict.get(key, settings_dict.get(fallback_key) if fallback_key else None)
        if value_to_set is not None:
            try:
                var.set(value_to_set)
            except Exception as e:
                print(f"Warning setting var '{key}': {e}")

    def update_color_label(self, label, color_code):
        if not color_code or color_code == 'None':
            color_code = "#FFFFFF"
            text = "Auto"
        else:
            text = color_code
        try:
            label.config(background=color_code, text=text, anchor=tk.CENTER)
        except tk.TclError:
            label.config(background="#FFFFFF", text="Invalid", anchor=tk.CENTER)

    # --- Plotting Method ---
    def plot_graph(self):
        try:
            self.fig.clear()
            if self.subplot_mode_var.get():
                self.ax = self.fig.add_subplot(211)
                self.ax2 = self.fig.add_subplot(212)
            else:
                self.ax = self.fig.add_subplot(111)
                self.ax2 = None
            self.canvas.draw()
        except Exception as e:
            messagebox.showerror("Internal Error", f"Failed to clear graph:\n{e}")
            return

        self.get_data_from_sheet()
        if not self.data_mgr.has_data():
            messagebox.showinfo("Info", "No data to plot.")
            self.canvas.draw()
            return

        # Retrieve active non-destructively filtered DataFrame
        plot_df = self.data_mgr.get_filtered_df(
            filter_enabled=self.data_filter_enabled_var.get(),
            filter_column=self.filter_column_var.get(),
            min_val_str=self.filter_min_var.get(),
            max_val_str=self.filter_max_var.get()
        )

        plot_type = self.plot_type_var.get()
        x_tabs_info = self.get_x_tabs_data()

        # Update Style Tab Combobox
        combined_list = []
        for tab_idx, tab_data in enumerate(x_tabs_info):
            t_name = f"T{tab_idx+1}"
            for c in tab_data["y1_cols"]:
                combined_list.append(f"({t_name}-Y1) {c}")
            for c in tab_data["y2_cols"]:
                combined_list.append(f"({t_name}-Y2) {c}")

        self.style_combo['values'] = combined_list
        if self.combined_style_target_var.get() not in combined_list:
            self.combined_style_target_var.set("")
            self.load_style_to_editor(None, True)

        # Basic validation
        has_any_y = any(tab["y1_cols"] or tab["y2_cols"] for tab in x_tabs_info)
        if not has_any_y:
            messagebox.showerror("Error", "Please select Y-Axis data in at least one X-Tab.")
            self.canvas.draw()
            return

        try:
            self.fig.set_size_inches(self.fig_width_var.get(), self.fig_height_var.get())
            self.fig.set_facecolor(self.fig_color_var.get())
            matplotlib.rcParams['font.family'] = self.font_family_var.get()

            # Check if Y2 is used anywhere across tabs
            has_y2 = any(tab["y2_cols"] for tab in x_tabs_info)
            if has_y2 and not self.subplot_mode_var.get():
                self.ax2 = self.ax.twinx()
        except Exception as e:
            messagebox.showerror("Settings Error", f"Failed to apply basic settings:\n{e}")
            return

        try:
            # Helper to plot a series
            def plot_series(ax, x_col, y_col, is_twin_ax, label_prefix=""):
                styles_dict = self.y1_series_styles if not is_twin_ax else self.y2_series_styles
                series_style = self.get_or_create_default_style(y_col, styles_dict)

                x_data_raw = plot_df[x_col]
                y_data_cleaned = plot_df[y_col].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                y_data_numeric = pd.to_numeric(y_data_cleaned, errors='coerce')

                color = series_style.get('color', None)
                if color == 'None':
                    color = None
                linestyle = series_style.get('linestyle', '-')
                linewidth = series_style.get('linewidth', 1.5)
                markersize = series_style.get('markersize', 6.0)
                alpha = series_style.get('alpha', 1.0)

                markerstyle = series_style.get('marker', 'o')
                if not self.marker_var.get():
                    markerstyle = 'None'

                display_label = f"{label_prefix}: {y_col}" if label_prefix else y_col

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

                    if plot_type == "line" and self.enable_smoothing_var.get() and len(plot_y) >= self.smoothing_window_var.get():
                        plot_y = plot_y.rolling(window=self.smoothing_window_var.get(), center=True).mean().fillna(plot_y)

                    kwargs = {'marker': markerstyle, 'markersize': markersize, 'alpha': alpha, 'label': display_label}
                    if color:
                        kwargs['color'] = color

                    if plot_type == "line":
                        kwargs.update({'linestyle': linestyle, 'linewidth': linewidth})
                        ax.plot(plot_x, plot_y, **kwargs)
                    elif plot_type == "scatter":
                        ax.scatter(plot_x, plot_y, **kwargs)
                    elif plot_type == "step":
                        kwargs.update({'linestyle': linestyle, 'linewidth': linewidth})
                        ax.step(plot_x, plot_y, where='mid', **kwargs)
                    elif plot_type == "area":
                        kwargs.update({'linestyle': linestyle, 'linewidth': linewidth})
                        ax.fill_between(plot_x, 0, plot_y, **kwargs)

            # Loop through all X-Tabs and plot
            num_tabs = len(x_tabs_info)
            for tab_idx, tab_data in enumerate(x_tabs_info):
                x_col = tab_data["x_axis"]
                if not x_col or x_col not in plot_df.columns:
                    continue

                prefix = f"X-Tab {tab_idx+1}" if num_tabs > 1 else ""

                for y_col in tab_data["y1_cols"]:
                    if y_col in plot_df.columns:
                        plot_series(self.ax, x_col, y_col, is_twin_ax=False, label_prefix=prefix)

                target_ax = self.ax2 if self.ax2 else self.ax
                for y_col in tab_data["y2_cols"]:
                    if y_col in plot_df.columns:
                        plot_series(target_ax, x_col, y_col, is_twin_ax=True, label_prefix=prefix)

            # Apply Scales
            try:
                self.ax.set_xscale('log' if self.x_log_scale_var.get() else 'linear')
            except ValueError:
                self.ax.set_xscale('linear')
            try:
                self.ax.set_yscale('log' if self.y1_log_scale_var.get() else 'linear')
            except ValueError:
                self.ax.set_yscale('linear')

            # Axis Labels & Title Handling
            font_family = self.font_family_var.get()

            # Automatic title if title_var is empty
            if self.title_var.get().strip():
                auto_title = self.title_var.get()
            else:
                y_names = []
                for tab in x_tabs_info:
                    y_names.extend(tab["y1_cols"] + tab["y2_cols"])
                auto_title = f"{', '.join(y_names)} vs {x_tabs_info[0]['x_axis']}" if y_names else "Data Plot"

            self.ax.set_title(auto_title, fontsize=self.title_fontsize_var.get(), fontfamily=font_family)
            self.ax.set_xlabel(self.xlabel_var.get() if self.xlabel_var.get() else x_tabs_info[0]["x_axis"], fontsize=self.xlabel_fontsize_var.get(), fontfamily=font_family)
            self.ax.set_ylabel(self.ylabel_var.get() if self.ylabel_var.get() else "Y-Values", fontsize=self.ylabel_fontsize_var.get(), fontfamily=font_family)

            # Grid
            if self.grid_var.get():
                self.ax.grid(
                    True, alpha=self.grid_alpha_var.get(),
                    linestyle=self.grid_linestyle_var.get(),
                    linewidth=self.grid_linewidth_var.get()
                )
            else:
                self.ax.grid(False)

            # Axis limits & Inversion
            self.set_axis_limits(self.ax, 'x', self.xlim_min_var.get(), self.xlim_max_var.get())
            self.set_axis_limits(self.ax, 'y', self.ylim_min_var.get(), self.ylim_max_var.get())
            if self.x_invert_var.get():
                self.ax.invert_xaxis()
            if self.y1_invert_var.get():
                self.ax.invert_yaxis()

            # Tickers
            self.ax.tick_params(
                axis='x', which='both',
                direction=self.xtick_direction_var.get(),
                bottom=self.xtick_show_var.get(),
                labelbottom=self.xtick_label_show_var.get(),
                labelsize=self.tick_fontsize_var.get()
            )
            self.ax.tick_params(
                axis='y', which='both',
                direction=self.ytick_direction_var.get(),
                left=self.ytick_show_var.get(),
                labelleft=self.ytick_label_show_var.get(),
                labelsize=self.tick_fontsize_var.get()
            )

            apply_major_ticker(self.ax.xaxis, self.xtick_major_interval_var.get(), self.x_log_scale_var.get())
            apply_major_ticker(self.ax.yaxis, self.ytick_major_interval_var.get(), self.y1_log_scale_var.get())
            apply_minor_ticker(self.ax.xaxis, self.xtick_minor_show_var.get(), self.xtick_minor_interval_var.get(), self.x_log_scale_var.get())
            apply_minor_ticker(self.ax.yaxis, self.ytick_minor_show_var.get(), self.ytick_minor_interval_var.get(), self.y1_log_scale_var.get())

            # Spines & Facecolor
            self.ax.set_facecolor(self.face_color_var.get())
            self.ax.spines['top'].set_visible(self.spine_top_var.get())
            self.ax.spines['bottom'].set_visible(self.spine_bottom_var.get())
            self.ax.spines['left'].set_visible(self.spine_left_var.get())
            self.ax.spines['right'].set_visible(self.spine_right_var.get())

            if self.ax2:
                self.ax2.set_ylabel(self.ylabel2_var.get() if self.ylabel2_var.get() else "Y2-Values", fontsize=self.ylabel2_fontsize_var.get(), fontfamily=font_family)
                self.set_axis_limits(self.ax2, 'y', self.ylim2_min_var.get(), self.ylim2_max_var.get())
                apply_major_ticker(self.ax2.yaxis, self.ytick2_major_interval_var.get(), self.y2_log_scale_var.get())
                apply_minor_ticker(self.ax2.yaxis, self.ytick2_minor_show_var.get(), self.ytick2_minor_interval_var.get(), self.y2_log_scale_var.get())

            # Legend
            if self.legend_show_var.get():
                legend_props = {'family': font_family, 'size': self.legend_fontsize_var.get()}
                h1, l1 = self.ax.get_legend_handles_labels()
                h2, l2 = (self.ax2.get_legend_handles_labels() if self.ax2 else ([], []))
                self.ax.legend(handles=h1 + h2, labels=l1 + l2, loc=self.legend_loc_var.get(), prop=legend_props)

            pad_val = self.tight_layout_pad_var.get()
            self.fig.tight_layout(pad=pad_val)
            self.canvas.draw()

            fig_width_px = self.fig.get_figwidth() * self.fig.dpi
            fig_height_px = self.fig.get_figheight() * self.fig.dpi
            self.canvas.get_tk_widget().config(width=int(fig_width_px), height=int(fig_height_px))
            self.graph_frame.update_idletasks()
            self.on_graph_frame_configure(None)

        except Exception as e:
            messagebox.showerror("Plot Error", f"Failed to plot graph:\n{e}")
            self.canvas.draw()

    def set_axis_limits(self, ax, axis_name, min_val, max_val):
        try:
            min_v = float(min_val) if min_val else None
            max_v = float(max_val) if max_val else None
            if axis_name == 'x':
                ax.set_xlim(min_v, max_v)
            elif axis_name == 'y':
                ax.set_ylim(min_v, max_v)
        except ValueError:
            pass

    def export_graph(self):
        file_path = filedialog.asksaveasfilename(
            title="Save Graph",
            filetypes=[("PNG files", "*.png"), ("JPEG files", "*.jpg *.jpeg"), ("SVG files", "*.svg"), ("PDF files", "*.pdf"), ("All files", "*.*")],
            defaultextension=".png"
        )
        if not file_path:
            return

        try:
            dpi_val = self.export_dpi_var.get()
            self.fig.savefig(file_path, dpi=dpi_val, bbox_inches='tight', facecolor=self.fig_color_var.get())
            messagebox.showinfo("Success", f"Graph saved to {file_path} ({dpi_val} DPI).")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save graph:\n{e}")

    def export_filtered_data(self):
        if not self.data_mgr.has_data():
            messagebox.showinfo("Info", "No data to export.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Export Data",
            filetypes=[("CSV files", "*.csv")],
            defaultextension=".csv"
        )
        if not file_path:
            return

        try:
            df_to_export = self.data_mgr.get_filtered_df(
                filter_enabled=self.data_filter_enabled_var.get(),
                filter_column=self.filter_column_var.get(),
                min_val_str=self.filter_min_var.get(),
                max_val_str=self.filter_max_var.get()
            )
            df_to_export.to_csv(file_path, index=False)
            messagebox.showinfo("Success", f"Data exported to {file_path}.")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export data:\n{e}")

    def clear_all(self):
        if messagebox.askyesno("Confirm Clear", "Clear all data and settings?"):
            self.data_mgr.clear()
            self.current_project_path = ""
            if self.sheet:
                self.sheet.destroy()
                self.sheet = None
            self.fig.clear()
            self.ax = self.fig.add_subplot(111)
            self.ax2 = None
            self.canvas.draw()

            self.plot_button['state'] = 'disabled'
            self.export_button['state'] = 'disabled'
            self.export_data_button['state'] = 'disabled'
            self.overwrite_save_button['state'] = 'disabled'

    def reset_settings(self):
        if messagebox.askyesno("Confirm Reset", "Reset all settings to default values (keep data)?"):
            self.plot_type_var.set("line")
            self.title_var.set("")
            self.xlabel_var.set("")
            self.ylabel_var.set("")
            self.ylabel2_var.set("")
            self.y1_series_styles = {}
            self.y2_series_styles = {}
            self.grid_var.set(False)
            self.marker_var.set(True)
            messagebox.showinfo("Reset Complete", "All settings reset to defaults.")

    def setup_drag_and_drop(self):
        if DND_AVAILABLE:
            try:
                from tkinterdnd2 import DND_FILES
                self.drop_target_register(DND_FILES)
                self.dnd_bind('<<Drop>>', self.on_drop)
            except Exception as e:
                print(f"DnD error: {e}")

    def on_drop(self, event):
        try:
            files = getattr(event, 'data', '')
            if not files:
                return
            try:
                files = self.tk.splitlist(files)
            except Exception:
                pass
            file_path = files[0] if isinstance(files, (list, tuple)) else files
            file_path = str(file_path).strip('{}').strip()

            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.pmggrp':
                self.load_project_file(file_path)
            elif ext in ['.csv', '.xlsx', '.xls']:
                self.load_data(file_path=file_path)
        except Exception as e:
            messagebox.showerror("Drop Error", f"Failed to load dropped file:\n{e}")

    def load_project_file(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to read project file:\n{e}")
            return

        if settings.get('dimension') == '3D':
            if messagebox.askyesno("3D Project", "This project is in 3D mode. Open in 3D mode?"):
                self.open_in_3d_mode(project_path=file_path)
                self.quit()
                return

        if 'edited_data' in settings and settings['edited_data']:
            d = settings['edited_data']
            self.df = pd.DataFrame(d['data'], columns=d['columns']).astype(str).fillna("")
            self.data_file_path = settings.get('original_file_path', '')

            if self.sheet:
                self.sheet.destroy()
            self.sheet = Sheet(
                self.sheet_frame,
                data=self.df.values.tolist(),
                headers=self.df.columns.tolist(),
                show_toolbar=True,
                show_top_left=True
            )
            self.sheet.enable_bindings()
            self.sheet.pack(fill=tk.BOTH, expand=True)
            self.update_plot_options()

        self.set_variable_from_dict(self.plot_type_var, settings, 'plot_type')
        self.set_variable_from_dict(self.title_var, settings, 'title')
        self.set_variable_from_dict(self.xlabel_var, settings, 'xlabel')
        self.set_variable_from_dict(self.ylabel_var, settings, 'ylabel')
        self.set_variable_from_dict(self.ylabel2_var, settings, 'ylabel2')

        self.current_project_path = file_path
        self.overwrite_save_button['state'] = 'normal'
        if self.df is not None:
            self.plot_graph()

    def open_in_3d_mode(self, project_path=None):
        import subprocess
        script_dir = os.path.dirname(os.path.abspath(__file__))
        main_3d_path = os.path.join(script_dir, "main_3d.py")
        target_path = project_path or self.current_project_path
        cmd = [sys.executable, main_3d_path]
        if target_path and os.path.exists(target_path):
            cmd.append(target_path)
        subprocess.Popen(cmd)

    def show_about(self):
        about_text = f"""HYGrapher ver. {VERSION}

A cross-platform graphing application for CSV/Excel data.

Author: Hiromichi Yokoyama
License: Apache-2.0 license
Repository: https://github.com/HiroYokoyama/matplotlib_graph_app
"""
        messagebox.showinfo("About HYGrapher", about_text)


def main():
    app = GraphApp()
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if os.path.exists(file_path):
            ext = os.path.splitext(file_path)[1].lower()
            def load_startup_file():
                if ext == '.pmggrp':
                    app.load_project_file(file_path)
                elif ext in ['.csv', '.xlsx', '.xls']:
                    app.load_data(file_path=file_path)
            app.after(100, load_startup_file)
    app.mainloop()


if __name__ == "__main__":
    main()
