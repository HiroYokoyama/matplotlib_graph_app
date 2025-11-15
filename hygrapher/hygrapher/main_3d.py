#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
hygrapher-3D

Author: Hiromichi Yokoyama
License: Apache-2.0 license
Repo: https://github.com/HiroYokoyama/matplotlib_graph_app
"""

VERSION = "0.5.2"

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.font_manager as fm
import matplotlib
import matplotlib.ticker as ticker
from mpl_toolkits.mplot3d import Axes3D
import numpy as np
import os
import json
from tksheet import Sheet
import sys

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
        # 1. UI English: Window Title
        self.title(f"HYGrapher 3D ver. {VERSION}")
        self.geometry("1600x900") # Keep window size

        self.df = None
        self.sheet = None
        self.data_file_path = "" # Store the path of the loaded file
        self.current_project_path = "" # Store the path of the current project file
        
        # Enable drag and drop
        self.setup_drag_and_drop()

        # --- Get Font List ---
        self.font_list = self.get_font_list()

        # --- Create all tk.Variables ---
        self.create_all_tk_variables()

        # --- Figure ---
        self.fig = Figure(figsize=(self.fig_width_var.get(), self.fig_height_var.get()), dpi=100)
        self.ax = self.fig.add_subplot(111, projection='3d')
        self.ax2 = None # For 2nd Z-axis (not typically used in 3D)

        # === Menu Bar ===
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Load Data (CSV/Excel)...", command=self.load_data, accelerator="")
        file_menu.add_separator()
        file_menu.add_command(label="Open Project...", command=self.load_settings, accelerator="Ctrl+O")
        file_menu.add_command(label="Save Project", command=self.overwrite_save, accelerator="Ctrl+S")
        file_menu.add_command(label="Save Project As...", command=self.save_settings, accelerator="")
        file_menu.add_separator()
        file_menu.add_command(label="Export Graph...", command=self.export_graph, accelerator="")
        file_menu.add_command(label="Export Data (CSV)...", command=self.export_filtered_data, accelerator="")
        file_menu.add_separator()
        file_menu.add_command(label="Open in 2D Mode...", command=self.open_in_2d_mode, accelerator="")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit, accelerator="")
        
        # Edit Menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Clear All", command=self.clear_all, accelerator="")
        edit_menu.add_command(label="Reset Settings", command=self.reset_settings, accelerator="")
        
        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about, accelerator="")

        # === Main Layout ===
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Top Frame (File Operations) ---
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 5))

        # Data Operations Frame (Left)
        data_ops_frame = ttk.LabelFrame(top_frame, text="Data")
        data_ops_frame.pack(side=tk.LEFT, padx=(0, 5))
        
        self.load_button = ttk.Button(data_ops_frame, text="Load Data", command=self.load_data, width=12)
        self.load_button.pack(side=tk.LEFT, padx=2, pady=2)

        self.clear_button = ttk.Button(data_ops_frame, text="Clear All", command=self.clear_all, width=12)
        self.clear_button.pack(side=tk.LEFT, padx=2, pady=2)

        # Project Operations Frame (Center)
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

        # Export Operations Frame (Right)
        export_ops_frame = ttk.LabelFrame(top_frame, text="Export")
        export_ops_frame.pack(side=tk.LEFT, padx=(5, 0))
        
        self.export_button = ttk.Button(export_ops_frame, text="Graph", command=self.export_graph, width=10)
        self.export_button.pack(side=tk.LEFT, padx=2, pady=2)
        self.export_button['state'] = 'disabled'
        
        self.export_data_button = ttk.Button(export_ops_frame, text="Data", command=self.export_filtered_data, width=10)
        self.export_data_button.pack(side=tk.LEFT, padx=2, pady=2)
        self.export_data_button['state'] = 'disabled'
        
        # Bind keyboard shortcuts
        self.bind('<Control-s>', lambda e: self.overwrite_save())
        self.bind('<Control-o>', lambda e: self.load_settings())

        # --- Content Frame (Split) ---
        # 2. Layout: Use PanedWindow for resizable 1:1 split
        content_frame = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # --- Left Panel (Data and Settings) ---
        # Create a container frame for left panel
        left_container = ttk.Frame(content_frame)
        content_frame.add(left_container, weight=1)
        
        # Create PanedWindow for vertical split (settings top, data editor bottom)
        left_paned = ttk.PanedWindow(left_container, orient=tk.VERTICAL)
        left_paned.pack(fill=tk.BOTH, expand=True)
        
        # --- Top section: Settings (scrollable) ---
        settings_container = ttk.Frame(left_paned)
        left_paned.add(settings_container, weight=1)
        
        # --- Plot Button (Fixed at bottom of settings area) ---
        # 1. UI English: Button text
        self.plot_button = ttk.Button(settings_container, text="Plot/Update Graph", command=self.plot_graph)
        self.plot_button.pack(side=tk.BOTTOM, fill=tk.X, pady=5, padx=5)
        self.plot_button['state'] = 'disabled'
        
        # --- Graph Settings (Notebook) - No outer scroll ---
        self.settings_notebook = ttk.Notebook(settings_container)
        self.settings_notebook.pack(fill=tk.BOTH, expand=True, pady=2)
        
        # Helper function to create scrollable tab content
        def create_scrollable_tab(parent_notebook, tab_name, create_content_func):
            # Create container frame for the tab
            tab_container = ttk.Frame(parent_notebook)
            parent_notebook.add(tab_container, text=tab_name)
            
            # Create canvas and scrollbar
            tab_canvas = tk.Canvas(tab_container, borderwidth=0, highlightthickness=0)
            tab_scrollbar = ttk.Scrollbar(tab_container, orient=tk.VERTICAL, command=tab_canvas.yview)
            tab_canvas.configure(yscrollcommand=tab_scrollbar.set)
            
            tab_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            tab_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Create content frame inside canvas
            content_frame = ttk.Frame(tab_canvas, padding=5)
            canvas_window = tab_canvas.create_window((0, 0), window=content_frame, anchor="nw")
            
            # Configure scroll region
            def on_content_configure(event):
                tab_canvas.configure(scrollregion=tab_canvas.bbox("all"))
            content_frame.bind("<Configure>", on_content_configure)
            
            # Adjust canvas window width to match canvas
            def on_canvas_configure(event):
                tab_canvas.itemconfig(canvas_window, width=event.width)
            tab_canvas.bind("<Configure>", on_canvas_configure)
            
            # Mouse wheel scrolling - bind to canvas and recursively to all children
            def on_tab_mousewheel(event):
                tab_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
            def bind_to_mousewheel(widget):
                widget.bind("<MouseWheel>", on_tab_mousewheel)
                for child in widget.winfo_children():
                    bind_to_mousewheel(child)
            
            tab_canvas.bind("<MouseWheel>", on_tab_mousewheel)
            
            # Call the function to create content
            create_content_func(content_frame)
            
            # Bind mousewheel to all widgets in the content frame
            bind_to_mousewheel(content_frame)
            
            return content_frame
        
        # === Tab 1: Basic Settings (Y-axis as Listbox) ===
        create_scrollable_tab(self.settings_notebook, "Basic Settings", self.create_basic_settings_tab)

        # === Tab 2: Style Settings (★ Major Change) ===
        create_scrollable_tab(self.settings_notebook, "Style", self.create_style_settings_tab)
        
        # === Tab 3: Font Settings ===
        create_scrollable_tab(self.settings_notebook, "Font", self.create_font_size_tab)
        
        # === Tab 4: Axis & Ticks Settings ===
        create_scrollable_tab(self.settings_notebook, "Axis/Ticks", self.create_axis_ticks_tab)

        # === Tab 5: Spines & Background ===
        create_scrollable_tab(self.settings_notebook, "Spines/BG", self.create_spines_tab)

        # === Tab 6: Legend Settings ===
        create_scrollable_tab(self.settings_notebook, "Legend", self.create_legend_tab)

        # === Tab 7: Advanced Settings ===
        create_scrollable_tab(self.settings_notebook, "Advanced", self.create_advanced_tab)

        # --- Bottom section: Data Edit Area ---
        # 1. UI English: LabelFrame text
        data_frame = ttk.LabelFrame(left_paned, text="Data Editor")
        left_paned.add(data_frame, weight=3)
        
        self.sheet_frame = ttk.Frame(data_frame)
        self.sheet_frame.pack(fill=tk.BOTH, expand=True)

        # === Right Panel (Graph) ===
        # 1. UI English: LabelFrame text
        right_panel = ttk.LabelFrame(content_frame, text="Graph Preview")
        # 2. Layout: Add to PanedWindow with weight 1
        content_frame.add(right_panel, weight=1)

        def force_sash_position(event=None):
            # This needs to be done *after* the window is fully drawn and sized
            try:
                # Get the total width of the paned window
                width = content_frame.winfo_width()
                # Set the sash position to be in the middle
                # The first sash (index 0) controls the boundary between panel 0 and 1
                content_frame.sashpos(0, width // 2)
                # Unbind after first run to allow user resizing
                content_frame.unbind("<Configure>")
            except Exception as e:
                print(f"Error setting sash position: {e}")
        
        # We bind to <Configure> for the *first* time the window is configured
        content_frame.bind("<Configure>", force_sash_position)


        # Scrollbars for the graph
        self.scrollable_canvas = tk.Canvas(right_panel, borderwidth=0, highlightthickness=0)
        
        v_scroll = ttk.Scrollbar(right_panel, orient=tk.VERTICAL, command=self.scrollable_canvas.yview)
        h_scroll = ttk.Scrollbar(right_panel, orient=tk.HORIZONTAL, command=self.scrollable_canvas.xview)
        self.scrollable_canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.scrollable_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Internal frame for Matplotlib graph and toolbar
        self.graph_frame = ttk.Frame(self.scrollable_canvas)
        self.scrollable_canvas.create_window((0, 0), window=self.graph_frame, anchor="nw")

        # Graph and Toolbar master changed to self.graph_frame
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_frame)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.graph_frame)
        self.toolbar.update()
        self.toolbar.pack(side=tk.TOP, fill=tk.X) 
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.NONE, expand=False, padx=5, pady=5)

        self.graph_frame.bind("<Configure>", self.on_graph_frame_configure)

    def on_graph_frame_configure(self, event):
        """ Update scroll region when graph frame size changes """
        self.scrollable_canvas.configure(scrollregion=self.scrollable_canvas.bbox("all"))


    # --- Variable Initialization ---
    def get_font_list(self):
        try:
            font_list = sorted(list(set(fm.fontManager.get_font_names())))
            # 1. UI English: Keep font names as they are, but prioritize common ones
            common_fonts = ['sans-serif', 'serif', 'monospace', 'Arial', 'Times New Roman', 'Courier New', 'Yu Gothic', 'Meiryo', 'MS Gothic']
            for f in reversed(common_fonts):
                if f in font_list:
                    font_list.remove(f)
                    font_list.insert(0, f)
            return font_list
        except Exception as e:
            print(f"Failed to load system fonts: {e}")
            return ['sans-serif', 'serif', 'monospace', 'Arial']

    def create_all_tk_variables(self):
        # Basic
        self.plot_type_var = tk.StringVar(value="surface")
        self.x_axis_var = tk.StringVar()
        self.y_axis_var = tk.StringVar()
        self.z_axis_var = tk.StringVar()
        self.title_var = tk.StringVar()
        self.xlabel_var = tk.StringVar()
        self.ylabel_var = tk.StringVar()
        self.zlabel_var = tk.StringVar()
        
        # Style dictionary for Z-axis series
        self.y1_series_styles = {}

        # ★ 1. Consolidate: Remove separate target vars
        # self.y1_style_target_var = tk.StringVar()
        # self.y2_style_target_var = tk.StringVar()
        # ★ 1. Consolidate: Add one combined target var
        self.combined_style_target_var = tk.StringVar()

        # Add transient tk.Vars for the Style Editor widgets
        # These hold the style for the *currently selected series*
        self.current_style_color_var = tk.StringVar(value="#000000")
        self.current_style_linestyle_var = tk.StringVar(value="-")
        self.current_style_marker_var = tk.StringVar(value="o")
        self.current_style_linewidth_var = tk.DoubleVar(value=1.5)
        self.current_style_alpha_var = tk.DoubleVar(value=1.0)
        # --- (End Style Refactor) ---

        self.grid_var = tk.BooleanVar(value=False)
        self.marker_var = tk.BooleanVar(value=True) # Common Show Markers toggle

        # Font/Size
        self.font_family_var = tk.StringVar(value=self.font_list[0] if self.font_list else 'sans-serif')
        self.title_fontsize_var = tk.DoubleVar(value=16.0)
        self.xlabel_fontsize_var = tk.DoubleVar(value=14.0)
        self.ylabel_fontsize_var = tk.DoubleVar(value=14.0)
        self.zlabel_fontsize_var = tk.DoubleVar(value=14.0)
        self.tick_fontsize_var = tk.DoubleVar(value=14.0)
        self.fig_width_var = tk.DoubleVar(value=7.0)
        self.fig_height_var = tk.DoubleVar(value=6.0)
        
        # Axis/Ticks
        self.xlim_min_var = tk.StringVar()
        self.xlim_max_var = tk.StringVar()
        self.ylim_min_var = tk.StringVar()
        self.ylim_max_var = tk.StringVar()
        self.zlim_min_var = tk.StringVar()
        self.zlim_max_var = tk.StringVar()
        self.xtick_show_var = tk.BooleanVar(value=True)
        self.xtick_label_show_var = tk.BooleanVar(value=True)
        self.xtick_direction_var = tk.StringVar(value='out')
        self.ytick_show_var = tk.BooleanVar(value=True)
        self.ytick_label_show_var = tk.BooleanVar(value=True)
        self.ytick_direction_var = tk.StringVar(value='out')
        self.ztick_show_var = tk.BooleanVar(value=True)
        self.ztick_label_show_var = tk.BooleanVar(value=True)
        self.ztick_direction_var = tk.StringVar(value='out')
        
        self.xaxis_plain_format_var = tk.BooleanVar(value=False)
        self.yaxis_plain_format_var = tk.BooleanVar(value=False)
        self.zaxis_plain_format_var = tk.BooleanVar(value=False)
        
        self.xtick_major_interval_var = tk.StringVar()
        self.ytick_major_interval_var = tk.StringVar()
        self.ztick_major_interval_var = tk.StringVar()
        
        # Spines/BG
        self.spine_top_var = tk.BooleanVar(value=True)
        self.spine_bottom_var = tk.BooleanVar(value=True)
        self.spine_left_var = tk.BooleanVar(value=True)
        self.spine_right_var = tk.BooleanVar(value=True)
        self.face_color_var = tk.StringVar(value='#FFFFFF') # Axes background
        self.fig_color_var = tk.StringVar(value='#FFFFFF') # Figure background (★ ADDED)
        
        # Legend
        self.legend_show_var = tk.BooleanVar(value=False)
        self.legend_loc_var = tk.StringVar(value='best')

        # Log Scale Vars for 3D
        self.x_log_scale_var = tk.BooleanVar(value=False)
        self.y_log_scale_var = tk.BooleanVar(value=False)
        self.z_log_scale_var = tk.BooleanVar(value=False)

        self.x_invert_var = tk.BooleanVar(value=False)
        self.y_invert_var = tk.BooleanVar(value=False)
        self.z_invert_var = tk.BooleanVar(value=False)

        # Grid settings for 3D
        self.grid_alpha_var = tk.DoubleVar(value=0.3)
        
        # 3D View angles
        self.view_elev_var = tk.IntVar(value=30)
        self.view_azim_var = tk.IntVar(value=-60)
        
        # Mesh resolution for surface/wireframe
        self.mesh_resolution_var = tk.IntVar(value=50)
        
        # Colormap for surface/contour
        self.colormap_var = tk.StringVar(value='viridis')


    # --- Tab Creation Methods ---
    def create_basic_settings_tab(self, frame):
        # 1. UI English: Labels
        
        # --- Top-level settings ---
        top_settings_frame = ttk.Frame(frame)
        top_settings_frame.pack(fill=tk.X, pady=2)

        ttk.Label(top_settings_frame, text="Graph Title:").grid(row=0, column=0, padx=3, pady=2, sticky=tk.W)
        self.title_entry = ttk.Entry(top_settings_frame, textvariable=self.title_var, width=25)
        self.title_entry.grid(row=0, column=1, columnspan=3, padx=3, pady=2, sticky=tk.EW)

        ttk.Label(top_settings_frame, text="Plot Type:").grid(row=1, column=0, padx=3, pady=2, sticky=tk.W)
        self.plot_type_combo = ttk.Combobox(top_settings_frame, textvariable=self.plot_type_var, 
                                            values=["surface", "wireframe", "scatter3d", "line3d", "contour3d"], state='readonly', width=24)
        self.plot_type_combo.grid(row=1, column=1, columnspan=3, padx=3, pady=2, sticky=tk.EW)
        
        top_settings_frame.columnconfigure(1, weight=1)

        # --- Axis Settings Frames ---
        axis_frames_container = ttk.Frame(frame)
        axis_frames_container.pack(fill=tk.X, expand=False, pady=2)
        
        # --- X-Axis Frame ---
        x_axis_frame = ttk.LabelFrame(axis_frames_container, text="X-Axis")
        x_axis_frame.pack(fill=tk.X, padx=3, pady=2)
        
        ttk.Label(x_axis_frame, text="Label:").grid(row=0, column=0, padx=2, pady=1, sticky=tk.W)
        self.xlabel_entry = ttk.Entry(x_axis_frame, textvariable=self.xlabel_var, width=25)
        self.xlabel_entry.grid(row=0, column=1, columnspan=2, padx=2, pady=1, sticky=tk.EW) 
        
        ttk.Label(x_axis_frame, text="Data:").grid(row=1, column=0, padx=2, pady=1, sticky=tk.W)
        self.x_axis_combo = ttk.Combobox(x_axis_frame, textvariable=self.x_axis_var, state='disabled', width=24)
        self.x_axis_combo.grid(row=1, column=1, columnspan=2, padx=2, pady=1, sticky=tk.EW) 
        
        # (★ ADDED) Log scale checkbox
        self.x_log_scale_check = ttk.Checkbutton(x_axis_frame, text="Log Scale", variable=self.x_log_scale_var)
        self.x_log_scale_check.grid(row=2, column=0, padx=2, pady=1, sticky=tk.W) 
        
        self.x_invert_check = ttk.Checkbutton(x_axis_frame, text="Invert Axis", variable=self.x_invert_var)
        self.x_invert_check.grid(row=2, column=1, padx=5, pady=1, sticky=tk.W) # Place next to log scale
        
        x_axis_frame.columnconfigure(1, weight=1)

        # --- Y-Axis Frames (in a PanedWindow) ---
        y_axis_paned_window = ttk.PanedWindow(axis_frames_container, orient=tk.HORIZONTAL)
        y_axis_paned_window.pack(fill=tk.X, expand=False, padx=3, pady=2)

        # --- Y-Axis Frame ---
        y_axis_frame = ttk.LabelFrame(axis_frames_container, text="Y-Axis")
        y_axis_frame.pack(fill=tk.X, padx=3, pady=2)
        
        ttk.Label(y_axis_frame, text="Label:").grid(row=0, column=0, padx=2, pady=1, sticky=tk.W)
        self.ylabel_entry = ttk.Entry(y_axis_frame, textvariable=self.ylabel_var, width=25)
        self.ylabel_entry.grid(row=0, column=1, columnspan=2, padx=2, pady=1, sticky=tk.EW)
        
        ttk.Label(y_axis_frame, text="Data:").grid(row=1, column=0, padx=2, pady=1, sticky=tk.W)
        self.y_axis_combo = ttk.Combobox(y_axis_frame, textvariable=self.y_axis_var, state='disabled', width=24)
        self.y_axis_combo.grid(row=1, column=1, columnspan=2, padx=2, pady=1, sticky=tk.EW)
        
        self.y_log_scale_check = ttk.Checkbutton(y_axis_frame, text="Log Scale", variable=self.y_log_scale_var)
        self.y_log_scale_check.grid(row=2, column=0, padx=2, pady=1, sticky=tk.W)
        
        self.y_invert_check = ttk.Checkbutton(y_axis_frame, text="Invert Axis", variable=self.y_invert_var)
        self.y_invert_check.grid(row=2, column=1, padx=5, pady=1, sticky=tk.W)
        
        y_axis_frame.columnconfigure(1, weight=1)

        # --- Z-Axis Frame ---
        z_axis_frame = ttk.LabelFrame(axis_frames_container, text="Z-Axis (Value Axis)")
        z_axis_frame.pack(fill=tk.X, padx=3, pady=2)
        
        ttk.Label(z_axis_frame, text="Label (e.g., 'Temperature', 'Value'):").grid(row=0, column=0, columnspan=3, padx=2, pady=1, sticky=tk.W)
        self.zlabel_entry = ttk.Entry(z_axis_frame, textvariable=self.zlabel_var, width=25)
        self.zlabel_entry.grid(row=1, column=0, columnspan=3, padx=2, pady=1, sticky=tk.EW)
        
        ttk.Label(z_axis_frame, text="Data Series (Multi-select):").grid(row=2, column=0, columnspan=3, padx=2, pady=1, sticky=tk.W)
        
        self.z_listbox_frame = ttk.Frame(z_axis_frame, height=80)
        self.z_listbox_scroll = ttk.Scrollbar(self.z_listbox_frame, orient=tk.VERTICAL)
        self.z_listbox = tk.Listbox(self.z_listbox_frame, selectmode=tk.MULTIPLE, yscrollcommand=self.z_listbox_scroll.set, exportselection=False, height=4)
        self.z_listbox_scroll.config(command=self.z_listbox.yview)
        self.z_listbox_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.z_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.z_listbox_frame.grid(row=3, column=0, columnspan=3, padx=3, pady=2, sticky=tk.EW)
        self.z_listbox_frame.pack_propagate(False)
        
        self.z_log_scale_check = ttk.Checkbutton(z_axis_frame, text="Log Scale", variable=self.z_log_scale_var)
        self.z_log_scale_check.grid(row=4, column=0, padx=2, pady=1, sticky=tk.W)
        
        self.z_invert_check = ttk.Checkbutton(z_axis_frame, text="Invert Axis", variable=self.z_invert_var)
        self.z_invert_check.grid(row=4, column=1, padx=5, pady=1, sticky=tk.W)
        
        z_axis_frame.columnconfigure(1, weight=1)

        # --- Figure Size ---
        fig_size_frame = ttk.Frame(frame)
        fig_size_frame.pack(fill=tk.X, pady=(2,0))
        
        ttk.Label(fig_size_frame, text="Figure Width (inch):").grid(row=0, column=0, padx=2, pady=1, sticky=tk.W)
        self.fig_width_spin = ttk.Spinbox(fig_size_frame, from_=3, to=20, increment=0.5, textvariable=self.fig_width_var, width=10)
        self.fig_width_spin.grid(row=0, column=1, padx=2, pady=1, sticky=tk.W)

        ttk.Label(fig_size_frame, text="Figure Height (inch):").grid(row=0, column=2, padx=5, pady=1, sticky=tk.W)
        self.fig_height_spin = ttk.Spinbox(fig_size_frame, from_=3, to=20, increment=0.5, textvariable=self.fig_height_var, width=10)
        self.fig_height_spin.grid(row=0, column=3, padx=2, pady=1, sticky=tk.W)

    def create_style_settings_tab(self, frame):
        # 3D Style Settings
        
        common_frame = ttk.LabelFrame(frame, text="Common Settings")
        common_frame.pack(fill=tk.X, padx=5, pady=5)

        self.grid_check = ttk.Checkbutton(common_frame, text="Show Grid", variable=self.grid_var)
        self.grid_check.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        self.marker_check = ttk.Checkbutton(common_frame, text="Show Markers (line3d only)", variable=self.marker_var)
        self.marker_check.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        # Info label
        info_frame = ttk.LabelFrame(frame, text="3D Style Information")
        info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        info_text = """Style settings apply to different plot types as follows:
        
• surface: Color, Alpha (transparency)
• wireframe: Color, Line Width, Alpha
• scatter3d: Color, Marker Style, Alpha
• line3d: All settings (Color, Line Style, Line Width, Marker, Alpha)
• contour3d: Alpha only (uses colormap)

Select Z-axis series below to customize individual series styles."""
        
        ttk.Label(info_frame, text=info_text, justify=tk.LEFT).pack(padx=5, pady=5)
        
        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # --- Mesh Resolution Section ---
        mesh_frame = ttk.LabelFrame(frame, text="Surface/Wireframe Quality")
        mesh_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(mesh_frame, text="Mesh Resolution:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.mesh_resolution_spin = ttk.Spinbox(mesh_frame, from_=10, to=200, increment=10, textvariable=self.mesh_resolution_var, width=10)
        self.mesh_resolution_spin.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(mesh_frame, text="(Higher = smoother)").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        # --- Colormap Section ---
        cmap_frame = ttk.LabelFrame(frame, text="Colormap (Surface/Contour)")
        cmap_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(cmap_frame, text="Color Gradient:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # Comprehensive list of matplotlib colormaps
        colormaps = [
            # Perceptually Uniform Sequential
            'viridis', 'plasma', 'inferno', 'magma', 'cividis',
            # Sequential
            'Greys', 'Purples', 'Blues', 'Greens', 'Oranges', 'Reds',
            'YlOrBr', 'YlOrRd', 'OrRd', 'PuRd', 'RdPu', 'BuPu',
            'GnBu', 'PuBu', 'YlGnBu', 'PuBuGn', 'BuGn', 'YlGn',
            # Sequential (2)
            'binary', 'gist_yarg', 'gist_gray', 'gray', 'bone', 'pink',
            'spring', 'summer', 'autumn', 'winter', 'cool', 'Wistia',
            'hot', 'afmhot', 'gist_heat', 'copper',
            # Diverging
            'PiYG', 'PRGn', 'BrBG', 'PuOr', 'RdGy', 'RdBu',
            'RdYlBu', 'RdYlGn', 'Spectral', 'coolwarm', 'bwr', 'seismic',
            # Cyclic
            'twilight', 'twilight_shifted', 'hsv',
            # Qualitative
            'Pastel1', 'Pastel2', 'Paired', 'Accent',
            'Dark2', 'Set1', 'Set2', 'Set3',
            'tab10', 'tab20', 'tab20b', 'tab20c',
            # Miscellaneous
            'flag', 'prism', 'ocean', 'gist_earth', 'terrain', 'gist_stern',
            'gnuplot', 'gnuplot2', 'CMRmap', 'cubehelix', 'brg',
            'gist_rainbow', 'rainbow', 'jet', 'turbo', 'nipy_spectral',
            'gist_ncar'
        ]
        
        self.colormap_combo = ttk.Combobox(cmap_frame, textvariable=self.colormap_var, 
                                           values=colormaps, state='readonly', width=20)
        self.colormap_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(cmap_frame, text="(When Series Color = Auto)").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        # Selector frame
        selector_frame = ttk.Frame(frame)
        selector_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(selector_frame, text="Select Z Series to Edit:").pack(side=tk.LEFT, padx=5)
        self.style_combo = ttk.Combobox(selector_frame, textvariable=self.combined_style_target_var, state='readonly', width=25)
        self.style_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.style_combo.bind("<<ComboboxSelected>>", self.on_combined_series_select)

        # Style Editor Frame
        editor_frame = ttk.LabelFrame(frame, text="Style Editor (for selected series)")
        editor_frame.pack(fill=tk.X, padx=5, pady=5)

        # Row 0: Color
        ttk.Label(editor_frame, text="Color:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.style_editor_color_btn = ttk.Button(editor_frame, text="Choose Color", command=self.on_style_editor_color_pick, width=12)
        self.style_editor_color_btn.grid(row=0, column=1, padx=5, pady=5)
        
        self.style_editor_color_auto_btn = ttk.Button(editor_frame, text="Auto", command=self.on_style_editor_color_auto, width=8)
        self.style_editor_color_auto_btn.grid(row=0, column=2, padx=5, pady=5)
        
        self.style_editor_color_label = ttk.Label(editor_frame, text="Auto", width=15, relief=tk.SUNKEN, anchor=tk.CENTER)
        self.style_editor_color_label.grid(row=0, column=3, padx=5, pady=5)

        # Row 1: Alpha
        ttk.Label(editor_frame, text="Alpha (Transparency):").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.style_editor_alpha_spin = ttk.Spinbox(editor_frame, from_=0.0, to=1.0, increment=0.1, textvariable=self.current_style_alpha_var, width=10,
                                                   command=self.on_style_editor_change)
        self.style_editor_alpha_spin.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        self.style_editor_alpha_spin.bind("<Return>", self.on_style_editor_change)
        ttk.Label(editor_frame, text="(0.0 = transparent, 1.0 = opaque)").grid(row=1, column=2, columnspan=2, padx=5, pady=5, sticky=tk.W)

        # Row 2: Line Width
        ttk.Label(editor_frame, text="Line Width:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.style_editor_linewidth_spin = ttk.Spinbox(editor_frame, from_=0.5, to=10.0, increment=0.5, textvariable=self.current_style_linewidth_var, width=10,
                                                       command=self.on_style_editor_change)
        self.style_editor_linewidth_spin.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        self.style_editor_linewidth_spin.bind("<Return>", self.on_style_editor_change)
        ttk.Label(editor_frame, text="(for wireframe and line3d)").grid(row=2, column=2, columnspan=2, padx=5, pady=5, sticky=tk.W)

        # Row 3: Line Style
        ttk.Label(editor_frame, text="Line Style:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.style_editor_linestyle_combo = ttk.Combobox(editor_frame, textvariable=self.current_style_linestyle_var, 
                                            values=['-', '--', ':', '-.', 'None'], state='readonly', width=10)
        self.style_editor_linestyle_combo.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        self.style_editor_linestyle_combo.bind("<<ComboboxSelected>>", self.on_style_editor_change)
        ttk.Label(editor_frame, text="(for line3d only)").grid(row=3, column=2, columnspan=2, padx=5, pady=5, sticky=tk.W)

        # Row 4: Marker Style
        ttk.Label(editor_frame, text="Marker Style:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.style_editor_marker_combo = ttk.Combobox(editor_frame, textvariable=self.current_style_marker_var, 
                                              values=['o', '.', ',', 's', 'p', '*', '^', '<', '>', 'D', 'H', 'None'], state='readonly', width=10)
        self.style_editor_marker_combo.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)
        self.style_editor_marker_combo.bind("<<ComboboxSelected>>", self.on_style_editor_change)
        ttk.Label(editor_frame, text="(for scatter3d and line3d)").grid(row=4, column=2, columnspan=2, padx=5, pady=5, sticky=tk.W)


    def create_font_size_tab(self, frame):
        # 1. UI English: Labels
        ttk.Label(frame, text="Font Family:").grid(row=0, column=0, padx=2, pady=1, sticky=tk.W)
        self.font_family_combo = ttk.Combobox(frame, textvariable=self.font_family_var, 
                                              values=self.font_list, state='readonly', width=15)
        self.font_family_combo.grid(row=0, column=1, columnspan=2, padx=2, pady=1, sticky=tk.W)

        # Font Size (Left Column)
        ttk.Label(frame, text="Title Size:").grid(row=1, column=0, padx=2, pady=1, sticky=tk.W)
        self.title_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.title_fontsize_var, width=6)
        self.title_fontsize_spin.grid(row=1, column=1, padx=2, pady=1, sticky=tk.W)

        ttk.Label(frame, text="X-Label Size:").grid(row=2, column=0, padx=2, pady=1, sticky=tk.W)
        self.xlabel_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.xlabel_fontsize_var, width=6)
        self.xlabel_fontsize_spin.grid(row=2, column=1, padx=2, pady=1, sticky=tk.W)

        ttk.Label(frame, text="Y-Label Size:").grid(row=3, column=0, padx=2, pady=1, sticky=tk.W)
        self.ylabel_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.ylabel_fontsize_var, width=6)
        self.ylabel_fontsize_spin.grid(row=3, column=1, padx=2, pady=1, sticky=tk.W)

        ttk.Label(frame, text="Z-Label Size:").grid(row=4, column=0, padx=2, pady=1, sticky=tk.W)
        self.zlabel_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.zlabel_fontsize_var, width=6)
        self.zlabel_fontsize_spin.grid(row=4, column=1, padx=2, pady=1, sticky=tk.W)

        # Font Size (Right Column)
        ttk.Label(frame, text="Tick Label Size:").grid(row=1, column=2, padx=5, pady=1, sticky=tk.W)
        self.tick_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.tick_fontsize_var, width=6)
        self.tick_fontsize_spin.grid(row=1, column=3, padx=2, pady=1, sticky=tk.W)

        # 3. Layout: Figure Size is now in Basic Settings tab
        
    def create_axis_ticks_tab(self, frame):
        # 1. UI English: Labels and Checkbuttons
        # --- X-Axis ---
        ttk.Label(frame, text="X-Axis", font=("-weight bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="Range (Min):").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.xlim_min_entry = ttk.Entry(frame, textvariable=self.xlim_min_var, width=10)
        self.xlim_min_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(frame, text="Range (Max):").grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        self.xlim_max_entry = ttk.Entry(frame, textvariable=self.xlim_max_var, width=10)
        self.xlim_max_entry.grid(row=1, column=3, padx=5, pady=5)
        
        ttk.Label(frame, text="Major Tick Interval:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W) # New row 2
        self.xtick_major_interval_entry = ttk.Entry(frame, textvariable=self.xtick_major_interval_var, width=10)
        self.xtick_major_interval_entry.grid(row=2, column=1, padx=5, pady=5)
        ttk.Label(frame, text="(Linear scale only)").grid(row=2, column=2, padx=5, pady=5, sticky=tk.W) # Note

        ttk.Label(frame, text="Tick Direction:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W) 
        self.xtick_direction_combo = ttk.Combobox(frame, textvariable=self.xtick_direction_var, 
                                                  values=['out', 'in', 'inout'], state='readonly', width=8)
        self.xtick_direction_combo.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W) 
        self.xtick_show_check = ttk.Checkbutton(frame, text="Show Ticks", variable=self.xtick_show_var)
        self.xtick_show_check.grid(row=3, column=2, padx=5, pady=5, sticky=tk.W)
        self.xtick_label_show_check = ttk.Checkbutton(frame, text="Show Labels", variable=self.xtick_label_show_var)
        self.xtick_label_show_check.grid(row=3, column=3, padx=5, pady=5, sticky=tk.W) 
        
        self.xaxis_plain_format_check = ttk.Checkbutton(frame, 
            text="Disable Scientific Notation", 
            variable=self.xaxis_plain_format_var)
        self.xaxis_plain_format_check.grid(row=4, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W) 

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=5, column=0, columnspan=5, sticky="ew", pady=10)

        # --- Y-Axis ---
        ttk.Label(frame, text="Y-Axis", font=("-weight bold")).grid(row=6, column=0, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="Range (Min):").grid(row=7, column=0, padx=5, pady=5, sticky=tk.W)
        self.ylim_min_entry = ttk.Entry(frame, textvariable=self.ylim_min_var, width=10)
        self.ylim_min_entry.grid(row=7, column=1, padx=5, pady=5) # Shifted from 6 to 7
        ttk.Label(frame, text="Range (Max):").grid(row=7, column=2, padx=5, pady=5, sticky=tk.W) # Shifted from 6 to 7
        self.ylim_max_entry = ttk.Entry(frame, textvariable=self.ylim_max_var, width=10)
        self.ylim_max_entry.grid(row=7, column=3, padx=5, pady=5) # Shifted from 6 to 7
        
        ttk.Label(frame, text="Major Tick Interval:").grid(row=8, column=0, padx=5, pady=5, sticky=tk.W) # New row 8
        self.ytick_major_interval_entry = ttk.Entry(frame, textvariable=self.ytick_major_interval_var, width=10)
        self.ytick_major_interval_entry.grid(row=8, column=1, padx=5, pady=5)
        ttk.Label(frame, text="(Linear scale only)").grid(row=8, column=2, padx=5, pady=5, sticky=tk.W) # Note

        ttk.Label(frame, text="Tick Direction:").grid(row=9, column=0, padx=5, pady=5, sticky=tk.W) 
        self.ytick_direction_combo = ttk.Combobox(frame, textvariable=self.ytick_direction_var, 
                                                  values=['out', 'in', 'inout'], state='readonly', width=8)
        self.ytick_direction_combo.grid(row=9, column=1, padx=5, pady=5, sticky=tk.W) 
        self.ytick_show_check = ttk.Checkbutton(frame, text="Show Ticks", variable=self.ytick_show_var)
        self.ytick_show_check.grid(row=9, column=2, padx=5, pady=5, sticky=tk.W) 
        self.ytick_label_show_check = ttk.Checkbutton(frame, text="Show Labels", variable=self.ytick_label_show_var)
        self.ytick_label_show_check.grid(row=9, column=3, padx=5, pady=5, sticky=tk.W) 


        self.yaxis_plain_format_check = ttk.Checkbutton(frame, 
            text="Disable Scientific Notation", 
            variable=self.yaxis_plain_format_var)
        self.yaxis_plain_format_check.grid(row=10, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W) 
        
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=11, column=0, columnspan=5, sticky="ew", pady=10)
        
        # --- Z-Axis ---
        ttk.Label(frame, text="Z-Axis", font=("-weight bold")).grid(row=12, column=0, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="Range (Min):").grid(row=13, column=0, padx=5, pady=5, sticky=tk.W)
        self.zlim_min_entry = ttk.Entry(frame, textvariable=self.zlim_min_var, width=10)
        self.zlim_min_entry.grid(row=13, column=1, padx=5, pady=5)
        ttk.Label(frame, text="Range (Max):").grid(row=13, column=2, padx=5, pady=5, sticky=tk.W)
        self.zlim_max_entry = ttk.Entry(frame, textvariable=self.zlim_max_var, width=10)
        self.zlim_max_entry.grid(row=13, column=3, padx=5, pady=5)
        
        ttk.Label(frame, text="Major Tick Interval:").grid(row=14, column=0, padx=5, pady=5, sticky=tk.W)
        self.ztick_major_interval_entry = ttk.Entry(frame, textvariable=self.ztick_major_interval_var, width=10)
        self.ztick_major_interval_entry.grid(row=14, column=1, padx=5, pady=5)
        ttk.Label(frame, text="(Linear scale only)").grid(row=14, column=2, padx=5, pady=5, sticky=tk.W)

        ttk.Label(frame, text="Tick Direction:").grid(row=15, column=0, padx=5, pady=5, sticky=tk.W) 
        self.ztick_direction_combo = ttk.Combobox(frame, textvariable=self.ztick_direction_var, 
                                                  values=['out', 'in', 'inout'], state='readonly', width=8)
        self.ztick_direction_combo.grid(row=15, column=1, padx=5, pady=5, sticky=tk.W)
        self.ztick_show_check = ttk.Checkbutton(frame, text="Show Ticks", variable=self.ztick_show_var)
        self.ztick_show_check.grid(row=15, column=2, padx=5, pady=5, sticky=tk.W) 
        self.ztick_label_show_check = ttk.Checkbutton(frame, text="Show Labels", variable=self.ztick_label_show_var)
        self.ztick_label_show_check.grid(row=15, column=3, padx=5, pady=5, sticky=tk.W)
        

        self.zaxis_plain_format_check = ttk.Checkbutton(frame, 
            text="Disable Scientific Notation", 
            variable=self.zaxis_plain_format_var)

        self.zaxis_plain_format_check.grid(row=16, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W) 

        frame.columnconfigure(4, weight=1)

    def create_spines_tab(self, frame):
        # 1. UI English: Labels, Checkbuttons, Button
        ttk.Label(frame, text="Show Graph Spines:").grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=5)
        
        self.spine_top_check = ttk.Checkbutton(frame, text="Top", variable=self.spine_top_var)
        self.spine_top_check.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.spine_bottom_check = ttk.Checkbutton(frame, text="Bottom", variable=self.spine_bottom_var)
        self.spine_bottom_check.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        self.spine_left_check = ttk.Checkbutton(frame, text="Left", variable=self.spine_left_var)
        self.spine_left_check.grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        
        self.spine_right_check = ttk.Checkbutton(frame, text="Right", variable=self.spine_right_var)
        self.spine_right_check.grid(row=1, column=3, padx=5, pady=5, sticky=tk.W)
        
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=2, column=0, columnspan=4, sticky="ew", pady=10)
        
        ttk.Label(frame, text="Graph Background Color (Axes):").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.face_color_btn = ttk.Button(frame, text="Select", command=lambda: self.choose_color(self.face_color_var, self.face_color_label))
        self.face_color_btn.grid(row=3, column=1, padx=5, pady=5)
        self.face_color_label = ttk.Label(frame, text=self.face_color_var.get(), background=self.face_color_var.get(), width=10)
        self.face_color_label.grid(row=3, column=2, padx=5, pady=5)

        # (★ ADDED) Figure Background Color
        ttk.Label(frame, text="Figure Background Color (Outside):").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.fig_color_btn = ttk.Button(frame, text="Select", command=lambda: self.choose_color(self.fig_color_var, self.fig_color_label))
        self.fig_color_btn.grid(row=4, column=1, padx=5, pady=5)
        self.fig_color_label = ttk.Label(frame, text=self.fig_color_var.get(), background=self.fig_color_var.get(), width=10)
        self.fig_color_label.grid(row=4, column=2, padx=5, pady=5)


    def create_legend_tab(self, frame):
        # 1. UI English: Labels, Checkbutton
        self.legend_show_check = ttk.Checkbutton(frame, text="Show Legend (Auto-combined)", variable=self.legend_show_var)
        self.legend_show_check.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(frame, text="Legend Location:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.legend_loc_combo = ttk.Combobox(frame, textvariable=self.legend_loc_var,
                                             values=['best', 'upper right', 'upper left', 'lower left', 
                                                     'lower right', 'right', 'center left', 'center right', 
                                                     'lower center', 'upper center', 'center'],
                                             state='readonly', width=28)
        self.legend_loc_combo.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(frame, text="(Legend labels use Z-axis column names automatically)").grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)

    def create_advanced_tab(self, frame):
        # Advanced 3D Settings Tab
        
        # --- Grid Style Section ---
        grid_style_frame = ttk.LabelFrame(frame, text="Grid Customization")
        grid_style_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(grid_style_frame, text="Grid Alpha (Transparency):").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.grid_alpha_spin = ttk.Spinbox(grid_style_frame, from_=0.0, to=1.0, increment=0.1, textvariable=self.grid_alpha_var, width=10)
        self.grid_alpha_spin.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # --- 3D View Angle Section ---
        view_frame = ttk.LabelFrame(frame, text="3D View Angle")
        view_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(view_frame, text="Elevation (degrees):").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.view_elev_spin = ttk.Spinbox(view_frame, from_=-90, to=90, increment=5, textvariable=self.view_elev_var, width=10)
        self.view_elev_spin.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(view_frame, text="Azimuth (degrees):").grid(row=0, column=2, padx=15, pady=5, sticky=tk.W)
        self.view_azim_spin = ttk.Spinbox(view_frame, from_=-180, to=180, increment=5, textvariable=self.view_azim_var, width=10)
        self.view_azim_spin.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(view_frame, text="(Adjust the 3D viewing angle)").grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W)


    
    def export_filtered_data(self):
        """Export the current (filtered) data to CSV"""
        if self.df is None or self.df.empty:
            messagebox.showinfo("Info", "No data to export.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Export Filtered Data",
            filetypes=[("CSV files", "*.csv")],
            defaultextension=".csv"
        )
        
        if not file_path:
            return
        
        try:
            self.df.to_csv(file_path, index=False)
            messagebox.showinfo("Success", f"Data exported to {file_path}.")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export data:\n{e}")
    
    def clear_all(self):
        """Clear all data and reset the application"""
        # Confirm with user
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all data and settings?\nThis cannot be undone."):
            # Clear data
            self.df = None
            self.data_file_path = ""
            self.current_project_path = ""
            
            # Clear sheet
            if self.sheet:
                self.sheet.destroy()
                self.sheet = None
            
            # Clear graph
            self.fig.clear()
            self.ax = self.fig.add_subplot(111)
            self.ax2 = None
            self.canvas.draw()
            
            # Reset listboxes
            self.z_listbox.delete(0, tk.END)
            
            # Reset combos
            self.x_axis_combo['values'] = []
            self.x_axis_combo['state'] = 'disabled'
            self.y_axis_combo['values'] = []
            self.y_axis_combo['state'] = 'disabled'
            if hasattr(self, 'errorbar_column_combo'):
                self.errorbar_column_combo['values'] = []
            if hasattr(self, 'filter_column_combo'):
                self.filter_column_combo['values'] = []
            if hasattr(self, 'style_combo'):
                self.style_combo['values'] = []
            
            # Reset all variables to defaults
            self.x_axis_var.set("")
            self.y_axis_var.set("")
            self.z_axis_var.set("")
            self.title_var.set("")
            self.xlabel_var.set("")
            self.ylabel_var.set("")
            self.zlabel_var.set("")
            
            # Clear style dictionary
            self.y1_series_styles = {}
            
            # Reset style editor
            self.combined_style_target_var.set("")
            self.load_style_to_editor(None, True)
            
            # Disable buttons
            self.plot_button['state'] = 'disabled'
            self.export_button['state'] = 'disabled'
            self.export_data_button['state'] = 'disabled'
            self.overwrite_save_button['state'] = 'disabled'
    
    def reset_settings(self):
        """Reset all settings to default values (keep data)"""
        if messagebox.askyesno("Confirm Reset", "Reset all graph settings to default values?\n(Data will be kept)"):
            # Reset all setting variables to defaults
            self.plot_type_var.set("surface")
            self.title_var.set("")
            self.xlabel_var.set("")
            self.ylabel_var.set("")
            self.zlabel_var.set("")
            
            # Clear style dictionary
            self.y1_series_styles = {}
            self.combined_style_target_var.set("")
            self.load_style_to_editor(None, True)
            
            # Reset style/appearance settings
            self.grid_var.set(False)
            self.marker_var.set(True)
            
            # Reset font settings
            self.font_family_var.set(self.font_list[0] if self.font_list else 'sans-serif')
            self.title_fontsize_var.set(16.0)
            self.xlabel_fontsize_var.set(14.0)
            self.ylabel_fontsize_var.set(14.0)
            self.zlabel_fontsize_var.set(14.0)
            self.tick_fontsize_var.set(14.0)
            self.fig_width_var.set(7.0)
            self.fig_height_var.set(6.0)
            
            # Reset axis limits
            self.xlim_min_var.set("")
            self.xlim_max_var.set("")
            self.ylim_min_var.set("")
            self.ylim_max_var.set("")
            self.zlim_min_var.set("")
            self.zlim_max_var.set("")
            
            # Reset tick settings
            self.xtick_show_var.set(True)
            self.xtick_label_show_var.set(True)
            self.xtick_direction_var.set('out')
            self.ytick_show_var.set(True)
            self.ytick_label_show_var.set(True)
            self.ytick_direction_var.set('out')
            self.ztick_show_var.set(True)
            self.ztick_label_show_var.set(True)
            self.ztick_direction_var.set('out')
            
            self.xaxis_plain_format_var.set(False)
            self.yaxis1_plain_format_var.set(False)
            self.zaxis_plain_format_var.set(False)
            
            self.xtick_major_interval_var.set("")
            self.ytick_major_interval_var.set("")
            self.ztick_major_interval_var.set("")
            
            # Reset spine settings
            self.spine_top_var.set(True)
            self.spine_bottom_var.set(True)
            self.spine_left_var.set(True)
            self.spine_right_var.set(True)
            self.face_color_var.set('#FFFFFF')
            self.fig_color_var.set('#FFFFFF')
            self.update_color_label(self.face_color_label, '#FFFFFF')
            self.update_color_label(self.fig_color_label, '#FFFFFF')
            
            # Reset legend
            self.legend_show_var.set(False)
            self.legend_loc_var.set('best')
            
            # Reset log scale and invert
            self.x_log_scale_var.set(False)
            self.y_log_scale_var.set(False)
            self.z_log_scale_var.set(False)
            self.x_invert_var.set(False)
            self.y_invert_var.set(False)
            self.z_invert_var.set(False)
            
            # Reset 3D-specific advanced features
            if hasattr(self, 'view_elev_var'):
                self.view_elev_var.set(30)
            if hasattr(self, 'view_azim_var'):
                self.view_azim_var.set(-60)
            if hasattr(self, 'grid_alpha_var'):
                self.grid_alpha_var.set(0.3)
            if hasattr(self, 'mesh_resolution_var'):
                self.mesh_resolution_var.set(50)
            if hasattr(self, 'colormap_var'):
                self.colormap_var.set('viridis')
            
            messagebox.showinfo("Reset Complete", "All settings have been reset to default values.")
    
    # --- (★ Style Refactor) Callbacks for Style Editor ---
    
    # ★ 1. Consolidate: Remove on_y1_series_select and on_y2_series_select
    # def on_y1_series_select(self, event=None): ... (removed)
    # def on_y2_series_select(self, event=None): ... (removed)

    # ★ 1. Consolidate: Add new combined callback
    def on_combined_series_select(self, event=None):
        """ Loads the style for the selected Z series into the editor """
        selected_item = self.combined_style_target_var.get()
        if not selected_item:
            return

        is_y1 = True  # In 3D mode, all Z series use y1_series_styles
        series_name = ""

        if selected_item.startswith("(Z) "):
            is_y1 = True
            series_name = selected_item[4:]  # Get name after "(Z) "
        else:
            return  # Should not happen

        if series_name:
            self.load_style_to_editor(series_name, is_y1=is_y1)

            
    def load_style_to_editor(self, series_name, is_y1):
        """ Helper to load a series' style into the 'current_style' vars """
        if series_name is None:
            self.current_style_color_var.set("#000000")
            self.current_style_linestyle_var.set("-")
            self.current_style_marker_var.set("o")
            self.current_style_linewidth_var.set(1.5)
            self.current_style_alpha_var.set(1.0)
            self.update_color_label(self.style_editor_color_label, "#000000")
            return

        # In 3D mode, all Z series use y1_series_styles
        styles_dict = self.y1_series_styles
        
        # Get the style for this series, or create a default if it's new
        series_style = self.get_or_create_default_style(series_name, styles_dict)
        
        # Set the editor's tk.Vars
        self.current_style_color_var.set(series_style.get('color', 'None')) 
        self.current_style_linestyle_var.set(series_style.get('linestyle', '-'))
        self.current_style_marker_var.set(series_style.get('marker', 'o'))
        self.current_style_linewidth_var.set(series_style.get('linewidth', 1.5))
        self.current_style_alpha_var.set(series_style.get('alpha', 1.0))
        
        # Update the color label
        self.update_color_label(self.style_editor_color_label, self.current_style_color_var.get())
        
    def get_or_create_default_style(self, series_name, styles_dict):
        """
        Gets the style dict for a series_name.
        If it doesn't exist, creates a default, saves it, and returns it.
        """
        if series_name not in styles_dict:
            # Create a default style. 
            # Note: We set color to 'None' to let matplotlib auto-color by default
            styles_dict[series_name] = {
                'color': None, # Let matplotlib auto-color initially
                'linestyle': '-',
                'marker': 'o',
                'linewidth': 1.5,
                'alpha': 1.0
            }
        return styles_dict[series_name]
        
    def on_style_editor_change(self, event=None):
        """
        Saves the current editor values back to the
        style dictionary for the currently selected series.
        """
        # ★ 1. Consolidate: Determine selected series from the combined var
        # Determine which series is selected (Y1 or Y2)
        selected_item = self.combined_style_target_var.get()
        if not selected_item:
            return # No series selected, do nothing

        series_name = None
        styles_dict = None

        if selected_item.startswith("(Z) "):
            series_name = selected_item[4:]
            styles_dict = self.y1_series_styles
        else:
            return # Should not happen

        # Ensure the style dict entry exists
        if series_name not in styles_dict:
            styles_dict[series_name] = {}
            
        # Save the current editor values into the dictionary
        try:
            styles_dict[series_name]['color'] = self.current_style_color_var.get()
            styles_dict[series_name]['linestyle'] = self.current_style_linestyle_var.get()
            styles_dict[series_name]['marker'] = self.current_style_marker_var.get()
            styles_dict[series_name]['linewidth'] = self.current_style_linewidth_var.get()
            styles_dict[series_name]['alpha'] = self.current_style_alpha_var.get()
        except tk.TclError as e:
            print(f"Error updating style (likely invalid spinbox value): {e}")

    def on_style_editor_color_pick(self):
        """ Handles the color chooser button press """
        # Get the currently selected series' color
        initial_color = self.current_style_color_var.get()
        if not initial_color or initial_color == 'None':
             initial_color = '#000000'
             
        color_code = colorchooser.askcolor(title="Choose Color", initialcolor=initial_color)[1]
        
        if color_code:
            # Set the editor var
            self.current_style_color_var.set(color_code)
            # Update the label
            self.update_color_label(self.style_editor_color_label, color_code)
            # Trigger a save
            self.on_style_editor_change()

    # (★ ADDED) Callback for Auto Color button
    def on_style_editor_color_auto(self):
        """ Handles the Auto color button press """
        # Set the editor var to 'None' (string)
        self.current_style_color_var.set('None')
        # Update the label to show "Auto"
        self.update_color_label(self.style_editor_color_label, 'None')
        # Trigger a save
        self.on_style_editor_change()

    # --- Core Function Methods ---

    def choose_color(self, color_var, color_label):
        # 1. UI English: Dialog title
        # This is for the *other* color pickers (e.g., background)
        color_code = colorchooser.askcolor(title="Choose Color", initialcolor=color_var.get())[1]
        if color_code:
            color_var.set(color_code)
            color_label.config(background=color_code, text=color_code)

    def load_data(self, file_path=None):
        if file_path is None:
            # 1. UI English: Dialog title
            file_path = filedialog.askopenfilename(
                title="Select Data File",
                filetypes=[("CSV files", "*.csv"),
                           ("Excel files", "*.xlsx *.xls"),
                           ("All files", "*.*")]
            )
        if not file_path:
            return

        try:
            if file_path.endswith('.csv'):
                self.df = pd.read_csv(file_path, dtype=str)
            else:
                self.df = pd.read_excel(file_path, dtype=str)
            
            self.data_file_path = file_path # Store path
        except Exception as e:
            # 1. UI English: Messagebox
            messagebox.showerror("Load Error", f"Failed to read file:\n{e}")
            self.data_file_path = ""
            return
            
        self.df.fillna("", inplace=True)

        if self.sheet:
            self.sheet.destroy()

        self.sheet = Sheet(self.sheet_frame,
                           data=self.df.values.tolist(),
                           headers=self.df.columns.tolist(),
                           show_toolbar=True,
                           show_top_left=True)
        self.sheet.enable_bindings()
        self.sheet.pack(fill=tk.BOTH, expand=True)
        
        self.update_plot_options()

    def update_plot_options(self):
        if self.df is None:
            return
        
        columns = self.df.columns.tolist()
        
        # X-Axis
        self.x_axis_combo['values'] = columns
        self.x_axis_combo['state'] = 'readonly'
        
        # Y-Axis
        self.y_axis_combo['values'] = columns
        self.y_axis_combo['state'] = 'readonly'
        
        # Z-Axis (Listbox)
        self.z_listbox.delete(0, tk.END)
        for col in columns:
            self.z_listbox.insert(tk.END, col)
            
        self.plot_button['state'] = 'normal'
        self.export_button['state'] = 'normal'

        # Default selection
        if columns:
            self.x_axis_var.set(columns[0])
            self.xlabel_var.set(columns[0])
            if len(columns) > 1:
                self.y_axis_var.set(columns[1])
                self.ylabel_var.set(columns[1])
            # Don't set zlabel automatically - let user define what Z represents
            # Z-axis represents values, not a coordinate dimension
            if len(columns) > 2:
                self.z_listbox.select_set(2)
            else:
                self.z_listbox.select_set(0)
        
        # Enable export data button
        if hasattr(self, 'export_data_button'):
            self.export_data_button['state'] = 'normal'

    def get_data_from_sheet(self):
        """Get data from tksheet as a DataFrame (v4.3 modified)"""
        if not self.sheet or self.df is None:
            return

        try:
            data = None
            # tksheet v7+
            if hasattr(self.sheet, 'get_sheet_data') and callable(self.sheet.get_sheet_data):
                data = self.sheet.get_sheet_data()
            # tksheet v6
            elif hasattr(self.sheet, 'data') and isinstance(self.sheet.data, list):
                data = self.sheet.data
            else:
                # 1. UI English: Messagebox
                messagebox.showerror("Compatibility Error", "Could not retrieve data from Sheet object.")
                return

            headers = None
            # tksheet v7+
            if hasattr(self.sheet, 'get_headers') and callable(self.sheet.get_headers):
                headers = self.sheet.get_headers()
            # tksheet v6
            elif hasattr(self.sheet, 'headers') and isinstance(self.sheet.headers, list):
                headers = self.sheet.headers
            # tksheet v6 (if headers is a property method)
            elif hasattr(self.sheet, 'headers') and callable(self.sheet.headers):
                 headers = self.sheet.headers()
            else:
                # 1. UI English: Messagebox
                messagebox.showerror("Compatibility Error", "Could not retrieve headers from Sheet object.")
                return
            
            # (Important) Filter data correctly as tksheet might return empty rows
            header_len = len(headers)
            if data and data[0] and len(data[0]) != header_len:
                # Slice data to match header count
                data = [row[:header_len] for row in data]

            # Create DataFrame, replace empty/whitespace strings with pd.NA, drop all-NA rows
            temp_df = pd.DataFrame(data, columns=headers).astype(str)
            temp_df.replace(r'^\s*$', pd.NA, regex=True, inplace=True)
            temp_df.dropna(how='all', inplace=True)
            # (Important) Fill NA back to '' for subsequent processing
            self.df = temp_df.fillna("")
            
        except Exception as e:
            # 1. UI English: Messagebox
            messagebox.showwarning("Data Retrieval Error", f"Failed to update data from sheet:\n{e}\n{type(e)}")
            print(f"Data Retrieval Error: {e}")
            
        # ★ 3. Debug: Remove debug prints


    # --- JSON Save/Load ---
    def save_settings(self):
        """Save all current settings and data to a .pmggrp file"""
        # 1. UI English: Dialog title
        file_path = filedialog.asksaveasfilename(
            title="Save Project",
            filetypes=[("Matplotlib Graphing App Graph Project", "*.pmggrp")],
            defaultextension=".pmggrp"
        )
        if not file_path:
            return

        # Get current edited data from sheet
        self.get_data_from_sheet()
        
        # Convert DataFrame to list format for JSON serialization
        data_dict = None
        if self.df is not None and not self.df.empty:
            data_dict = {
                "columns": self.df.columns.tolist(),
                "data": self.df.values.tolist()
            }
        
        from datetime import datetime
        
        settings = {
            "format": "Python Matplotlib Grapher App (HYGrapher) Graph Project",
            "version": "1.0",
            "application": "HYGrapher 3D",
            "application_version": VERSION,
            "dimension": "3D",
            "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "edited_data": data_dict,
            "original_file_path": self.data_file_path,
            "plot_type": self.plot_type_var.get(),
            "x_axis": self.x_axis_var.get(),
            "y_axis": self.y_axis_var.get(),
            "z_axis_indices": list(self.z_listbox.curselection()),
            "title": self.title_var.get(),
            "xlabel": self.xlabel_var.get(),
            "ylabel": self.ylabel_var.get(),
            "zlabel": self.zlabel_var.get(),
            
            # Save the series style dictionary
            "y1_series_styles": self.y1_series_styles,

            "grid": self.grid_var.get(),
            "marker": self.marker_var.get(),
            
            "font_family": self.font_family_var.get(),
            "title_fontsize": self.title_fontsize_var.get(),
            "xlabel_fontsize": self.xlabel_fontsize_var.get(),
            "ylabel_fontsize": self.ylabel_fontsize_var.get(),
            "zlabel_fontsize": self.zlabel_fontsize_var.get(),
            "tick_fontsize": self.tick_fontsize_var.get(),
            "fig_width": self.fig_width_var.get(),
            "fig_height": self.fig_height_var.get(),
            
            "xlim_min": self.xlim_min_var.get(),
            "xlim_max": self.xlim_max_var.get(),
            "ylim_min": self.ylim_min_var.get(),
            "ylim_max": self.ylim_max_var.get(),
            "zlim_min": self.zlim_min_var.get(),
            "zlim_max": self.zlim_max_var.get(),
            "xtick_show": self.xtick_show_var.get(),
            "xtick_label_show": self.xtick_label_show_var.get(),
            "xtick_direction": self.xtick_direction_var.get(),
            "ytick_show": self.ytick_show_var.get(),
            "ytick_label_show": self.ytick_label_show_var.get(),
            "ytick_direction": self.ytick_direction_var.get(),
            "ztick_show": self.ztick_show_var.get(),
            "ztick_label_show": self.ztick_label_show_var.get(),
            "ztick_direction": self.ztick_direction_var.get(),
            
            "xaxis_plain_format": self.xaxis_plain_format_var.get(),
            "yaxis_plain_format": self.yaxis_plain_format_var.get(),
            "zaxis_plain_format": self.zaxis_plain_format_var.get(),
            
            "xtick_major_interval": self.xtick_major_interval_var.get(),
            "ytick_major_interval": self.ytick_major_interval_var.get(),
            "ztick_major_interval": self.ztick_major_interval_var.get(),
            
            "spine_top": self.spine_top_var.get(),
            "spine_bottom": self.spine_bottom_var.get(),
            "spine_left": self.spine_left_var.get(),
            "spine_right": self.spine_right_var.get(),
            "face_color": self.face_color_var.get(),
            "fig_color": self.fig_color_var.get(),
            
            "legend_show": self.legend_show_var.get(),
            "legend_loc": self.legend_loc_var.get(),
            
            "x_log_scale": self.x_log_scale_var.get(),
            "y_log_scale": self.y_log_scale_var.get(),
            "z_log_scale": self.z_log_scale_var.get(),

            "x_invert": self.x_invert_var.get(),
            "y_invert": self.y_invert_var.get(),
            "z_invert": self.z_invert_var.get(),
            
            "grid_alpha": self.grid_alpha_var.get(),
            "view_elev": self.view_elev_var.get(),
            "view_azim": self.view_azim_var.get(),
            "mesh_resolution": self.mesh_resolution_var.get(),
            "colormap": self.colormap_var.get(),
        }
        
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
            # Store current project path and enable overwrite save button
            self.current_project_path = file_path
            self.overwrite_save_button['state'] = 'normal'
            # 1. UI English: Messagebox
            messagebox.showinfo("Success", f"Project saved to {file_path}.")
        except Exception as e:
            # 1. UI English: Messagebox
            messagebox.showerror("Save Error", f"Failed to save project:\n{e}")

    def overwrite_save(self):
        """Overwrite save to current project file (Ctrl+S)"""
        if not self.current_project_path:
            # No current project, do Save As
            self.save_settings()
            return
        
        # Get current edited data from sheet
        self.get_data_from_sheet()
        
        # Convert DataFrame to list format for JSON serialization
        data_dict = None
        if self.df is not None and not self.df.empty:
            data_dict = {
                "columns": self.df.columns.tolist(),
                "data": self.df.values.tolist()
            }
        
        from datetime import datetime
        
        settings = {
            "format": "Python Matplotlib Grapher App (HYGrapher) Graph Project",
            "version": "1.0",
            "application": "HYGrapher 3D",
            "application_version": VERSION,
            "dimension": "3D",
            "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "edited_data": data_dict,
            "original_file_path": self.data_file_path,
            "plot_type": self.plot_type_var.get(),
            "x_axis": self.x_axis_var.get(),
            "y_axis": self.y_axis_var.get(),
            "z_axis_indices": list(self.z_listbox.curselection()),
            "title": self.title_var.get(),
            "xlabel": self.xlabel_var.get(),
            "ylabel": self.ylabel_var.get(),
            "zlabel": self.zlabel_var.get(),
            
            "y1_series_styles": self.y1_series_styles,

            "grid": self.grid_var.get(),
            "marker": self.marker_var.get(),
            
            "font_family": self.font_family_var.get(),
            "title_fontsize": self.title_fontsize_var.get(),
            "xlabel_fontsize": self.xlabel_fontsize_var.get(),
            "ylabel_fontsize": self.ylabel_fontsize_var.get(),
            "zlabel_fontsize": self.zlabel_fontsize_var.get(),
            "tick_fontsize": self.tick_fontsize_var.get(),
            "fig_width": self.fig_width_var.get(),
            "fig_height": self.fig_height_var.get(),
            
            "xlim_min": self.xlim_min_var.get(),
            "xlim_max": self.xlim_max_var.get(),
            "ylim_min": self.ylim_min_var.get(),
            "ylim_max": self.ylim_max_var.get(),
            "zlim_min": self.zlim_min_var.get(),
            "zlim_max": self.zlim_max_var.get(),
            "xtick_show": self.xtick_show_var.get(),
            "xtick_label_show": self.xtick_label_show_var.get(),
            "xtick_direction": self.xtick_direction_var.get(),
            "ytick_show": self.ytick_show_var.get(),
            "ytick_label_show": self.ytick_label_show_var.get(),
            "ytick_direction": self.ytick_direction_var.get(),
            "ztick_show": self.ztick_show_var.get(),
            "ztick_label_show": self.ztick_label_show_var.get(),
            "ztick_direction": self.ztick_direction_var.get(),
            
            "xaxis_plain_format": self.xaxis_plain_format_var.get(),
            "yaxis_plain_format": self.yaxis_plain_format_var.get(),
            "zaxis_plain_format": self.zaxis_plain_format_var.get(),
            
            "xtick_major_interval": self.xtick_major_interval_var.get(),
            "ytick_major_interval": self.ytick_major_interval_var.get(),
            "ztick_major_interval": self.ztick_major_interval_var.get(),
            
            "spine_top": self.spine_top_var.get(),
            "spine_bottom": self.spine_bottom_var.get(),
            "spine_left": self.spine_left_var.get(),
            "spine_right": self.spine_right_var.get(),
            "face_color": self.face_color_var.get(),
            "fig_color": self.fig_color_var.get(),
            
            "legend_show": self.legend_show_var.get(),
            "legend_loc": self.legend_loc_var.get(),
            
            "x_log_scale": self.x_log_scale_var.get(),
            "y_log_scale": self.y_log_scale_var.get(),
            "z_log_scale": self.z_log_scale_var.get(),

            "x_invert": self.x_invert_var.get(),
            "y_invert": self.y_invert_var.get(),
            "z_invert": self.z_invert_var.get(),
            
            "grid_alpha": self.grid_alpha_var.get(),
            "view_elev": self.view_elev_var.get(),
            "view_azim": self.view_azim_var.get(),
            "mesh_resolution": self.mesh_resolution_var.get(),
            "colormap": self.colormap_var.get(),
        }
        
        try:
            with open(self.current_project_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
            messagebox.showinfo("Success", f"Project saved to {self.current_project_path}.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save project:\n{e}")

    def load_settings(self):
        """Load settings and data from a .pmggrp file"""
        # 1. UI English: Dialog title
        file_path = filedialog.askopenfilename(
            title="Select Project File",
            filetypes=[("Matplotlib Graph Project", "*.pmggrp")]
        )
        if not file_path:
            return
        
        self.load_project_file(file_path)
    def set_variable_from_dict(self, var, settings_dict, key, fallback_key=None):
        """ Set tk.Variable from a settings dictionary key, with existence check and fallback """
        value_to_set = None
        if key in settings_dict:
            value_to_set = settings_dict[key]
        elif fallback_key and fallback_key in settings_dict:
            value_to_set = settings_dict[fallback_key]
        
        if value_to_set is not None:
            try:
                var.set(value_to_set)
            except Exception as e:
                # 1. UI English: Warning message
                print(f"Warning: Failed to set variable for key '{key}' (fallback '{fallback_key}'): {e}")
                
    def update_color_label(self, label, color_code):
        if not color_code or color_code == 'None':
             color_code = "#FFFFFF" # Show white for 'None'
             text = "Auto"
        else:
             text = color_code
             
        try:
            label.config(background=color_code, text=text, anchor=tk.CENTER)
        except tk.TclError: # Handle invalid color code
            label.config(background="#FFFFFF", text="Invalid", anchor=tk.CENTER)

    # --- Plotting Method (3D Version) ---
    def plot_graph(self):
        
        # (1) Clear figure first
        try:
            self.fig.clear()
            self.ax = self.fig.add_subplot(111, projection='3d')
            self.ax2 = None
            self.canvas.draw()
        except Exception as e:
            messagebox.showerror("Internal Error", f"Failed to clear graph:\n{e}")
            return

        # 1. Get Data
        self.get_data_from_sheet()
        if self.df is None or self.df.empty:
            messagebox.showinfo("Info", "No data to plot.")
            self.canvas.draw()
            return

        # 2. Get All Settings
        x_col = self.x_axis_var.get()
        y_col = self.y_axis_var.get()
        plot_type = self.plot_type_var.get()
        
        try:
            z_cols = [self.z_listbox.get(i) for i in self.z_listbox.curselection()]
        except tk.TclError: 
            z_cols = []
        
        # Update Style Comboboxes
        combined_list = [f"(Z) {col}" for col in z_cols]
        self.style_combo['values'] = combined_list
        
        current_selection = self.combined_style_target_var.get()
        if current_selection not in combined_list:
            self.combined_style_target_var.set("")
            self.load_style_to_editor(None, True)

        # Column Existence Check
        if not x_col:
            messagebox.showerror("Error", "Please select an X-axis column.")
            self.canvas.draw()
            return
        if not y_col:
            messagebox.showerror("Error", "Please select a Y-axis column.")
            self.canvas.draw()
            return
        if not z_cols:
            messagebox.showerror("Error", "Please select data for Z-Axis.")
            self.canvas.draw()
            return
        
        all_cols = [x_col, y_col] + z_cols
        for col in all_cols:
            if col not in self.df.columns:
                messagebox.showerror("Plot Error", f"Selected column '{col}' does not exist in data.")
                self.canvas.draw()
                return

        # 3. Prepare Graph (Size & Font)
        try:
            self.fig.set_size_inches(self.fig_width_var.get(), self.fig_height_var.get())
            self.fig.set_facecolor(self.fig_color_var.get())
            matplotlib.rcParams['font.family'] = self.font_family_var.get()
        except Exception as e:
            messagebox.showerror("Settings Error", f"Failed to apply basic graph settings:\n{e}")
            return

        # 4. Plot 3D Graph
        try:
            # Get and clean data
            x_data_raw = self.df[x_col]
            y_data_raw = self.df[y_col]
            
            x_data_cleaned = x_data_raw.astype(str).str.replace(r'[^\d.-]', '', regex=True)
            x_data_numeric = pd.to_numeric(x_data_cleaned, errors='coerce')
            
            y_data_cleaned = y_data_raw.astype(str).str.replace(r'[^\d.-]', '', regex=True)
            y_data_numeric = pd.to_numeric(y_data_cleaned, errors='coerce')

            # Plot each Z column
            for z_idx, z_col in enumerate(z_cols):
                # Get style
                series_style = self.get_or_create_default_style(z_col, self.y1_series_styles)
                
                # Clean Z data
                z_data_cleaned = self.df[z_col].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                z_data_numeric = pd.to_numeric(z_data_cleaned, errors='coerce')
                
                # Create valid data frame
                valid_data = pd.DataFrame({
                    'x': x_data_numeric,
                    'y': y_data_numeric, 
                    'z': z_data_numeric
                }).dropna()
                
                if valid_data.empty:
                    continue
                
                # Get style properties
                color = series_style.get('color', None)
                if color == 'None' or color is None:
                    color = None
                linestyle = series_style.get('linestyle', '-')
                linewidth = series_style.get('linewidth', 1.5)
                alpha = series_style.get('alpha', 1.0)
                marker = series_style.get('marker', 'o')
                
                if plot_type == "surface" or plot_type == "wireframe":
                    # Create grid for surface/wireframe
                    try:
                        # Get mesh resolution
                        mesh_res = self.mesh_resolution_var.get()
                        
                        # Get data range
                        x_min, x_max = valid_data['x'].min(), valid_data['x'].max()
                        y_min, y_max = valid_data['y'].min(), valid_data['y'].max()
                        
                        # Create uniform grid with specified resolution
                        x_grid = np.linspace(x_min, x_max, mesh_res)
                        y_grid = np.linspace(y_min, y_max, mesh_res)
                        X, Y = np.meshgrid(x_grid, y_grid)
                        
                        # Interpolate Z values onto the grid
                        try:
                            from scipy.interpolate import griddata
                            points = valid_data[['x', 'y']].values
                            values = valid_data['z'].values
                            Z = griddata(points, values, (X, Y), method='linear')
                            
                            # Fill remaining NaN with nearest neighbor
                            if np.isnan(Z).any():
                                Z_nearest = griddata(points, values, (X, Y), method='nearest')
                                Z = np.where(np.isnan(Z), Z_nearest, Z)
                        except ImportError:
                            messagebox.showwarning("Missing Library", 
                                "scipy is required for surface/wireframe plots. Install with: pip install scipy")
                            continue
                        
                        if plot_type == "surface":
                            if color:
                                self.ax.plot_surface(X, Y, Z, alpha=alpha, label=z_col, 
                                                   color=color, edgecolor='none')
                            else:
                                cmap = self.colormap_var.get()
                                self.ax.plot_surface(X, Y, Z, alpha=alpha, label=z_col,
                                                   cmap=cmap, edgecolor='none')
                        else:  # wireframe
                            if color:
                                self.ax.plot_wireframe(X, Y, Z, alpha=alpha, label=z_col,
                                                      color=color, linewidth=linewidth)
                            else:
                                # Wireframe doesn't support cmap well, use default color
                                self.ax.plot_wireframe(X, Y, Z, alpha=alpha, label=z_col,
                                                      linewidth=linewidth)
                    except Exception as e:
                        messagebox.showwarning("Plot Warning", 
                            f"Could not create {plot_type} for {z_col}. Using scatter instead.\n{e}")
                        # Fallback to scatter
                        self.ax.scatter(valid_data['x'], valid_data['y'], valid_data['z'],
                                      c=color, marker=marker, alpha=alpha, label=z_col)
                
                elif plot_type == "scatter3d":
                    self.ax.scatter(valid_data['x'], valid_data['y'], valid_data['z'],
                                  c=color, marker=marker, alpha=alpha, label=z_col, s=50)
                
                elif plot_type == "line3d":
                    if color:
                        self.ax.plot(valid_data['x'], valid_data['y'], valid_data['z'],
                                   color=color, linestyle=linestyle, linewidth=linewidth,
                                   marker=marker if self.marker_var.get() else 'None',
                                   alpha=alpha, label=z_col)
                    else:
                        self.ax.plot(valid_data['x'], valid_data['y'], valid_data['z'],
                                   linestyle=linestyle, linewidth=linewidth,
                                   marker=marker if self.marker_var.get() else 'None',
                                   alpha=alpha, label=z_col)
                
                elif plot_type == "contour3d":
                    try:
                        # Get mesh resolution
                        mesh_res = self.mesh_resolution_var.get()
                        
                        # Get data range
                        x_min, x_max = valid_data['x'].min(), valid_data['x'].max()
                        y_min, y_max = valid_data['y'].min(), valid_data['y'].max()
                        
                        # Create uniform grid
                        x_grid = np.linspace(x_min, x_max, mesh_res)
                        y_grid = np.linspace(y_min, y_max, mesh_res)
                        X, Y = np.meshgrid(x_grid, y_grid)
                        
                        # Interpolate Z values
                        try:
                            from scipy.interpolate import griddata
                            points = valid_data[['x', 'y']].values
                            values = valid_data['z'].values
                            Z = griddata(points, values, (X, Y), method='linear')
                            
                            # Fill NaN with nearest neighbor
                            if np.isnan(Z).any():
                                Z_nearest = griddata(points, values, (X, Y), method='nearest')
                                Z = np.where(np.isnan(Z), Z_nearest, Z)
                        except ImportError:
                            messagebox.showwarning("Missing Library", 
                                "scipy is required for contour3d plots. Install with: pip install scipy")
                            continue
                        
                        # Number of contour levels (proportional to resolution)
                        num_levels = max(10, mesh_res // 5)
                        cmap = self.colormap_var.get()
                        self.ax.contour3D(X, Y, Z, num_levels, cmap=cmap, alpha=alpha)
                    except Exception as e:
                        messagebox.showwarning("Plot Warning",
                            f"Could not create contour3d for {z_col}.\n{e}")
            
            
            # 5. Apply All Settings
            
            # Log Scale Settings (3D axes)
            try:
                self.ax.set_xscale('log' if self.x_log_scale_var.get() else 'linear')
            except (ValueError, NotImplementedError):
                self.x_log_scale_var.set(False)
                self.ax.set_xscale('linear')
            try:
                self.ax.set_yscale('log' if self.y_log_scale_var.get() else 'linear')
            except (ValueError, NotImplementedError):
                self.y_log_scale_var.set(False)
                self.ax.set_yscale('linear')
            try:
                self.ax.set_zscale('log' if self.z_log_scale_var.get() else 'linear')
            except (ValueError, NotImplementedError):
                self.z_log_scale_var.set(False)
                self.ax.set_zscale('linear')

            # Labels and Title
            font_family = self.font_family_var.get()
            self.ax.set_xlabel(self.xlabel_var.get() if self.xlabel_var.get() else x_col, 
                             fontsize=self.xlabel_fontsize_var.get(), fontfamily=font_family)
            self.ax.set_ylabel(self.ylabel_var.get() if self.ylabel_var.get() else y_col, 
                             fontsize=self.ylabel_fontsize_var.get(), fontfamily=font_family)
            
            # Z-axis label: Use custom label if provided, otherwise leave as "Z" or appropriate value axis label
            # Don't use column names as default since Z represents values, not a coordinate
            z_label = self.zlabel_var.get() if self.zlabel_var.get() else "Value"
            self.ax.set_zlabel(z_label, fontsize=self.zlabel_fontsize_var.get(), fontfamily=font_family)
            
            self.ax.set_title(self.title_var.get(), fontsize=self.title_fontsize_var.get(), fontfamily=font_family)
            
            # Grid
            if self.grid_var.get():
                self.ax.grid(True, alpha=self.grid_alpha_var.get())
            else:
                self.ax.grid(False)
            
            # Axis limits
            self.set_axis_limits(self.ax, 'x', self.xlim_min_var.get(), self.xlim_max_var.get())
            self.set_axis_limits(self.ax, 'y', self.ylim_min_var.get(), self.ylim_max_var.get())
            self.set_axis_limits(self.ax, 'z', self.zlim_min_var.get(), self.zlim_max_var.get())

            if self.x_invert_var.get():
                self.ax.invert_xaxis()
            if self.y_invert_var.get():
                self.ax.invert_yaxis()
            if self.z_invert_var.get():
                self.ax.invert_zaxis()

            # Tick settings
            self.ax.tick_params(axis='x', labelsize=self.tick_fontsize_var.get())
            self.ax.tick_params(axis='y', labelsize=self.tick_fontsize_var.get())
            self.ax.tick_params(axis='z', labelsize=self.tick_fontsize_var.get())
            
            # Set font family for tick labels
            for label in self.ax.get_xticklabels() + self.ax.get_yticklabels() + self.ax.get_zticklabels():
                label.set_fontfamily(font_family)
            
            # Background pane colors
            self.ax.xaxis.pane.fill = self.spine_top_var.get()
            self.ax.yaxis.pane.fill = self.spine_left_var.get()
            self.ax.zaxis.pane.fill = self.spine_right_var.get()
            
            if self.spine_top_var.get() or self.spine_left_var.get() or self.spine_right_var.get():
                try:
                    # Parse face color
                    face_rgb = matplotlib.colors.to_rgb(self.face_color_var.get())
                    self.ax.xaxis.pane.set_facecolor(face_rgb + (0.3,))
                    self.ax.yaxis.pane.set_facecolor(face_rgb + (0.3,))
                    self.ax.zaxis.pane.set_facecolor(face_rgb + (0.3,))
                except:
                    pass
            
            # Legend
            if self.legend_show_var.get():
                legend_font_props = {'family': font_family, 'size': self.tick_fontsize_var.get()}
                handles, labels = self.ax.get_legend_handles_labels()
                if handles:
                    self.ax.legend(handles=handles, labels=labels, 
                                 loc=self.legend_loc_var.get(), prop=legend_font_props)
            
            # Set 3D view angle
            self.ax.view_init(elev=self.view_elev_var.get(), azim=self.view_azim_var.get())
            
            self.fig.tight_layout()

            # 6. Update Canvas
            self.canvas.draw()
            
            # Get the figure size in pixels
            fig_width_px = self.fig.get_figwidth() * self.fig.dpi
            fig_height_px = self.fig.get_figheight() * self.fig.dpi
            
            # Update canvas widget size to match figure
            canvas_widget = self.canvas.get_tk_widget()
            canvas_widget.config(width=int(fig_width_px), height=int(fig_height_px))
            
            # Force update to ensure proper sizing
            self.graph_frame.update_idletasks()
            self.on_graph_frame_configure(None)
            
            # Force redraw to clear any artifacts
            self.scrollable_canvas.update_idletasks()


        except Exception as e:
            # 1. UI English: Messagebox
            messagebox.showerror("Plot Error", f"Failed to plot graph:\n{e}\n{type(e)}")
            self.canvas.draw()

    def set_axis_limits(self, ax, axis_name, min_val, max_val):
        """ Set axis limits (handles conversion errors) """
        try:
            min_v = float(min_val) if min_val else None
            max_v = float(max_val) if max_val else None
            
            if axis_name == 'x':
                ax.set_xlim(min_v, max_v)
            elif axis_name == 'y':
                ax.set_ylim(min_v, max_v)
            elif axis_name == 'z':
                ax.set_zlim(min_v, max_v)
        except ValueError:
            pass # Ignore if conversion fails

    # (★ ADDED) Helper method to apply major ticker
    def apply_major_ticker(self, axis, interval_str, is_log_scale):
        """ Apply MultipleLocator if interval_str is valid float and scale is linear """
        if is_log_scale:
            return # Do not apply custom interval on log scale

        try:
            interval = float(interval_str)
            if interval > 0:
                axis.set_major_locator(ticker.MultipleLocator(interval))
        except (ValueError, TypeError):
            pass # Ignore if empty, invalid, or None

    def export_graph(self):
        # 1. UI English: Dialog title
        file_path = filedialog.asksaveasfilename(
            title="Save Graph",
            filetypes=[("PNG files", "*.png"),
                       ("JPEG files", "*.jpg *.jpeg"),
                       ("SVG files", "*.svg"),
                       ("PDF files", "*.pdf"),
                       ("TIFF files", "*.tiff *.tif"),
                       ("EPS files", "*.eps"),
                       ("All files", "*.*")],
            defaultextension=".png"
        )
        if not file_path:
            return

        try:
            matplotlib.rcParams['font.family'] = self.font_family_var.get()
            
            # Detect format from file extension
            file_ext = os.path.splitext(file_path)[1].lower()
            format_map = {
                '.png': 'png',
                '.jpg': 'jpg',
                '.jpeg': 'jpeg',
                '.svg': 'svg',
                '.pdf': 'pdf',
                '.tiff': 'tiff',
                '.tif': 'tiff',
                '.eps': 'eps'
            }
            
            save_format = format_map.get(file_ext, 'png')
            
            # Save with detected format
            self.fig.savefig(file_path, format=save_format, bbox_inches='tight', facecolor=self.fig_color_var.get())
            # 1. UI English: Messagebox
            messagebox.showinfo("Success", f"Graph saved to {file_path}.")
        except Exception as e:
            # 1. UI English: Messagebox
            messagebox.showerror("Save Error", f"Failed to save graph:\n{e}")

    def setup_drag_and_drop(self):
        """Setup drag and drop functionality for file loading"""
        if DND_AVAILABLE:
            try:
                from tkinterdnd2 import DND_FILES
                self.drop_target_register(DND_FILES)
                self.dnd_bind('<<Drop>>', self.on_drop)
            except Exception as e:
                print(f"Drag and drop setup failed: {e}")
        else:
            pass  # Silently fail if tkinterdnd2 not available
    
    def on_drop(self, event):
        """Handle file drop event"""
        try:
            # Get file path from drop event
            files = None
            
            # Try different ways to get the file path
            if hasattr(event, 'data') and event.data:
                files = event.data
                # Try to split if it's a string with multiple files
                try:
                    files = self.tk.splitlist(files)
                except:
                    pass
            
            if not files:
                return
            
            # Handle single file or first file from list
            file_path = files[0] if isinstance(files, (list, tuple)) else files
            
            # Convert to string if needed
            if not isinstance(file_path, str):
                file_path = str(file_path)
            
            # Remove curly braces if present (Windows formatting)
            file_path = file_path.strip('{}').strip()
            
            # Check file extension
            ext = os.path.splitext(file_path)[1].lower()
            
            if ext == '.pmggrp':
                # Load project file
                self.load_project_file(file_path)
            elif ext in ['.csv', '.xlsx', '.xls']:
                # Load data file
                self.load_data(file_path=file_path)
            else:
                messagebox.showwarning("Unsupported File", 
                                     f"File type '{ext}' is not supported.\n"
                                     "Supported: .pmggrp, .csv, .xlsx, .xls")
        except Exception as e:
            messagebox.showerror("Drop Error", f"Failed to load dropped file:\n{e}")
    
    def load_project_file(self, file_path):
        """Load project file (extracted from load_settings for reuse)"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to read project file:\n{e}")
            return

        # Check dimension and relaunch with appropriate app if needed
        file_dimension = settings.get('dimension', '3D')
        if file_dimension == '2D':
            response = messagebox.askyesno("2D Project Detected", 
                "This project was created in 2D mode.\nDo you want to open it in 2D mode?")
            if response:
                import subprocess
                script_dir = os.path.dirname(os.path.abspath(__file__))
                main_2d_path = os.path.join(script_dir, "main.py")
                if os.path.exists(main_2d_path):
                    subprocess.Popen([sys.executable, main_2d_path, file_path])
                    self.quit()
                    return
                else:
                    messagebox.showerror("File Not Found", "2D mode file not found.")
                    return

        # Load Data - always use embedded data if available
        if 'edited_data' not in settings or settings['edited_data'] is None:
            messagebox.showerror("Error", "Project file contains no data.")
            return
            
        # Load edited data from project file
        try:
            data_dict = settings['edited_data']
            self.df = pd.DataFrame(data_dict['data'], columns=data_dict['columns'])
            self.df = self.df.astype(str).fillna("")
            
            # Store original file path if available
            self.data_file_path = settings.get('original_file_path', '')
            
            # Update sheet display
            if self.sheet:
                self.sheet.destroy()

            self.sheet = Sheet(self.sheet_frame,
                               data=self.df.values.tolist(),
                               headers=self.df.columns.tolist(),
                               show_toolbar=True,
                               show_top_left=True)
            self.sheet.enable_bindings()
            self.sheet.pack(fill=tk.BOTH, expand=True)
            
            self.update_plot_options()
            
        except Exception as e:
            messagebox.showerror("Data Load Error", f"Failed to load embedded data:\n{e}")
            return
        
        # Apply all settings
        self.set_variable_from_dict(self.plot_type_var, settings, 'plot_type')
        self.set_variable_from_dict(self.x_axis_var, settings, 'x_axis')
        self.set_variable_from_dict(self.y_axis_var, settings, 'y_axis')
        self.set_variable_from_dict(self.title_var, settings, 'title')
        self.set_variable_from_dict(self.xlabel_var, settings, 'xlabel')
        self.set_variable_from_dict(self.ylabel_var, settings, 'ylabel')
        self.set_variable_from_dict(self.zlabel_var, settings, 'zlabel')
        
        if 'y1_series_styles' in settings:
            self.y1_series_styles = settings['y1_series_styles']
        else:
            self.y1_series_styles = {}

        self.set_variable_from_dict(self.grid_var, settings, 'grid')
        self.set_variable_from_dict(self.marker_var, settings, 'marker')
        
        self.set_variable_from_dict(self.font_family_var, settings, 'font_family')
        self.set_variable_from_dict(self.title_fontsize_var, settings, 'title_fontsize')
        self.set_variable_from_dict(self.xlabel_fontsize_var, settings, 'xlabel_fontsize')
        self.set_variable_from_dict(self.ylabel_fontsize_var, settings, 'ylabel_fontsize')
        self.set_variable_from_dict(self.zlabel_fontsize_var, settings, 'zlabel_fontsize', fallback_key='ylabel2_fontsize')
        self.set_variable_from_dict(self.tick_fontsize_var, settings, 'tick_fontsize')
        self.set_variable_from_dict(self.fig_width_var, settings, 'fig_width')
        self.set_variable_from_dict(self.fig_height_var, settings, 'fig_height')
        
        self.set_variable_from_dict(self.xlim_min_var, settings, 'xlim_min')
        self.set_variable_from_dict(self.xlim_max_var, settings, 'xlim_max')
        self.set_variable_from_dict(self.ylim_min_var, settings, 'ylim_min')
        self.set_variable_from_dict(self.ylim_max_var, settings, 'ylim_max')
        self.set_variable_from_dict(self.zlim_min_var, settings, 'zlim_min', fallback_key='ylim2_min')
        self.set_variable_from_dict(self.zlim_max_var, settings, 'zlim_max', fallback_key='ylim2_max')
        self.set_variable_from_dict(self.xtick_show_var, settings, 'xtick_show')
        self.set_variable_from_dict(self.xtick_label_show_var, settings, 'xtick_label_show')
        self.set_variable_from_dict(self.xtick_direction_var, settings, 'xtick_direction')
        self.set_variable_from_dict(self.ytick_show_var, settings, 'ytick_show')
        self.set_variable_from_dict(self.ytick_label_show_var, settings, 'ytick_label_show')
        self.set_variable_from_dict(self.ytick_direction_var, settings, 'ytick_direction')
        self.set_variable_from_dict(self.ztick_show_var, settings, 'ztick_show', fallback_key='ytick2_show')
        self.set_variable_from_dict(self.ztick_label_show_var, settings, 'ztick_label_show', fallback_key='ytick2_label_show')
        self.set_variable_from_dict(self.ztick_direction_var, settings, 'ztick_direction', fallback_key='ytick2_direction')
        
        self.set_variable_from_dict(self.xaxis_plain_format_var, settings, 'xaxis_plain_format')
        self.set_variable_from_dict(self.yaxis_plain_format_var, settings, 'yaxis_plain_format', fallback_key='yaxis1_plain_format')
        self.set_variable_from_dict(self.zaxis_plain_format_var, settings, 'zaxis_plain_format', fallback_key='yaxis2_plain_format')

        self.set_variable_from_dict(self.xtick_major_interval_var, settings, 'xtick_major_interval')
        self.set_variable_from_dict(self.ytick_major_interval_var, settings, 'ytick_major_interval')
        self.set_variable_from_dict(self.ztick_major_interval_var, settings, 'ztick_major_interval', fallback_key='ytick2_major_interval')

        self.set_variable_from_dict(self.spine_top_var, settings, 'spine_top')
        self.set_variable_from_dict(self.spine_bottom_var, settings, 'spine_bottom')
        self.set_variable_from_dict(self.spine_left_var, settings, 'spine_left')
        self.set_variable_from_dict(self.spine_right_var, settings, 'spine_right')
        self.set_variable_from_dict(self.face_color_var, settings, 'face_color')
        self.set_variable_from_dict(self.fig_color_var, settings, 'fig_color')
        
        self.set_variable_from_dict(self.legend_show_var, settings, 'legend_show')
        self.set_variable_from_dict(self.legend_loc_var, settings, 'legend_loc')
        
        self.set_variable_from_dict(self.x_log_scale_var, settings, 'x_log_scale')
        self.set_variable_from_dict(self.y_log_scale_var, settings, 'y_log_scale', fallback_key='y1_log_scale')
        self.set_variable_from_dict(self.z_log_scale_var, settings, 'z_log_scale', fallback_key='y2_log_scale')

        self.set_variable_from_dict(self.x_invert_var, settings, 'x_invert')
        self.set_variable_from_dict(self.y_invert_var, settings, 'y_invert', fallback_key='y1_invert')
        self.set_variable_from_dict(self.z_invert_var, settings, 'z_invert', fallback_key='y2_invert')
        
        self.set_variable_from_dict(self.grid_alpha_var, settings, 'grid_alpha')
        self.set_variable_from_dict(self.view_elev_var, settings, 'view_elev')
        self.set_variable_from_dict(self.view_azim_var, settings, 'view_azim')
        self.set_variable_from_dict(self.mesh_resolution_var, settings, 'mesh_resolution')
        self.set_variable_from_dict(self.colormap_var, settings, 'colormap')

        # Restore Listbox selections
        if self.df is not None:
            self.z_listbox.select_clear(0, tk.END)
            if 'z_axis_indices' in settings:
                for i in settings['z_axis_indices']:
                    if i < self.z_listbox.size():
                        self.z_listbox.select_set(i)

        # Update color labels
        self.update_color_label(self.face_color_label, self.face_color_var.get())
        self.update_color_label(self.fig_color_label, self.fig_color_var.get())

        # Store current project path and enable overwrite save button
        self.current_project_path = file_path
        self.overwrite_save_button['state'] = 'normal'
        
        # Redraw graph
        if self.df is not None:
            self.plot_graph()
    
    def open_in_2d_mode(self):
        """Open the 2D version of the application"""
        import subprocess
        script_dir = os.path.dirname(os.path.abspath(__file__))
        main_2d_path = os.path.join(script_dir, "main.py")
        
        if os.path.exists(main_2d_path):
            try:
                # Launch 2D version with current project file if available
                if self.current_project_path and os.path.exists(self.current_project_path):
                    subprocess.Popen([sys.executable, main_2d_path, self.current_project_path])
                else:
                    subprocess.Popen([sys.executable, main_2d_path])
            except Exception as e:
                messagebox.showerror("Launch Error", f"Failed to launch 2D mode:\n{e}")
        else:
            messagebox.showerror("File Not Found", f"2D mode file not found:\n{main_2d_path}")
    
    def show_about(self):
        """Show about dialog"""
        about_text = f"""HYGrapher 3D ver. {VERSION}

A flexible 3D graphing application for CSV and Excel data.

Author: Hiromichi Yokoyama
License: Apache-2.0 license
Repository: https://github.com/HiroYokoyama/matplotlib_graph_app

Keyboard Shortcuts:
  Ctrl+O - Open Project
  Ctrl+S - Save Project
"""
        messagebox.showinfo("About HYGrapher 3D", about_text)

def main():
    """アプリケーションを起動するメイン関数"""
    # Create the application instance
    app = GraphApp()
    
    # Handle command line arguments
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        
        if os.path.exists(file_path):
            ext = os.path.splitext(file_path)[1].lower()
            
            # Load file after main window is ready
            def load_startup_file():
                try:
                    if ext == '.pmggrp':
                        app.load_project_file(file_path)
                    elif ext in ['.csv', '.xlsx', '.xls']:
                        app.load_data(file_path=file_path)
                    else:
                        messagebox.showwarning("Unsupported File", 
                                             f"File type '{ext}' is not supported.\n"
                                             "Supported: .pmggrp, .csv, .xlsx, .xls")
                except Exception as e:
                    messagebox.showerror("Startup Error", f"Failed to load file:\n{e}")
            
            # Schedule file loading after GUI is initialized
            app.after(100, load_startup_file)
        else:
            messagebox.showerror("File Not Found", f"File does not exist:\n{file_path}")
    
    app.mainloop()

if __name__ == "__main__":
    main()

