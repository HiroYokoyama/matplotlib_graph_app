import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.font_manager as fm
import matplotlib
import os
import json
from tksheet import Sheet # Import tksheet

class GraphApp(tk.Tk):
    def __init__(self):
        super().__init__()
        # 1. UI English: Window Title
        self.title("Matplotlib Graphing App ver. 0.1.1")
        self.geometry("1600x900") # Keep window size

        self.df = None
        self.sheet = None
        self.data_file_path = "" # Store the path of the loaded file

        # --- Get Font List ---
        self.font_list = self.get_font_list()

        # --- Create all tk.Variables ---
        self.create_all_tk_variables()

        # --- Figure ---
        self.fig = Figure(figsize=(self.fig_width_var.get(), self.fig_height_var.get()), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax2 = None # For 2nd Y-axis

        # === Main Layout ===
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Top Frame (File Operations) ---
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X)

        # 1. UI English: Button text
        self.load_button = ttk.Button(top_frame, text="Load Data (CSV/Excel)", command=self.load_data)
        self.load_button.pack(side=tk.LEFT, padx=5)

        self.save_settings_button = ttk.Button(top_frame, text="Save Settings (.json)", command=self.save_settings)
        self.save_settings_button.pack(side=tk.LEFT, padx=5)
        
        self.load_settings_button = ttk.Button(top_frame, text="Load Settings (.json)", command=self.load_settings)
        self.load_settings_button.pack(side=tk.LEFT, padx=5)

        self.export_button = ttk.Button(top_frame, text="Export Graph (Image)", command=self.export_graph)
        self.export_button.pack(side=tk.LEFT, padx=5)
        self.export_button['state'] = 'disabled'

        # --- Content Frame (Split) ---
        # 2. Layout: Use PanedWindow for resizable 1:1 split
        content_frame = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # --- Left Panel (Data and Settings) ---
        # 2. Layout: Remove fixed width and pack_propagate
        left_panel = ttk.Frame(content_frame)
        # 2. Layout: Add to PanedWindow with weight 1
        content_frame.add(left_panel, weight=1)

        # --- Graph Settings (Notebook) ---
        self.settings_notebook = ttk.Notebook(left_panel)
        self.settings_notebook.pack(fill=tk.X, pady=5)
        
        # === Tab 1: Basic Settings (Y-axis as Listbox) ===
        basic_settings_frame = ttk.Frame(self.settings_notebook, padding=10)
        # 1. UI English: Tab text
        self.settings_notebook.add(basic_settings_frame, text="Basic Settings")
        self.create_basic_settings_tab(basic_settings_frame) # (★ MODIFIED)

        # === Tab 2: Style Settings (★ Major Change) ===
        style_settings_frame = ttk.Frame(self.settings_notebook, padding=10)
        # 1. UI English: Tab text
        # (★ MODIFIED) 4. Remove (Series) from tab name
        self.settings_notebook.add(style_settings_frame, text="Style")
        self.create_style_settings_tab(style_settings_frame)
        
        # (★ MODIFIED) 2. Re-add Font tab
        font_size_frame = ttk.Frame(self.settings_notebook, padding=10)
        self.settings_notebook.add(font_size_frame, text="Font")
        self.create_font_size_tab(font_size_frame) # (★ MODIFIED)
        
        # === Tab 4: Axis & Ticks Settings === (Now Tab 4, was 3)
        axis_ticks_frame = ttk.Frame(self.settings_notebook, padding=10)
        # 1. UI English: Tab text
        self.settings_notebook.add(axis_ticks_frame, text="Axis/Ticks")
        self.create_axis_ticks_tab(axis_ticks_frame) # (★ MODIFIED)

        # === Tab 5: Spines & Background === (Now Tab 4)
        spines_frame = ttk.Frame(self.settings_notebook, padding=10)
        # 1. UI English: Tab text
        self.settings_notebook.add(spines_frame, text="Spines/BG")
        self.create_spines_tab(spines_frame)

        # === Tab 6: Legend Settings === (Now Tab 5)
        legend_frame = ttk.Frame(self.settings_notebook, padding=10)
        # 1. UI English: Tab text
        self.settings_notebook.add(legend_frame, text="Legend")
        self.create_legend_tab(legend_frame)

        # --- Plot Button ---
        # 1. UI English: Button text
        self.plot_button = ttk.Button(left_panel, text="Plot/Update Graph", command=self.plot_graph)
        self.plot_button.pack(fill=tk.X, pady=10)
        self.plot_button['state'] = 'disabled'

        # --- Data Edit Area ---
        # 1. UI English: LabelFrame text
        data_frame = ttk.LabelFrame(left_panel, text="Data Editor")
        data_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.sheet_frame = ttk.Frame(data_frame)
        self.sheet_frame.pack(fill=tk.BOTH, expand=True)

        # === Right Panel (Graph) ===
        # 1. UI English: LabelFrame text
        right_panel = ttk.LabelFrame(content_frame, text="Graph Preview")
        # 2. Layout: Add to PanedWindow with weight 1
        content_frame.add(right_panel, weight=1)

        # (★ MODIFIED) 2. Layout: Try to force 1:1 split on start
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
        # (★ MODIFIED) 2. Scrollbar Fix: Make scrollable_canvas a self variable
        self.scrollable_canvas = tk.Canvas(right_panel, borderwidth=0, highlightthickness=0)
        
        v_scroll = ttk.Scrollbar(right_panel, orient=tk.VERTICAL, command=self.scrollable_canvas.yview)
        h_scroll = ttk.Scrollbar(right_panel, orient=tk.HORIZONTAL, command=self.scrollable_canvas.xview)
        self.scrollable_canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.scrollable_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Internal frame for Matplotlib graph and toolbar
        # (★ MODIFIED) 2. Scrollbar Fix: Use self.scrollable_canvas
        self.graph_frame = ttk.Frame(self.scrollable_canvas)
        self.scrollable_canvas.create_window((0, 0), window=self.graph_frame, anchor="nw")

        # Graph and Toolbar master changed to self.graph_frame
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_frame)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.graph_frame)
        self.toolbar.update()
        self.toolbar.pack(side=tk.TOP, fill=tk.X) 
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.NONE, expand=False, padx=5, pady=5)

        # (★ MODIFIED) 2. Scrollbar Fix: Bind frame configure event to a method
        self.graph_frame.bind("<Configure>", self.on_graph_frame_configure)

    # (★ MODIFIED) 2. Scrollbar Fix: Add method for frame configure
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
        self.plot_type_var = tk.StringVar(value="line")
        self.x_axis_var = tk.StringVar()
        self.title_var = tk.StringVar()
        self.xlabel_var = tk.StringVar()
        self.ylabel_var = tk.StringVar()
        # 1. UI English: Default variable text
        self.ylabel2_var = tk.StringVar(value="2nd Y-Axis Label")
        
        # --- (★ Style Refactor) ---
        # Remove axis-level style variables (e.g., linestyle_y1_var)
        # Add dictionaries to hold styles per series (by column name)
        self.y1_series_styles = {}
        self.y2_series_styles = {}

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
        # (★ MODIFIED) 5. Default font to Arial
        # (★ MODIFIED) 6. Default font to the *first* available font (which is likely Arial)
        self.font_family_var = tk.StringVar(value=self.font_list[0] if self.font_list else 'sans-serif')
        self.title_fontsize_var = tk.DoubleVar(value=16.0)
        self.xlabel_fontsize_var = tk.DoubleVar(value=14.0)
        self.ylabel_fontsize_var = tk.DoubleVar(value=14.0)
        self.ylabel2_fontsize_var = tk.DoubleVar(value=14.0)
        self.tick_fontsize_var = tk.DoubleVar(value=14.0)
        self.tick2_fontsize_var = tk.DoubleVar(value=14.0)
        self.fig_width_var = tk.DoubleVar(value=7.0)
        self.fig_height_var = tk.DoubleVar(value=6.0)
        
        # Axis/Ticks
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
        
        # (★ MODIFIED) 1. Scientific Notation Fix: Add variables
        self.xaxis_plain_format_var = tk.BooleanVar(value=False)
        self.yaxis1_plain_format_var = tk.BooleanVar(value=False) # Renamed
        self.yaxis2_plain_format_var = tk.BooleanVar(value=False)
        
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

        # (★ ADDED) Log Scale Vars
        self.x_log_scale_var = tk.BooleanVar(value=False)
        self.y1_log_scale_var = tk.BooleanVar(value=False)
        self.y2_log_scale_var = tk.BooleanVar(value=False)

        # (★ MODIFIED) Add Invert Axis Vars
        self.x_invert_var = tk.BooleanVar(value=False)
        self.y1_invert_var = tk.BooleanVar(value=False)
        self.y2_invert_var = tk.BooleanVar(value=False)


    # --- Tab Creation Methods ---
    def create_basic_settings_tab(self, frame):
        # (★ MODIFIED) 6. Complete rewrite for layout change
        # 1. UI English: Labels
        
        # --- Top-level settings ---
        top_settings_frame = ttk.Frame(frame)
        top_settings_frame.pack(fill=tk.X, pady=2)

        ttk.Label(top_settings_frame, text="Graph Title:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.title_entry = ttk.Entry(top_settings_frame, textvariable=self.title_var, width=50)
        self.title_entry.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky=tk.EW)

        ttk.Label(top_settings_frame, text="Plot Type:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.plot_type_combo = ttk.Combobox(top_settings_frame, textvariable=self.plot_type_var, 
                                            values=["line", "scatter", "bar"], state='readonly', width=48)
        self.plot_type_combo.grid(row=1, column=1, columnspan=3, padx=5, pady=5, sticky=tk.EW)
        
        top_settings_frame.columnconfigure(1, weight=1)

        # --- Axis Settings Frames ---
        axis_frames_container = ttk.Frame(frame)
        axis_frames_container.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # --- X-Axis Frame ---
        x_axis_frame = ttk.LabelFrame(axis_frames_container, text="X-Axis")
        x_axis_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(x_axis_frame, text="Label:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.xlabel_entry = ttk.Entry(x_axis_frame, textvariable=self.xlabel_var, width=50)
        self.xlabel_entry.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky=tk.EW) # (★ MODIFIED) Change columnspan
        
        ttk.Label(x_axis_frame, text="Data:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.x_axis_combo = ttk.Combobox(x_axis_frame, textvariable=self.x_axis_var, state='disabled', width=48)
        self.x_axis_combo.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky=tk.EW) # (★ MODIFIED) Change columnspan
        
        # (★ ADDED) Log scale checkbox
        self.x_log_scale_check = ttk.Checkbutton(x_axis_frame, text="Log Scale", variable=self.x_log_scale_var)
        self.x_log_scale_check.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W) # (★ MODIFIED) Change grid
        
        # (★ MODIFIED) Add Invert X-Axis Checkbox here
        self.x_invert_check = ttk.Checkbutton(x_axis_frame, text="Invert Axis", variable=self.x_invert_var)
        self.x_invert_check.grid(row=2, column=1, padx=15, pady=5, sticky=tk.W) # Place next to log scale
        
        x_axis_frame.columnconfigure(1, weight=1)

        # --- Y-Axis Frames (in a PanedWindow) ---
        y_axis_paned_window = ttk.PanedWindow(axis_frames_container, orient=tk.HORIZONTAL)
        y_axis_paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- Y-Axis (Left) Frame ---
        y1_axis_frame = ttk.LabelFrame(y_axis_paned_window, text="Y-Axis (Left)")
        y_axis_paned_window.add(y1_axis_frame, weight=1)
        
        ttk.Label(y1_axis_frame, text="Label:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.ylabel_entry = ttk.Entry(y1_axis_frame, textvariable=self.ylabel_var, width=24)
        self.ylabel_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        
        # (★ ADDED) Log scale checkbox
        self.y1_log_scale_check = ttk.Checkbutton(y1_axis_frame, text="Log Scale", variable=self.y1_log_scale_var)
        self.y1_log_scale_check.grid(row=0, column=2, padx=15, pady=5, sticky=tk.W) # (★ MODIFIED) Change grid
        
        # (★ MODIFIED) Add Invert Y1-Axis Checkbox here
        self.y1_invert_check = ttk.Checkbutton(y1_axis_frame, text="Invert Axis", variable=self.y1_invert_var)
        self.y1_invert_check.grid(row=0, column=3, padx=15, pady=5, sticky=tk.W) # Place next to log scale
        
        ttk.Label(y1_axis_frame, text="Data (Multi-select):").grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W) # (★ MODIFIED) Change columnspan
        
        self.y_listbox_frame = ttk.Frame(y1_axis_frame)
        self.y_listbox_scroll = ttk.Scrollbar(self.y_listbox_frame, orient=tk.VERTICAL)
        self.y_listbox = tk.Listbox(self.y_listbox_frame, selectmode=tk.MULTIPLE, yscrollcommand=self.y_listbox_scroll.set, exportselection=False, height=6)
        self.y_listbox_scroll.config(command=self.y_listbox.yview)
        self.y_listbox_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.y_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.y_listbox_frame.grid(row=2, column=0, columnspan=4, padx=5, pady=5, sticky=tk.NSEW) # (★ MODIFIED) Change columnspan
        
        y1_axis_frame.columnconfigure(1, weight=1)
        y1_axis_frame.rowconfigure(2, weight=1)

        # --- Y-Axis (Right) Frame ---
        y2_axis_frame = ttk.LabelFrame(y_axis_paned_window, text="Y-Axis (Right)")
        y_axis_paned_window.add(y2_axis_frame, weight=1)

        ttk.Label(y2_axis_frame, text="Label:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.ylabel2_entry = ttk.Entry(y2_axis_frame, textvariable=self.ylabel2_var, width=24)
        self.ylabel2_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)

        # (★ ADDED) Log scale checkbox
        self.y2_log_scale_check = ttk.Checkbutton(y2_axis_frame, text="Log Scale", variable=self.y2_log_scale_var)
        self.y2_log_scale_check.grid(row=0, column=2, padx=15, pady=5, sticky=tk.W) # (★ MODIFIED) Change grid

        # (★ MODIFIED) Add Invert Y2-Axis Checkbox here
        self.y2_invert_check = ttk.Checkbutton(y2_axis_frame, text="Invert Axis", variable=self.y2_invert_var)
        self.y2_invert_check.grid(row=0, column=3, padx=15, pady=5, sticky=tk.W) # Place next to log scale

        ttk.Label(y2_axis_frame, text="Data (Multi-select):").grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W) # (★ MODIFIED) Change columnspan

        self.y2_listbox_frame = ttk.Frame(y2_axis_frame)
        self.y2_listbox_scroll = ttk.Scrollbar(self.y2_listbox_frame, orient=tk.VERTICAL)
        self.y2_listbox = tk.Listbox(self.y2_listbox_frame, selectmode=tk.MULTIPLE, yscrollcommand=self.y2_listbox_scroll.set, exportselection=False, height=6)
        self.y2_listbox_scroll.config(command=self.y2_listbox.yview)
        self.y2_listbox_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.y2_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.y2_listbox_frame.grid(row=2, column=0, columnspan=4, padx=5, pady=5, sticky=tk.NSEW) # (★ MODIFIED) Change columnspan

        y2_axis_frame.columnconfigure(1, weight=1)
        y2_axis_frame.rowconfigure(2, weight=1)

        # --- Figure Size ---
        fig_size_frame = ttk.Frame(frame)
        fig_size_frame.pack(fill=tk.X, pady=(10,0))
        
        ttk.Label(fig_size_frame, text="Figure Width (inch):").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.fig_width_spin = ttk.Spinbox(fig_size_frame, from_=3, to=20, increment=0.5, textvariable=self.fig_width_var, width=10)
        self.fig_width_spin.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

        ttk.Label(fig_size_frame, text="Figure Height (inch):").grid(row=0, column=2, padx=15, pady=5, sticky=tk.W)
        self.fig_height_spin = ttk.Spinbox(fig_size_frame, from_=3, to=20, increment=0.5, textvariable=self.fig_height_var, width=10)
        self.fig_height_spin.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)

    def create_style_settings_tab(self, frame):
        # --- (★ Style Refactor) ---
        
        # (★ MODIFIED) 4. Move Common Settings to top
        common_frame = ttk.LabelFrame(frame, text="Common Style Settings")
        common_frame.pack(fill=tk.X, padx=5, pady=5)

        self.grid_check = ttk.Checkbutton(common_frame, text="Show Grid", variable=self.grid_var)
        self.grid_check.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        self.marker_check = ttk.Checkbutton(common_frame, text="Show Markers (Line/Scatter)", variable=self.marker_var)
        self.marker_check.grid(row=0, column=1, padx=15, pady=5, sticky=tk.W)
        
        ttk.Label(common_frame, text="(Note: Series-specific styles are always used)").grid(row=0, column=2, padx=15, pady=5, sticky=tk.W)

        # (★ MODIFIED) 2. Remove Font Settings from Style tab
        # font_settings_frame = ttk.LabelFrame(frame, text="Font Settings")
        # font_settings_frame.pack(fill=tk.X, padx=5, pady=5)
        # self.create_font_size_widgets(font_settings_frame) # Call helper
        
        # --- Separator ---
        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(10, 5))


        # --- Top frame for Y1/Y2 selectors ---
        selector_frame = ttk.Frame(frame)
        selector_frame.pack(fill=tk.X, pady=5)


        # ★ 1. Consolidate: Add a single, combined selector
        ttk.Label(selector_frame, text="Select Series (Y1/Y2):").pack(side=tk.LEFT, padx=5, pady=5)
        self.style_combo = ttk.Combobox(selector_frame, textvariable=self.combined_style_target_var, state='readonly', width=40)
        self.style_combo.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        # Bind selection event to the new consolidated callback
        self.style_combo.bind("<<ComboboxSelected>>", self.on_combined_series_select)

        # --- Common "Style Editor" Frame ---
        editor_frame = ttk.LabelFrame(frame, text="Style Editor (for selected series)")
        editor_frame.pack(fill=tk.X, padx=5, pady=5)

        # Editor Row 1: Color
        ttk.Label(editor_frame, text="Line/Bar Color:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.style_editor_color_btn = ttk.Button(editor_frame, text="Select", command=self.on_style_editor_color_pick, width=8)
        self.style_editor_color_btn.grid(row=0, column=1, padx=2, pady=5) # Reduced padx
        
        # (★ ADDED) Auto button
        self.style_editor_color_auto_btn = ttk.Button(editor_frame, text="Auto", command=self.on_style_editor_color_auto, width=8)
        self.style_editor_color_auto_btn.grid(row=0, column=2, padx=2, pady=5) # Add Auto button
        
        self.style_editor_color_label = ttk.Label(editor_frame, text="#000000", background="#000000", width=10, anchor=tk.CENTER)
        self.style_editor_color_label.grid(row=0, column=3, padx=5, pady=5) # Shifted column

        # Editor Row 1: Line Style (Right side)
        ttk.Label(editor_frame, text="Line Style:").grid(row=0, column=4, padx=15, pady=5, sticky=tk.W) # Shifted column
        self.style_editor_linestyle_combo = ttk.Combobox(editor_frame, textvariable=self.current_style_linestyle_var, 
                                            values=['-', '--', ':', '-.', 'None'], state='readonly', width=10)
        self.style_editor_linestyle_combo.grid(row=0, column=5, padx=5, pady=5, sticky=tk.W) # Shifted column
        self.style_editor_linestyle_combo.bind("<<ComboboxSelected>>", self.on_style_editor_change)

        # Editor Row 2: Marker Style
        ttk.Label(editor_frame, text="Marker Style:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.style_editor_marker_combo = ttk.Combobox(editor_frame, textvariable=self.current_style_marker_var, 
                                              values=['o', '.', ',', 's', 'p', '*', '^', '<', '>', 'D', 'H', 'None'], state='readonly', width=10)
        self.style_editor_marker_combo.grid(row=1, column=1, columnspan=3, padx=5, pady=5, sticky=tk.W) # Use columnspan 3
        self.style_editor_marker_combo.bind("<<ComboboxSelected>>", self.on_style_editor_change)

        # Editor Row 2: Line Width (Right side)
        ttk.Label(editor_frame, text="Line Width:").grid(row=1, column=4, padx=15, pady=5, sticky=tk.W) # Shifted column
        self.style_editor_linewidth_spin = ttk.Spinbox(editor_frame, from_=0.5, to=10.0, increment=0.5, textvariable=self.current_style_linewidth_var, width=10,
                                                       command=self.on_style_editor_change) # command handles spinbox change
        self.style_editor_linewidth_spin.grid(row=1, column=5, padx=5, pady=5, sticky=tk.W) # Shifted column
        # Also bind Return key for manual entry
        self.style_editor_linewidth_spin.bind("<Return>", self.on_style_editor_change)


        # Editor Row 3: Alpha
        ttk.Label(editor_frame, text="Alpha (Opacity):").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.style_editor_alpha_spin = ttk.Spinbox(editor_frame, from_=0.0, to=1.0, increment=0.1, textvariable=self.current_style_alpha_var, width=10,
                                                   command=self.on_style_editor_change)
        self.style_editor_alpha_spin.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky=tk.W) # Use columnspan 3
        self.style_editor_alpha_spin.bind("<Return>", self.on_style_editor_change)


    # (★ MODIFIED) 2. Re-create create_font_size_tab (was create_font_size_widgets)
    def create_font_size_tab(self, frame):
        # (★ MODIFIED) 3. This is the content from the old create_font_size_tab
        # 1. UI English: Labels
        ttk.Label(frame, text="Font Family:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.font_family_combo = ttk.Combobox(frame, textvariable=self.font_family_var, 
                                              values=self.font_list, state='readonly', width=30)
        self.font_family_combo.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky=tk.W)

        # Font Size (Left Column)
        ttk.Label(frame, text="Title Size:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.title_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.title_fontsize_var, width=10)
        self.title_fontsize_spin.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

        ttk.Label(frame, text="X-Label Size:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.xlabel_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.xlabel_fontsize_var, width=10)
        self.xlabel_fontsize_spin.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)

        ttk.Label(frame, text="Y-Left Label Size:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.ylabel_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.ylabel_fontsize_var, width=10)
        self.ylabel_fontsize_spin.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)

        ttk.Label(frame, text="Y-Right Label Size:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.ylabel2_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.ylabel2_fontsize_var, width=10)
        self.ylabel2_fontsize_spin.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)

        # Font Size (Right Column)
        ttk.Label(frame, text="Ticks (Left/X) Size:").grid(row=1, column=2, padx=15, pady=5, sticky=tk.W)
        self.tick_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.tick_fontsize_var, width=10)
        self.tick_fontsize_spin.grid(row=1, column=3, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(frame, text="Ticks (Right) Size:").grid(row=2, column=2, padx=15, pady=5, sticky=tk.W)
        self.tick2_fontsize_spin = ttk.Spinbox(frame, from_=6, to=48, increment=1, textvariable=self.tick2_fontsize_var, width=10)
        self.tick2_fontsize_spin.grid(row=2, column=3, padx=5, pady=5, sticky=tk.W)

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
        
        ttk.Label(frame, text="Tick Direction:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.xtick_direction_combo = ttk.Combobox(frame, textvariable=self.xtick_direction_var, 
                                                  values=['out', 'in', 'inout'], state='readonly', width=8)
        self.xtick_direction_combo.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        self.xtick_show_check = ttk.Checkbutton(frame, text="Show Ticks", variable=self.xtick_show_var)
        self.xtick_show_check.grid(row=2, column=2, padx=5, pady=5, sticky=tk.W)
        self.xtick_label_show_check = ttk.Checkbutton(frame, text="Show Labels", variable=self.xtick_label_show_var)
        self.xtick_label_show_check.grid(row=2, column=3, padx=5, pady=5, sticky=tk.W)
        
        # (★ MODIFIED) 1. Add X-Axis scientific notation checkbox
        # (★ MODIFIED) 3. Update text, remove 'e.g.'
        # (★ MODIFIED) 4. Remove example text
        self.xaxis_plain_format_check = ttk.Checkbutton(frame, 
            text="Disable Scientific Notation", 
            variable=self.xaxis_plain_format_var)
        # (★ MODIFIED) 5. Move to new row
        self.xaxis_plain_format_check.grid(row=3, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W) # (★ MODIFIED) Revert columnspan

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=4, column=0, columnspan=5, sticky="ew", pady=10) # Shifted from 3 to 4

        # --- Y-Axis (Left) ---
        ttk.Label(frame, text="Y-Axis (Left)", font=("-weight bold")).grid(row=5, column=0, sticky=tk.W, pady=5) # Shifted from 4 to 5
        
        ttk.Label(frame, text="Range (Min):").grid(row=6, column=0, padx=5, pady=5, sticky=tk.W) # Shifted from 5 to 6
        self.ylim_min_entry = ttk.Entry(frame, textvariable=self.ylim_min_var, width=10)
        self.ylim_min_entry.grid(row=6, column=1, padx=5, pady=5) # Shifted from 5 to 6
        ttk.Label(frame, text="Range (Max):").grid(row=6, column=2, padx=5, pady=5, sticky=tk.W) # Shifted from 5 to 6
        self.ylim_max_entry = ttk.Entry(frame, textvariable=self.ylim_max_var, width=10)
        self.ylim_max_entry.grid(row=6, column=3, padx=5, pady=5) # Shifted from 5 to 6
        
        ttk.Label(frame, text="Tick Direction:").grid(row=7, column=0, padx=5, pady=5, sticky=tk.W) # Shifted from 6 to 7
        self.ytick_direction_combo = ttk.Combobox(frame, textvariable=self.ytick_direction_var, 
                                                  values=['out', 'in', 'inout'], state='readonly', width=8)
        self.ytick_direction_combo.grid(row=7, column=1, padx=5, pady=5, sticky=tk.W) # Shifted from 6 to 7
        self.ytick_show_check = ttk.Checkbutton(frame, text="Show Ticks", variable=self.ytick_show_var)
        self.ytick_show_check.grid(row=7, column=2, padx=5, pady=5, sticky=tk.W) # Shifted from 6 to 7
        self.ytick_label_show_check = ttk.Checkbutton(frame, text="Show Labels", variable=self.ytick_label_show_var)
        self.ytick_label_show_check.grid(row=7, column=3, padx=5, pady=5, sticky=tk.W) # Shifted from 6 to 7

        # (★ MODIFIED) 1. Scientific Notation Fix: Add Checkbox, change text, change variable
        # (★ MODIFIED) 3. Update text, remove 'e.g.'
        # (★ MODIFIED) 4. Remove example text
        self.yaxis_plain_format_check = ttk.Checkbutton(frame, 
            text="Disable Scientific Notation", 
            variable=self.yaxis1_plain_format_var)
        # (★ MODIFIED) 5. Move to new row
        self.yaxis_plain_format_check.grid(row=8, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W) # Shifted from 6 to 8 (★ MODIFIED) Revert columnspan
        
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=9, column=0, columnspan=5, sticky="ew", pady=10) # Shifted from 7 to 9
        
        # --- Y-Axis (Right) ---
        ttk.Label(frame, text="Y-Axis (Right)", font=("-weight bold")).grid(row=10, column=0, sticky=tk.W, pady=5) # Shifted from 8 to 10
        
        ttk.Label(frame, text="Range (Min):").grid(row=11, column=0, padx=5, pady=5, sticky=tk.W) # Shifted from 9 to 11
        self.ylim2_min_entry = ttk.Entry(frame, textvariable=self.ylim2_min_var, width=10)
        self.ylim2_min_entry.grid(row=11, column=1, padx=5, pady=5) # Shifted from 9 to 11
        ttk.Label(frame, text="Range (Max):").grid(row=11, column=2, padx=5, pady=5, sticky=tk.W) # Shifted from 9 to 11
        self.ylim2_max_entry = ttk.Entry(frame, textvariable=self.ylim2_max_var, width=10)
        self.ylim2_max_entry.grid(row=11, column=3, padx=5, pady=5) # Shifted from 9 to 11
        
        ttk.Label(frame, text="Tick Direction:").grid(row=12, column=0, padx=5, pady=5, sticky=tk.W) # Shifted from 10 to 12
        self.ytick2_direction_combo = ttk.Combobox(frame, textvariable=self.ytick2_direction_var, 
                                                  values=['out', 'in', 'inout'], state='readonly', width=8)
        self.ytick2_direction_combo.grid(row=12, column=1, padx=5, pady=5, sticky=tk.W) # Shifted from 10 to 12
        self.ytick2_show_check = ttk.Checkbutton(frame, text="Show Ticks", variable=self.ytick2_show_var)
        self.ytick2_show_check.grid(row=12, column=2, padx=5, pady=5, sticky=tk.W) # Shifted from 10 to 12
        self.ytick2_label_show_check = ttk.Checkbutton(frame, text="Show Labels", variable=self.ytick2_label_show_var)
        self.ytick2_label_show_check.grid(row=12, column=3, padx=5, pady=5, sticky=tk.W) # Shifted from 10 to 12
        
        # (★ MODIFIED) 1. Add Y2-Axis scientific notation checkbox
        # (★ MODIFIED) 3. Update text, remove 'e.g.'
        # (★ MODIFIED) 4. Remove example text
        self.yaxis2_plain_format_check = ttk.Checkbutton(frame, 
            text="Disable Scientific Notation", 
            variable=self.yaxis2_plain_format_var)
        # (★ MODIFIED) 5. Move to new row
        self.yaxis2_plain_format_check.grid(row=13, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W) # Shifted from 10 to 13 (★ MODIFIED) Revert columnspan

        # (★ MODIFIED) 1. Make column 4 expandable for the checkbox text
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
        
        ttk.Label(frame, text="(Legend labels use Y-axis column names automatically)").grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)


    # --- (★ Style Refactor) Callbacks for Style Editor ---
    
    # ★ 1. Consolidate: Remove on_y1_series_select and on_y2_series_select
    # def on_y1_series_select(self, event=None): ... (removed)
    # def on_y2_series_select(self, event=None): ... (removed)

    # ★ 1. Consolidate: Add new combined callback
    def on_combined_series_select(self, event=None):
        """ Loads the style for the selected combined series (Y1 or Y2) into the editor """
        selected_item = self.combined_style_target_var.get()
        if not selected_item:
            return

        is_y1 = False
        series_name = ""

        if selected_item.startswith("(Y1) "):
            is_y1 = True
            series_name = selected_item[5:] # Get name after "(Y1) "
        elif selected_item.startswith("(Y2) "):
            is_y1 = False
            series_name = selected_item[5:] # Get name after "(Y2) "
        else:
            return # Should not happen

        if series_name:
            self.load_style_to_editor(series_name, is_y1=is_y1)

            
    def load_style_to_editor(self, series_name, is_y1):
        """ Helper to load a series' style into the 'current_style' vars """
        if series_name is None:
            # (★ MODIFIED) Clear editor
            self.current_style_color_var.set("#000000")
            self.current_style_linestyle_var.set("-")
            self.current_style_marker_var.set("o")
            self.current_style_linewidth_var.set(1.5)
            self.current_style_alpha_var.set(1.0)
            self.update_color_label(self.style_editor_color_label, "#000000")
            return

        styles_dict = self.y1_series_styles if is_y1 else self.y2_series_styles
        
        # Get the style for this series, or create a default if it's new
        series_style = self.get_or_create_default_style(series_name, styles_dict)
        
        # Set the editor's tk.Vars
        self.current_style_color_var.set(series_style.get('color', 'None')) # (★ MODIFIED) Default to 'None'
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

        if selected_item.startswith("(Y1) "):
            series_name = selected_item[5:]
            styles_dict = self.y1_series_styles
        elif selected_item.startswith("(Y2) "):
            series_name = selected_item[5:]
            styles_dict = self.y2_series_styles
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
        
        # Y-Axis (Listbox)
        self.y_listbox.delete(0, tk.END)
        self.y2_listbox.delete(0, tk.END)
        for col in columns:
            self.y_listbox.insert(tk.END, col)
            self.y2_listbox.insert(tk.END, col)
            
        self.plot_button['state'] = 'normal'
        self.export_button['state'] = 'normal'

        # Default selection
        if columns:
            self.x_axis_var.set(columns[0])
            self.xlabel_var.set(columns[0])
            if len(columns) > 1:
                self.y_listbox.select_set(1) # Select 2nd column by default
                self.ylabel_var.set(columns[1])
            else:
                self.y_listbox.select_set(0) # Select 1st column
                self.ylabel_var.set(columns[0])

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
        """Save all current settings to a JSON file"""
        if not self.data_file_path:
            # 1. UI English: Messagebox
            messagebox.showwarning("Save Error", "Please load a data file first before saving settings.")
            return
            
        # 1. UI English: Dialog title
        file_path = filedialog.asksaveasfilename(
            title="Save Settings",
            filetypes=[("JSON files", "*.json")],
            defaultextension=".json"
        )
        if not file_path:
            return

        # (★ MODIFIED) Get the directory where the settings file will be saved
        settings_dir = os.path.dirname(file_path)
        if not settings_dir:
            # If saving in the root or cwd, use "."
            settings_dir = os.curdir 

        # (★ MODIFIED) Calculate data path relative to the *settings file location*
        relative_data_path = self.data_file_path
        if self.data_file_path:
            try:
                # Calculate path relative to the settings directory
                relative_data_path = os.path.relpath(self.data_file_path, start=settings_dir)
            except ValueError:
                # Fallback (e.g., paths on different drives in Windows)
                relative_data_path = self.data_file_path # Save absolute path
            
        settings = {
            "data_file_path": relative_data_path,
            "plot_type": self.plot_type_var.get(),
            "x_axis": self.x_axis_var.get(),
            "y_axis_indices": list(self.y_listbox.curselection()),
            "y2_axis_indices": list(self.y2_listbox.curselection()),
            "title": self.title_var.get(),
            "xlabel": self.xlabel_var.get(),
            "ylabel": self.ylabel_var.get(),
            "ylabel2": self.ylabel2_var.get(),
            
            # --- (★ Style Refactor) ---
            # Save the series style dictionaries
            "y1_series_styles": self.y1_series_styles,
            "y2_series_styles": self.y2_series_styles,
            # --- (End Style Refactor) ---

            "grid": self.grid_var.get(),
            "marker": self.marker_var.get(),
            
            "font_family": self.font_family_var.get(),
            "title_fontsize": self.title_fontsize_var.get(),
            "xlabel_fontsize": self.xlabel_fontsize_var.get(),
            "ylabel_fontsize": self.ylabel_fontsize_var.get(),
            "ylabel2_fontsize": self.ylabel2_fontsize_var.get(),
            "tick_fontsize": self.tick_fontsize_var.get(),
            "tick2_fontsize": self.tick2_fontsize_var.get(),
            "fig_width": self.fig_width_var.get(),
            "fig_height": self.fig_height_var.get(),
            
            "xlim_min": self.xlim_min_var.get(),
            "xlim_max": self.xlim_max_var.get(),
            "ylim_min": self.ylim_min_var.get(),
            "ylim_max": self.ylim_max_var.get(),
            "ylim2_min": self.ylim2_min_var.get(),
            "ylim2_max": self.ylim2_max_var.get(),
            "xtick_show": self.xtick_show_var.get(),
            "xtick_label_show": self.xtick_label_show_var.get(),
            "xtick_direction": self.xtick_direction_var.get(),
            "ytick_show": self.ytick_show_var.get(),
            "ytick_label_show": self.ytick_label_show_var.get(),
            "ytick_direction": self.ytick_direction_var.get(),
            "ytick2_show": self.ytick2_show_var.get(),
            "ytick2_label_show": self.ytick2_label_show_var.get(),
            "ytick2_direction": self.ytick2_direction_var.get(),
            
            # (★ MODIFIED) 1. Scientific Notation Fix: Save settings
            "xaxis_plain_format": self.xaxis_plain_format_var.get(),
            "yaxis1_plain_format": self.yaxis1_plain_format_var.get(),
            "yaxis2_plain_format": self.yaxis2_plain_format_var.get(),
            
            "spine_top": self.spine_top_var.get(),
            "spine_bottom": self.spine_bottom_var.get(),
            "spine_left": self.spine_left_var.get(),
            "spine_right": self.spine_right_var.get(),
            "face_color": self.face_color_var.get(),
            "fig_color": self.fig_color_var.get(), # (★ ADDED)
            
            "legend_show": self.legend_show_var.get(),
            "legend_loc": self.legend_loc_var.get(),
            
            # (★ ADDED) Log scale settings
            "x_log_scale": self.x_log_scale_var.get(),
            "y1_log_scale": self.y1_log_scale_var.get(),
            "y2_log_scale": self.y2_log_scale_var.get(),

            # (★ MODIFIED) Add Invert Axis settings
            "x_invert": self.x_invert_var.get(),
            "y1_invert": self.y1_invert_var.get(),
            "y2_invert": self.y2_invert_var.get(),
        }
        
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
            # 1. UI English: Messagebox
            messagebox.showinfo("Success", f"Settings saved to {file_path}.")
        except Exception as e:
            # 1. UI English: Messagebox
            messagebox.showerror("Save Error", f"Failed to save settings:\n{e}")

    def load_settings(self):
        """Load settings from a JSON file"""
        # 1. UI English: Dialog title
        file_path = filedialog.askopenfilename(
            title="Select Settings File",
            filetypes=[("JSON files", "*.json")]
        )
        if not file_path:
            return

        # (★ MODIFIED) Get the directory *from which* the settings file is loaded
        settings_dir = os.path.dirname(file_path)
        if not settings_dir:
            settings_dir = os.curdir

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
        except Exception as e:
            # 1. UI English: Messagebox
            messagebox.showerror("Load Error", f"Failed to read settings file:\n{e}")
            return

        # 1. Load Data
        if 'data_file_path' not in settings:
            # 1. UI English: Messagebox
            messagebox.showerror("Error", "Settings file does not contain 'data_file_path'.")
            return
            
        data_path_from_settings = settings['data_file_path']
        if not data_path_from_settings:
            # 1. UI English: Messagebox
            messagebox.showerror("Error", "Settings file contains an empty 'data_file_path'.")
            return

        # (★ MODIFIED) Resolve the data path relative to the *settings file's directory*
        # os.path.join(A, B) correctly handles if B is an absolute path (it just returns B)
        # os.path.abspath() ensures the combined path is fully resolved
        absolute_data_path = os.path.abspath(os.path.join(settings_dir, data_path_from_settings))
            
        if not os.path.exists(absolute_data_path):
            # 1. UI English: Messagebox
            messagebox.showwarning("Data File Not Found", 
                                   f"The data file referenced in settings was not found:\n{data_path_from_settings}\n"
                                   f"(Looked relative to: {settings_dir})\n"
                                   f"(Resolved to: {absolute_data_path})\n\n"
                                   "Please load the correct data file manually, then load settings again if Y-axis selections are wrong.")
            pass
        else:
             self.load_data(file_path=absolute_data_path) # (★ MODIFIED) Use resolved path
        
        # 2. Apply settings to GUI (setters)
        self.set_variable_from_dict(self.plot_type_var, settings, 'plot_type')
        self.set_variable_from_dict(self.x_axis_var, settings, 'x_axis')
        self.set_variable_from_dict(self.title_var, settings, 'title')
        self.set_variable_from_dict(self.xlabel_var, settings, 'xlabel')
        self.set_variable_from_dict(self.ylabel_var, settings, 'ylabel')
        self.set_variable_from_dict(self.ylabel2_var, settings, 'ylabel2')
        
        # --- (★ Style Refactor) ---
        # Load the series style dictionaries
        if 'y1_series_styles' in settings:
            self.y1_series_styles = settings['y1_series_styles']
        else:
            self.y1_series_styles = {} # Initialize if not in file
            
        if 'y2_series_styles' in settings:
            self.y2_series_styles = settings['y2_series_styles']
        else:
            self.y2_series_styles = {}
        # --- (End Style Refactor) ---

        self.set_variable_from_dict(self.grid_var, settings, 'grid')
        self.set_variable_from_dict(self.marker_var, settings, 'marker')
        
        self.set_variable_from_dict(self.font_family_var, settings, 'font_family')
        self.set_variable_from_dict(self.title_fontsize_var, settings, 'title_fontsize')
        self.set_variable_from_dict(self.xlabel_fontsize_var, settings, 'xlabel_fontsize')
        self.set_variable_from_dict(self.ylabel_fontsize_var, settings, 'ylabel_fontsize')
        self.set_variable_from_dict(self.ylabel2_fontsize_var, settings, 'ylabel2_fontsize')
        self.set_variable_from_dict(self.tick_fontsize_var, settings, 'tick_fontsize')
        self.set_variable_from_dict(self.tick2_fontsize_var, settings, 'tick2_fontsize')
        self.set_variable_from_dict(self.fig_width_var, settings, 'fig_width')
        self.set_variable_from_dict(self.fig_height_var, settings, 'fig_height')
        
        self.set_variable_from_dict(self.xlim_min_var, settings, 'xlim_min')
        self.set_variable_from_dict(self.xlim_max_var, settings, 'xlim_max')
        self.set_variable_from_dict(self.ylim_min_var, settings, 'ylim_min')
        self.set_variable_from_dict(self.ylim_max_var, settings, 'ylim_max')
        self.set_variable_from_dict(self.ylim2_min_var, settings, 'ylim2_min')
        self.set_variable_from_dict(self.ylim2_max_var, settings, 'ylim2_max')
        self.set_variable_from_dict(self.xtick_show_var, settings, 'xtick_show')
        self.set_variable_from_dict(self.xtick_label_show_var, settings, 'xtick_label_show')
        self.set_variable_from_dict(self.xtick_direction_var, settings, 'xtick_direction')
        self.set_variable_from_dict(self.ytick_show_var, settings, 'ytick_show')
        self.set_variable_from_dict(self.ytick_label_show_var, settings, 'ytick_label_show')
        self.set_variable_from_dict(self.ytick_direction_var, settings, 'ytick_direction')
        self.set_variable_from_dict(self.ytick2_show_var, settings, 'ytick2_show')
        self.set_variable_from_dict(self.ytick2_label_show_var, settings, 'ytick2_label_show')
        self.set_variable_from_dict(self.ytick2_direction_var, settings, 'ytick2_direction')
        
        # (★ MODIFIED) 1. Scientific Notation Fix: Load setting
        self.set_variable_from_dict(self.xaxis_plain_format_var, settings, 'xaxis_plain_format')
        self.set_variable_from_dict(self.yaxis1_plain_format_var, settings, 'yaxis1_plain_format', fallback_key='yaxis_plain_format')
        self.set_variable_from_dict(self.yaxis2_plain_format_var, settings, 'yaxis2_plain_format')

        self.set_variable_from_dict(self.spine_top_var, settings, 'spine_top')
        self.set_variable_from_dict(self.spine_bottom_var, settings, 'spine_bottom')
        self.set_variable_from_dict(self.spine_left_var, settings, 'spine_left')
        self.set_variable_from_dict(self.spine_right_var, settings, 'spine_right')
        self.set_variable_from_dict(self.face_color_var, settings, 'face_color')
        self.set_variable_from_dict(self.fig_color_var, settings, 'fig_color') # (★ ADDED)
        
        self.set_variable_from_dict(self.legend_show_var, settings, 'legend_show')
        self.set_variable_from_dict(self.legend_loc_var, settings, 'legend_loc')
        
        # (★ ADDED) Log scale settings
        self.set_variable_from_dict(self.x_log_scale_var, settings, 'x_log_scale')
        self.set_variable_from_dict(self.y1_log_scale_var, settings, 'y1_log_scale')
        self.set_variable_from_dict(self.y2_log_scale_var, settings, 'y2_log_scale')

        # (★ MODIFIED) Load Invert Axis settings
        self.set_variable_from_dict(self.x_invert_var, settings, 'x_invert')
        self.set_variable_from_dict(self.y1_invert_var, settings, 'y1_invert')
        self.set_variable_from_dict(self.y2_invert_var, settings, 'y2_invert')

        # Restore Listbox selections (only if data is loaded)
        if self.df is not None:
            self.y_listbox.select_clear(0, tk.END)
            if 'y_axis_indices' in settings:
                for i in settings['y_axis_indices']:
                    if i < self.y_listbox.size():
                        self.y_listbox.select_set(i)
                    
            self.y2_listbox.select_clear(0, tk.END)
            if 'y2_axis_indices' in settings:
                for i in settings['y2_axis_indices']:
                    if i < self.y2_listbox.size():
                        self.y2_listbox.select_set(i)
        else:
             # 1. UI English: Messagebox
             messagebox.showwarning("Data Mismatch", "Y-axis selections were not restored because data file was not loaded.")

        # Update color labels
        self.update_color_label(self.face_color_label, self.face_color_var.get())
        self.update_color_label(self.fig_color_label, self.fig_color_var.get()) # (★ ADDED)
        # (Style editor labels will be updated when a series is selected)

        # 3. Redraw graph (if data is loaded)
        if self.df is not None:
            self.plot_graph()
            # 1. UI English: Messagebox
            messagebox.showinfo("Success", f"Settings loaded from {file_path}.")
        else:
            # 1. UI English: Messagebox
            messagebox.showinfo("Settings Applied", f"Settings from {file_path} applied. Data was not loaded.")


    def set_variable_from_dict(self, var, settings_dict, key, fallback_key=None):
        """ Set tk.Variable from a settings dictionary key, with existence check and fallback """
        # (★ MODIFIED) 9. Add fallback logic
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
            # (★ MODIFIED) Set anchor and update text
            label.config(background=color_code, text=text, anchor=tk.CENTER)
        except tk.TclError: # Handle invalid color code
            label.config(background="#FFFFFF", text="Invalid", anchor=tk.CENTER)

    # --- Plotting Method (v5.0 Series Style) ---
    def plot_graph(self):
        
        # (1) Clear figure first
        try:
            self.fig.clear()
            self.ax = self.fig.add_subplot(111) # Recreate main axis
            self.ax2 = None # Reset 2nd axis
        except Exception as e:
            # 1. UI English: Messagebox
            messagebox.showerror("Internal Error", f"Failed to clear graph:\n{e}")
            return

        # 1. Get Data
        self.get_data_from_sheet()
        if self.df is None or self.df.empty:
            # 1. UI English: Messagebox
            messagebox.showinfo("Info", "No data to plot.")
            self.canvas.draw()
            return

        # 2. Get All Settings
        x_col = self.x_axis_var.get()
        plot_type = self.plot_type_var.get()
        
        try:
            y_cols_1 = [self.y_listbox.get(i) for i in self.y_listbox.curselection()]
            y_cols_2 = [self.y2_listbox.get(i) for i in self.y2_listbox.curselection()]
        except tk.TclError: 
            y_cols_1 = []
            y_cols_2 = []
            
        # --- (★ Style Refactor) Update Style Tab Comboboxes ---
        # ★ 1. Consolidate: Update the single combined combobox
        
        # Build the combined list with prefixes
        combined_list = []
        for col in y_cols_1:
            combined_list.append(f"(Y1) {col}")
        for col in y_cols_2:
            combined_list.append(f"(Y2) {col}")
            
        # Set the new values
        self.style_combo['values'] = combined_list

        # Clear the current selection if it's no longer valid
        current_selection = self.combined_style_target_var.get()
        if current_selection not in combined_list:
            self.combined_style_target_var.set("")
            # (★ MODIFIED) Clear editor if selection is invalid
            self.load_style_to_editor(None, True) # Clear editor
            
        # ★ 1. Consolidate: Remove old combo updates
        # self.y1_style_combo['values'] = y_cols_1
        # self.y2_style_combo['values'] = y_cols_2
        # if self.y1_style_target_var.get() not in y_cols_1:
        #     self.y1_style_target_var.set("")
        # if self.y2_style_target_var.get() not in y_cols_2:
        #     self.y2_style_target_var.set("")
        
        # --- (End Style Refactor) ---


        # --- Column Existence Check ---
        if not x_col:
            # 1. UI English: Messagebox
            messagebox.showerror("Error", "Please select an X-axis column.")
            self.canvas.draw()
            return
        if not y_cols_1 and not y_cols_2:
            # 1. UI English: Messagebox
            messagebox.showerror("Error", "Please select data for Y-Axis (Left) or Y-Axis (Right).")
            self.canvas.draw()
            return
        
        all_cols = [x_col] + y_cols_1 + y_cols_2
        for col in all_cols:
            if col not in self.df.columns:
                # 1. UI English: Messagebox
                messagebox.showerror("Plot Error", f"Selected column '{col}' does not exist in data.")
                self.canvas.draw()
                return

        # 3. Prepare Graph (Size & Font)
        try:
            self.fig.set_size_inches(self.fig_width_var.get(), self.fig_height_var.get())
            self.fig.set_facecolor(self.fig_color_var.get()) # (★ ADDED)
            matplotlib.rcParams['font.family'] = self.font_family_var.get()
            
            # Create 2nd Y-axis if needed
            if y_cols_2:
                self.ax2 = self.ax.twinx()
                
        except Exception as e:
            # 1. UI English: Messagebox
            messagebox.showerror("Settings Error", f"Failed to apply basic graph settings:\n{e}")
            return

        # 4. Plot Graph
        try:
            x_data_raw = self.df[x_col]
            
            # Try numeric conversion for X-axis (except for bar)
            x_data_numeric = None
            if plot_type != "bar":
                x_data_cleaned = x_data_raw.astype(str).str.replace(r'[^\d.-]', '', regex=True)
                x_data_numeric = pd.to_numeric(x_data_cleaned, errors='coerce')

            # --- Internal Plotting Function (★ Style Refactor) ---
            def plot_series(ax, y_col, x_data_raw, is_twin_ax):
                
                # Get the style dictionary for this series
                styles_dict = self.y1_series_styles if not is_twin_ax else self.y2_series_styles
                # Get or create the specific style entry for this column
                series_style = self.get_or_create_default_style(y_col, styles_dict)

                # Prepare Y-data (clean and convert to numeric)
                y_data_cleaned = self.df[y_col].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                y_data_numeric = pd.to_numeric(y_data_cleaned, errors='coerce')

                # Get style properties, providing defaults
                # IMPORTANT: Use .get(key, default)
                color = series_style.get('color', None)
                if color == 'None' or color is None: # Allow 'None' string or Python None
                    color = None # Let matplotlib auto-assign
                    
                linestyle = series_style.get('linestyle', '-')
                if linestyle == 'None': linestyle = 'None' # String 'None' is valid for no line

                linewidth = series_style.get('linewidth', 1.5)
                alpha = series_style.get('alpha', 1.0)
                
                # Marker style depends on the global "Show Markers" checkbox
                markerstyle = series_style.get('marker', 'o')
                if not self.marker_var.get() or markerstyle == 'None':
                    markerstyle = 'None' # 'None' string is valid for no marker

                if plot_type == "bar":
                    x_data = x_data_raw.astype(str) # Bar uses X as string
                    valid_mask = ~y_data_numeric.isnull()
                    plot_x = x_data[valid_mask]
                    plot_y = y_data_numeric[valid_mask]
                    
                    if plot_y.empty:
                        return
                    
                    kwargs = {'alpha': alpha, 'label': y_col}
                    
                    if color:
                        kwargs['color'] = color
                        
                    ax.bar(plot_x, plot_y, **kwargs)

                else: # line / scatter
                    valid_data = pd.DataFrame({'x': x_data_numeric, 'y': y_data_numeric}).dropna()
                    
                    if valid_data.empty:
                        return

                    plot_x = valid_data['x']
                    plot_y = valid_data['y']
                    
                    kwargs = {
                        'marker': markerstyle,
                        'alpha': alpha,
                        'label': y_col
                    }
                    
                    if color:
                         kwargs['color'] = color

                    if plot_type == "line":
                        kwargs.update({
                            'linestyle': linestyle,
                            'linewidth': linewidth,
                        })
                        ax.plot(plot_x, plot_y, **kwargs)
                    elif plot_type == "scatter":
                         kwargs['marker'] = markerstyle if markerstyle != 'None' else 'o'
                         ax.scatter(plot_x, plot_y, **kwargs)

            # --- Execute Plotting ---
            
            # (★ Style Refactor) Remove all the old logic about y1_color_to_use etc.
            # Just loop and plot.
            
            # 1st Y-Axis
            for y_col in y_cols_1:
                plot_series(self.ax, y_col, x_data_raw, is_twin_ax=False)

            # 2nd Y-Axis
            if self.ax2:
                for y_col in y_cols_2:
                    plot_series(self.ax2, y_col, x_data_raw, is_twin_ax=True)
            
            
            # 5. Apply All Settings
            
            # (★ ADDED) Log Scale Settings
            try:
                self.ax.set_xscale('log' if self.x_log_scale_var.get() else 'linear')
            except ValueError:
                self.x_log_scale_var.set(False) # Uncheck if log fails
                self.ax.set_xscale('linear') # Fallback if log scale fails (e.g., zero/negative data)
            try:
                self.ax.set_yscale('log' if self.y1_log_scale_var.get() else 'linear')
            except ValueError:
                 self.y1_log_scale_var.set(False) # Uncheck if log fails
                 self.ax.set_yscale('linear')

            # Labels and Title
            # (★ MODIFIED) 4. Apply font family explicitly
            font_family = self.font_family_var.get()
            self.ax.set_xlabel(self.xlabel_var.get() if self.xlabel_var.get() else x_col, fontsize=self.xlabel_fontsize_var.get(), fontfamily=font_family)
            self.ax.set_ylabel(self.ylabel_var.get() if self.ylabel_var.get() else ", ".join(y_cols_1), fontsize=self.ylabel_fontsize_var.get(), fontfamily=font_family)
            self.ax.set_title(self.title_var.get(), fontsize=self.title_fontsize_var.get(), fontfamily=font_family)
            
            # Grid
            self.ax.grid(self.grid_var.get())
            
            # Axis limits (X, Y1)
            self.set_axis_limits(self.ax, 'x', self.xlim_min_var.get(), self.xlim_max_var.get())
            self.set_axis_limits(self.ax, 'y', self.ylim_min_var.get(), self.ylim_max_var.get())

            # (★ MODIFIED) Apply Invert Axis (X, Y1)
            if self.x_invert_var.get():
                self.ax.invert_xaxis()
            if self.y1_invert_var.get():
                self.ax.invert_yaxis()

            # Tick settings (X, Y1)
            self.ax.tick_params(axis='x', which='both', 
                                direction=self.xtick_direction_var.get(), 
                                bottom=self.xtick_show_var.get(),
                                labelbottom=self.xtick_label_show_var.get(),
                                labelsize=self.tick_fontsize_var.get())
            self.ax.tick_params(axis='y', which='both', 
                                direction=self.ytick_direction_var.get(), 
                                left=self.ytick_show_var.get(), 
                                labelleft=self.ytick_label_show_var.get(),
                                labelsize=self.tick_fontsize_var.get())
            
            # (★ MODIFIED) 6. Explicitly set tick label font family
            font_family = self.font_family_var.get()
            for label in self.ax.get_xticklabels() + self.ax.get_yticklabels():
                label.set_fontfamily(font_family)
            
            # (★ MODIFIED) 1. Scientific Notation Fix: Apply plain format if checked
            if self.xaxis_plain_format_var.get():
                self.ax.ticklabel_format(style='plain', axis='x', useOffset=False)
            if self.yaxis1_plain_format_var.get():
                self.ax.ticklabel_format(style='plain', axis='y', useOffset=False)

                                
            # Spines and Background Color (Axes)
            self.ax.set_facecolor(self.face_color_var.get())
            self.ax.spines['top'].set_visible(self.spine_top_var.get())
            self.ax.spines['bottom'].set_visible(self.spine_bottom_var.get())
            self.ax.spines['left'].set_visible(self.spine_left_var.get())
            
            # 2nd Y-Axis Settings
            if self.ax2:
                # (★ ADDED) Log Scale Setting
                try:
                    self.ax2.set_yscale('log' if self.y2_log_scale_var.get() else 'linear')
                except ValueError:
                    self.y2_log_scale_var.set(False) # Uncheck if log fails
                    self.ax2.set_yscale('linear')
                
                # (★ MODIFIED) 4. Apply font family explicitly
                self.ax2.set_ylabel(self.ylabel2_var.get() if self.ylabel2_var.get() else ", ".join(y_cols_2), fontsize=self.ylabel2_fontsize_var.get(), fontfamily=font_family)
                self.set_axis_limits(self.ax2, 'y', self.ylim2_min_var.get(), self.ylim2_max_var.get())
                
                # (★ MODIFIED) Apply Invert Axis (Y2)
                if self.y2_invert_var.get():
                    self.ax2.invert_yaxis()
                
                self.ax2.tick_params(axis='y', which='both',
                                     direction=self.ytick2_direction_var.get(),
                                     right=self.ytick2_show_var.get(),
                                     labelright=self.ytick2_label_show_var.get(),
                                     labelsize=self.tick2_fontsize_var.get())
                
                # (★ MODIFIED) 6. Explicitly set ax2 tick label font family
                for label in self.ax2.get_yticklabels():
                    label.set_fontfamily(font_family)
                
                # (★ MODIFIED) 1. Scientific Notation Fix: Apply plain format if checked
                if self.yaxis2_plain_format_var.get():
                    self.ax2.ticklabel_format(style='plain', axis='y', useOffset=False)
                
                # Right spine visibility controlled by ax2
                self.ax.spines['right'].set_visible(False)
                self.ax2.spines['top'].set_visible(self.spine_top_var.get())
                self.ax2.spines['bottom'].set_visible(self.spine_bottom_var.get())
                self.ax2.spines['left'].set_visible(False)
                self.ax2.spines['right'].set_visible(self.spine_right_var.get())
            else:
                # No 2nd Y-axis
                self.ax.spines['right'].set_visible(self.spine_right_var.get())
            
            # Legend (Combined)
            if self.legend_show_var.get():
                h1, l1 = self.ax.get_legend_handles_labels()
                h2, l2 = [], []
                if self.ax2:
                    h2, l2 = self.ax2.get_legend_handles_labels()
                # (★ MODIFIED) 4. Apply font family to legend via prop
                legend_font_props = {'family': font_family, 'size': self.tick_fontsize_var.get()}
                self.ax.legend(handles=h1+h2, labels=l1+l2, 
                               loc=self.legend_loc_var.get(), 
                               prop=legend_font_props)
            
            self.fig.tight_layout()

            # 6. Update Canvas
            self.canvas.draw()

            # (★ MODIFIED) 2. Scrollbar Fix: Force update widget size and manually trigger scroll region update
            self.graph_frame.update_idletasks() # Ensure sizes are calculated
            self.on_graph_frame_configure(None) # Call the update function

            # (Original code removed)
            # self.graph_frame.update_idletasks()
            # scrollable_canvas = self.graph_frame.master
            # scrollable_canvas.configure(scrollregion=scrollable_canvas.bbox("all"))


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
        except ValueError:
            pass # Ignore if conversion fails

    def export_graph(self):
        # 1. UI English: Dialog title
        file_path = filedialog.asksaveasfilename(
            title="Save Graph",
            filetypes=[("PNG files", "*.png"),
                       ("SVG files", "*.svg"),
                       ("PDF files", ".*pdf")],
        )
        if not file_path:
            return

        try:
            matplotlib.rcParams['font.family'] = self.font_family_var.get()
            # (★ MODIFIED) Save with the set facecolor (using fig_color_var)
            self.fig.savefig(file_path, bbox_inches='tight', facecolor=self.fig_color_var.get())
            # 1. UI English: Messagebox
            messagebox.showinfo("Success", f"Graph saved to {file_path}.")
        except Exception as e:
            # 1. UI English: Messagebox
            messagebox.showerror("Save Error", f"Failed to save graph:\n{e}")

if __name__ == "__main__":
    app = GraphApp()
    app.mainloop()
