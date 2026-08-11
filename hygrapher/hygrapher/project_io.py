# -*- coding: utf-8 -*-
"""
hygrapher.project_io

Project File (.pmggrp) Serialization & Deserialization Module.
"""

import json
from datetime import datetime


def _safe_get(var, default=""):
    if var is None:
        return default
    try:
        val = var.get()
        if isinstance(val, (str, int, float, bool, list, dict)):
            return val
        return str(val)
    except Exception:
        return default


def _safe_dict(val):
    return val if isinstance(val, dict) else {}


def _safe_list(val):
    return val if isinstance(val, list) else []



def build_project_dict(app, version_str="0.6.0", dimension="2D"):
    """
    Consolidate application state into a project dictionary for .pmggrp export.
    """
    if hasattr(app, 'get_data_from_sheet'):
        try:
            app.get_data_from_sheet()
        except Exception:
            pass

    data_dict = None
    if getattr(app, 'df', None) is not None and not app.df.empty:
        data_dict = {
            "columns": app.df.columns.tolist(),
            "data": app.df.values.tolist()
        }

    x_tabs_data = []
    if hasattr(app, 'get_x_tabs_data'):
        try:
            x_tabs_data = app.get_x_tabs_data()
        except Exception:
            x_tabs_data = []

    y1_idx = []
    if hasattr(app, 'y_listbox') and hasattr(app.y_listbox, 'curselection'):
        try:
            y1_idx = list(app.y_listbox.curselection())
        except Exception:
            y1_idx = []

    y2_idx = []
    if hasattr(app, 'y2_listbox') and hasattr(app.y2_listbox, 'curselection'):
        try:
            y2_idx = list(app.y2_listbox.curselection())
        except Exception:
            y2_idx = []

    settings = {
        "format": "Python Matplotlib Grapher App (HYGrapher) Graph Project",
        "version": "1.1",
        "application": "HYGrapher",
        "application_version": version_str,
        "dimension": dimension,
        "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "edited_data": data_dict,
        "original_file_path": str(getattr(app, 'data_file_path', '') or ''),
        
        "plot_type": _safe_get(getattr(app, 'plot_type_var', None), "line"),
        "x_axis": _safe_get(getattr(app, 'x_axis_var', None), ""),
        "y_axis_indices": _safe_list(y1_idx),
        "y2_axis_indices": _safe_list(y2_idx),
        "x_tabs": _safe_list(x_tabs_data),
        
        "title": _safe_get(getattr(app, 'title_var', None), ""),
        "xlabel": _safe_get(getattr(app, 'xlabel_var', None), ""),
        "ylabel": _safe_get(getattr(app, 'ylabel_var', None), ""),
        "ylabel2": _safe_get(getattr(app, 'ylabel2_var', None), ""),
        
        "y1_series_styles": _safe_dict(getattr(app, 'y1_series_styles', None)),
        "y2_series_styles": _safe_dict(getattr(app, 'y2_series_styles', None)),


        "grid": _safe_get(getattr(app, 'grid_var', None), False),
        "marker": _safe_get(getattr(app, 'marker_var', None), True),
        
        "font_family": _safe_get(getattr(app, 'font_family_var', None), "sans-serif"),
        "title_fontsize": _safe_get(getattr(app, 'title_fontsize_var', None), 16.0),
        "xlabel_fontsize": _safe_get(getattr(app, 'xlabel_fontsize_var', None), 14.0),
        "ylabel_fontsize": _safe_get(getattr(app, 'ylabel_fontsize_var', None), 14.0),
        "ylabel2_fontsize": _safe_get(getattr(app, 'ylabel2_fontsize_var', None), 14.0),
        "tick_fontsize": _safe_get(getattr(app, 'tick_fontsize_var', None), 14.0),
        "tick2_fontsize": _safe_get(getattr(app, 'tick2_fontsize_var', None), 14.0),
        "legend_fontsize": _safe_get(getattr(app, 'legend_fontsize_var', None), 12.0),
        "fig_width": _safe_get(getattr(app, 'fig_width_var', None), 7.0),
        "fig_height": _safe_get(getattr(app, 'fig_height_var', None), 6.0),
        
        "xlim_min": _safe_get(getattr(app, 'xlim_min_var', None), ""),
        "xlim_max": _safe_get(getattr(app, 'xlim_max_var', None), ""),
        "ylim_min": _safe_get(getattr(app, 'ylim_min_var', None), ""),
        "ylim_max": _safe_get(getattr(app, 'ylim_max_var', None), ""),
        "ylim2_min": _safe_get(getattr(app, 'ylim2_min_var', None), ""),
        "ylim2_max": _safe_get(getattr(app, 'ylim2_max_var', None), ""),
        
        "xtick_show": _safe_get(getattr(app, 'xtick_show_var', None), True),
        "xtick_label_show": _safe_get(getattr(app, 'xtick_label_show_var', None), True),
        "xtick_direction": _safe_get(getattr(app, 'xtick_direction_var', None), "out"),
        "ytick_show": _safe_get(getattr(app, 'ytick_show_var', None), True),
        "ytick_label_show": _safe_get(getattr(app, 'ytick_label_show_var', None), True),
        "ytick_direction": _safe_get(getattr(app, 'ytick_direction_var', None), "out"),
        "ytick2_show": _safe_get(getattr(app, 'ytick2_show_var', None), True),
        "ytick2_label_show": _safe_get(getattr(app, 'ytick2_label_show_var', None), True),
        "ytick2_direction": _safe_get(getattr(app, 'ytick2_direction_var', None), "out"),
        
        "xtick_minor_show": _safe_get(getattr(app, 'xtick_minor_show_var', None), False),
        "ytick_minor_show": _safe_get(getattr(app, 'ytick_minor_show_var', None), False),
        "ytick2_minor_show": _safe_get(getattr(app, 'ytick2_minor_show_var', None), False),
        "xtick_minor_interval": _safe_get(getattr(app, 'xtick_minor_interval_var', None), ""),
        "ytick_minor_interval": _safe_get(getattr(app, 'ytick_minor_interval_var', None), ""),
        "ytick2_minor_interval": _safe_get(getattr(app, 'ytick2_minor_interval_var', None), ""),
        
        "xaxis_plain_format": _safe_get(getattr(app, 'xaxis_plain_format_var', None), False),
        "yaxis1_plain_format": _safe_get(getattr(app, 'yaxis1_plain_format_var', None), False),
        "yaxis2_plain_format": _safe_get(getattr(app, 'yaxis2_plain_format_var', None), False),
        
        "xtick_major_interval": _safe_get(getattr(app, 'xtick_major_interval_var', None), ""),
        "ytick_major_interval": _safe_get(getattr(app, 'ytick_major_interval_var', None), ""),
        "ytick2_major_interval": _safe_get(getattr(app, 'ytick2_major_interval_var', None), ""),
        
        "spine_top": _safe_get(getattr(app, 'spine_top_var', None), True),
        "spine_bottom": _safe_get(getattr(app, 'spine_bottom_var', None), True),
        "spine_left": _safe_get(getattr(app, 'spine_left_var', None), True),
        "spine_right": _safe_get(getattr(app, 'spine_right_var', None), True),
        "face_color": _safe_get(getattr(app, 'face_color_var', None), "#FFFFFF"),
        "fig_color": _safe_get(getattr(app, 'fig_color_var', None), "#FFFFFF"),
        
        "legend_show": _safe_get(getattr(app, 'legend_show_var', None), False),
        "legend_loc": _safe_get(getattr(app, 'legend_loc_var', None), "best"),
        
        "x_log_scale": _safe_get(getattr(app, 'x_log_scale_var', None), False),
        "y1_log_scale": _safe_get(getattr(app, 'y1_log_scale_var', None), False),
        "y2_log_scale": _safe_get(getattr(app, 'y2_log_scale_var', None), False),

        "x_invert": _safe_get(getattr(app, 'x_invert_var', None), False),
        "y1_invert": _safe_get(getattr(app, 'y1_invert_var', None), False),
        "y2_invert": _safe_get(getattr(app, 'y2_invert_var', None), False),
        
        "enable_smoothing": _safe_get(getattr(app, 'enable_smoothing_var', None), False),
        "smoothing_window": _safe_get(getattr(app, 'smoothing_window_var', None), 5),
        "enable_errorbar": _safe_get(getattr(app, 'enable_errorbar_var', None), False),
        "errorbar_column": _safe_get(getattr(app, 'errorbar_column_var', None), ""),
        "enable_annotation": _safe_get(getattr(app, 'enable_annotation_var', None), False),
        
        "data_filter_enabled": _safe_get(getattr(app, 'data_filter_enabled_var', None), False),
        "filter_min": _safe_get(getattr(app, 'filter_min_var', None), ""),
        "filter_max": _safe_get(getattr(app, 'filter_max_var', None), ""),
        "filter_column": _safe_get(getattr(app, 'filter_column_var', None), ""),
        "grid_alpha": _safe_get(getattr(app, 'grid_alpha_var', None), 0.3),
        "grid_linestyle": _safe_get(getattr(app, 'grid_linestyle_var', None), "--"),
        "grid_linewidth": _safe_get(getattr(app, 'grid_linewidth_var', None), 0.5),
        "subplot_mode": _safe_get(getattr(app, 'subplot_mode_var', None), False),
        "rotate_labels": _safe_get(getattr(app, 'rotate_labels_var', None), False),
        "rotation_angle": _safe_get(getattr(app, 'rotation_angle_var', None), 45),

        "export_dpi": _safe_get(getattr(app, 'export_dpi_var', None), 150),
        "colormap": _safe_get(getattr(app, 'colormap_var', None), "viridis"),
        "tight_layout_pad": _safe_get(getattr(app, 'tight_layout_pad_var', None), 1.0),
    }

    if dimension == "3D":
        settings.update({
            "y_axis": _safe_get(getattr(app, 'y_axis_var', None), ""),
            "z_axis_indices": y1_idx,
            "zlabel": _safe_get(getattr(app, 'zlabel_var', None), ""),
            "z_log_scale": _safe_get(getattr(app, 'z_log_scale_var', None), False),
            "z_invert": _safe_get(getattr(app, 'z_invert_var', None), False),
            "view_elev": _safe_get(getattr(app, 'view_elev_var', None), 30),
            "view_azim": _safe_get(getattr(app, 'view_azim_var', None), -60),
            "mesh_resolution": _safe_get(getattr(app, 'mesh_resolution_var', None), 50),
        })

    return settings


def save_project_file(app, file_path, version_str="0.6.0", dimension="2D"):
    """
    Save current app state to a JSON file.
    """
    settings = build_project_dict(app, version_str=version_str, dimension=dimension)
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(settings, f, indent=4)
    return settings
