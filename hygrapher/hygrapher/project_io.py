# -*- coding: utf-8 -*-
"""
hygrapher.project_io

Project File (.pmggrp) Serialization & Deserialization Module.

Maps the PyQt6 widget state on ``GraphApp`` (and ``GraphApp3D``) to/from a
plain JSON-serializable dict. Every entry in ``_FIELD_SPEC`` names the real
widget attribute set in main.py / main_3d.py — keep this list in sync with
the UI or a round trip through Save/Load Project will silently drop that
setting.
"""

import json
from datetime import datetime

from hygrapher.utils import get_app_version

# (settings_key, widget_attr, kind, default)
# kind: "text" (QLineEdit), "check" (QCheckBox), "int" (QSpinBox),
#       "float" (QDoubleSpinBox), "combo" (QComboBox by text)
_FIELD_SPEC = [
    ("title", "title_input", "text", ""),
    ("xlabel", "xlabel_input", "text", ""),
    ("ylabel", "ylabel_input", "text", ""),
    ("ylabel2", "ylabel2_input", "text", ""),
    ("plot_type", "plot_type_combo", "combo", "line"),
    ("x_log_scale", "x_log_check", "check", False),
    ("y1_log_scale", "y1_log_check", "check", False),
    ("y2_log_scale", "y2_log_check", "check", False),
    ("y1_invert", "y1_invert_check", "check", False),
    ("y2_invert", "y2_invert_check", "check", False),
    ("grid", "grid_check", "check", True),
    ("subplot_mode", "subplot_mode_check", "check", False),
    ("xlim_min", "xlim_min_input", "text", ""),
    ("xlim_max", "xlim_max_input", "text", ""),
    ("ylim_min", "ylim_min_input", "text", ""),
    ("ylim_max", "ylim_max_input", "text", ""),
    ("ylim2_min", "ylim2_min_input", "text", ""),
    ("ylim2_max", "ylim2_max_input", "text", ""),
    ("xtick_major_interval", "xtick_major_interval_input", "text", ""),
    ("ytick_major_interval", "ytick_major_interval_input", "text", ""),
    ("ytick2_major_interval", "ytick2_major_interval_input", "text", ""),
    ("xtick_minor_show", "xtick_minor_check", "check", False),
    ("xtick_minor_interval", "xtick_minor_interval_input", "text", ""),
    ("ytick_minor_show", "ytick_minor_check", "check", False),
    ("ytick_minor_interval", "ytick_minor_interval_input", "text", ""),
    ("ytick2_minor_show", "ytick2_minor_check", "check", False),
    ("ytick2_minor_interval", "ytick2_minor_interval_input", "text", ""),
    ("rotate_labels", "rotate_labels_check", "check", False),
    ("rotation_angle", "rotation_angle_spin", "int", 45),
    ("xaxis_plain_format", "xaxis_plain_check", "check", False),
    ("yaxis1_plain_format", "yaxis1_plain_check", "check", False),
    ("yaxis2_plain_format", "yaxis2_plain_check", "check", False),
    ("font_family", "font_family_combo", "combo", "sans-serif"),
    ("title_fontsize", "title_fontsize_spin", "int", 14),
    ("xlabel_fontsize", "xlabel_fontsize_spin", "int", 12),
    ("ylabel_fontsize", "ylabel_fontsize_spin", "int", 12),
    ("ylabel2_fontsize", "ylabel2_fontsize_spin", "int", 12),
    ("tick_fontsize", "tick_fontsize_spin", "int", 10),
    ("tick2_fontsize", "tick2_fontsize_spin", "int", 10),
    ("legend_fontsize", "legend_fontsize_spin", "int", 10),
    ("spine_top", "spine_top_check", "check", True),
    ("spine_bottom", "spine_bottom_check", "check", True),
    ("spine_left", "spine_left_check", "check", True),
    ("spine_right", "spine_right_check", "check", True),
    ("legend_show", "legend_show_check", "check", False),
    ("legend_loc", "legend_loc_combo", "combo", "best"),
    ("enable_smoothing", "enable_smoothing_check", "check", False),
    ("smoothing_window", "smoothing_window_spin", "int", 5),
    ("enable_errorbar", "enable_errorbar_check", "check", False),
    ("errorbar_column", "errorbar_column_combo", "combo", ""),
    ("enable_annotation", "enable_annotation_check", "check", False),
    ("data_filter_enabled", "data_filter_check", "check", False),
    ("filter_column", "filter_column_combo", "combo", ""),
    ("filter_min", "filter_min_input", "text", ""),
    ("filter_max", "filter_max_input", "text", ""),
    ("colormap", "colormap_combo", "combo", "viridis"),
    ("export_dpi", "export_dpi_spin", "int", 300),
    ("grid_alpha", "grid_alpha_spin", "float", 0.3),
    ("grid_linestyle", "grid_linestyle_combo", "combo", "--"),
    ("grid_linewidth", "grid_linewidth_spin", "float", 0.5),
]

# Same idea for the 3D plotter window (GraphApp3D).
_FIELD_SPEC_3D = [
    ("title", "title_input", "text", ""),
    ("plot_type", "plot_type_combo", "combo", "surface"),
    ("x_axis", "x_axis_combo", "combo", ""),
    ("y_axis", "y_axis_combo", "combo", ""),
    ("view_elev", "elev_spin", "int", 30),
    ("view_azim", "azim_spin", "int", -60),
    ("mesh_resolution", "resolution_spin", "int", 30),
    ("colormap", "colormap_combo", "combo", "viridis"),
]


def _get_field(app, attr, kind, default):
    widget = getattr(app, attr, None)
    if widget is None:
        return default
    try:
        if kind == "text":
            return widget.text()
        if kind == "check":
            return bool(widget.isChecked())
        if kind in ("int", "float"):
            return widget.value()
        if kind == "combo":
            return widget.currentText()
    except Exception:
        return default
    return default


def _set_field(app, attr, kind, value):
    widget = getattr(app, attr, None)
    if widget is None or value is None:
        return
    try:
        if kind == "text":
            widget.setText(str(value))
        elif kind == "check":
            widget.setChecked(bool(value))
        elif kind == "int":
            widget.setValue(int(value))
        elif kind == "float":
            widget.setValue(float(value))
        elif kind == "combo":
            idx = widget.findText(str(value))
            if idx >= 0:
                widget.setCurrentIndex(idx)
            else:
                widget.setCurrentText(str(value))
    except Exception:
        pass


def _spec_for(dimension):
    return _FIELD_SPEC_3D if dimension == "3D" else _FIELD_SPEC


def reset_to_defaults(app, dimension="2D"):
    """
    Reset every widget named in the field spec back to its factory default.
    Used by "Reset All" so it actually resets everything the UI exposes,
    instead of the handful of fields each window used to reset by hand.
    """
    for _key, attr, kind, default in _spec_for(dimension):
        _set_field(app, attr, kind, default)


def build_project_dict(app, version_str=None, dimension="2D"):
    """
    Consolidate application state into a project dictionary for .pmggrp export.
    """
    if version_str is None:
        version_str = get_app_version()

    if hasattr(app, "get_data_from_table"):
        try:
            app.get_data_from_table()
        except Exception:
            pass

    data_dict = None
    if getattr(app, "df", None) is not None and not app.df.empty:
        data_dict = {"columns": app.df.columns.tolist(), "data": app.df.values.tolist()}

    x_tabs_data = []
    if hasattr(app, "get_x_tabs_data"):
        try:
            x_tabs_data = app.get_x_tabs_data()
        except Exception:
            x_tabs_data = []

    settings = {
        "format": "Python Matplotlib Grapher App (HYGrapher) Graph Project",
        "version": "1.2",
        "application": "HYGrapher",
        "application_version": version_str,
        "dimension": dimension,
        "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "edited_data": data_dict,
        "x_tabs": x_tabs_data,
        "y1_series_styles": dict(getattr(app, "y1_series_styles", None) or {}),
        "y2_series_styles": dict(getattr(app, "y2_series_styles", None) or {}),
    }

    for key, attr, kind, default in _spec_for(dimension):
        settings[key] = _get_field(app, attr, kind, default)

    if dimension == "3D" and hasattr(app, "z_listbox"):
        settings["z_columns"] = [
            item.text() for item in app.z_listbox.selectedItems()
        ]

    return settings


def save_project_file(app, file_path, version_str=None, dimension="2D"):
    """
    Save current app state to a JSON file.
    """
    settings = build_project_dict(app, version_str=version_str, dimension=dimension)
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=4)
    return settings


def _restore_x_tabs(app, x_tabs_data):
    if not hasattr(app, "x_tab_widgets") or not x_tabs_data:
        return

    # Ensure we have exactly as many X-tabs as were saved.
    while len(app.x_tab_widgets) < len(x_tabs_data) and hasattr(app, "add_x_tab"):
        app.add_x_tab()
    while len(app.x_tab_widgets) > len(x_tabs_data) and hasattr(app, "remove_x_tab"):
        app.remove_x_tab(app.x_tab_widgets[-1]["tab_widget"])

    for info, saved in zip(app.x_tab_widgets, x_tabs_data):
        x_combo = info.get("x_combo")
        if x_combo is not None:
            x_col = saved.get("x_axis", "")
            idx = x_combo.findText(x_col)
            if idx >= 0:
                x_combo.setCurrentIndex(idx)

        for key, listbox_key in (("y1_cols", "y1_listbox"), ("y2_cols", "y2_listbox")):
            listbox = info.get(listbox_key)
            if listbox is None:
                continue
            wanted = set(saved.get(key, []))
            for row in range(listbox.count()):
                item = listbox.item(row)
                item.setSelected(item.text() in wanted)


def load_project_file(app, file_path):
    """
    Load project JSON file and restore app data and settings.
    """
    with open(file_path, "r", encoding="utf-8") as f:
        settings = json.load(f)

    dimension = settings.get("dimension", "2D")

    edited_data = settings.get("edited_data")
    if edited_data and isinstance(edited_data, dict):
        cols = edited_data.get("columns", [])
        rows = edited_data.get("data", [])
        if cols and rows:
            import pandas as pd

            df = pd.DataFrame(rows, columns=cols)
            app.data_mgr.set_dataframe(df)
            app.df = df
            if hasattr(app, "populate_data_table"):
                app.populate_data_table()
            if hasattr(app, "update_plot_options"):
                app.update_plot_options()
            elif hasattr(app, "update_combos"):
                app.update_combos()

    for key, attr, kind, _default in _spec_for(dimension):
        if key in settings:
            _set_field(app, attr, kind, settings[key])

    if hasattr(app, "y1_series_styles") and isinstance(
        settings.get("y1_series_styles"), dict
    ):
        app.y1_series_styles = dict(settings["y1_series_styles"])
    if hasattr(app, "y2_series_styles") and isinstance(
        settings.get("y2_series_styles"), dict
    ):
        app.y2_series_styles = dict(settings["y2_series_styles"])

    _restore_x_tabs(app, settings.get("x_tabs", []))

    if dimension == "3D" and hasattr(app, "z_listbox"):
        wanted = set(settings.get("z_columns", []))
        for row in range(app.z_listbox.count()):
            item = app.z_listbox.item(row)
            item.setSelected(item.text() in wanted)

    if hasattr(app, "update_style_editor_targets"):
        try:
            app.update_style_editor_targets()
        except Exception:
            pass

    return settings
