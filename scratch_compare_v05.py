# -*- coding: utf-8 -*-
import re

v05 = open('scratch_v05_main.txt', encoding='utf-8').read()
curr = open('hygrapher/hygrapher/main.py', encoding='utf-8').read()

print("=== DEGRADATION CHECK: v0.5 vs CURRENT PyQt6 ===")

# 1. Tk variables in v0.5 vs options in Current
v05_vars = set(re.findall(r'self\.(\w+_var)\s*=\s*tk\.\w+Var', v05))

print(f"\n1. v0.5 Variables Audit (Total: {len(v05_vars)})")

# Map of Tk variable names to their corresponding PyQt6 controls/attributes
var_mapping = {
    "plot_type_var": "plot_type_combo",
    "x_axis_var": "x_combo",
    "title_var": "title_input",
    "xlabel_var": "xlabel_input",
    "ylabel_var": "ylabel_input",
    "ylabel2_var": "ylabel2_input",
    "grid_var": "grid_check",
    "marker_var": "style_marker_combo",
    "font_family_var": "font_family_combo",
    "title_fontsize_var": "title_fontsize_spin",
    "xlabel_fontsize_var": "xlabel_fontsize_spin",
    "ylabel_fontsize_var": "ylabel_fontsize_spin",
    "ylabel2_fontsize_var": "ylabel2_fontsize_spin",
    "tick_fontsize_var": "tick_fontsize_spin",
    "tick2_fontsize_var": "tick2_fontsize_spin",
    "legend_fontsize_var": "legend_fontsize_spin",
    "fig_width_var": "figsize",
    "fig_height_var": "figsize",
    "xlim_min_var": "xlim_min_input",
    "xlim_max_var": "xlim_max_input",
    "ylim_min_var": "ylim_min_input",
    "ylim_max_var": "ylim_max_input",
    "ylim2_min_var": "ylim2_min_input",
    "ylim2_max_var": "ylim2_max_input",
    "xtick_show_var": "tick_params",
    "xtick_label_show_var": "tick_params",
    "xtick_direction_var": "tick_params",
    "ytick_show_var": "tick_params",
    "ytick_label_show_var": "tick_params",
    "ytick_direction_var": "tick_params",
    "ytick2_show_var": "tick_params",
    "ytick2_label_show_var": "tick_params",
    "ytick2_direction_var": "tick_params",
    "xtick_minor_show_var": "apply_minor_ticker",
    "ytick_minor_show_var": "apply_minor_ticker",
    "ytick2_minor_show_var": "apply_minor_ticker",
    "xtick_minor_interval_var": "apply_minor_ticker",
    "ytick_minor_interval_var": "apply_minor_ticker",
    "ytick2_minor_interval_var": "apply_minor_ticker",
    "xaxis_plain_format_var": "xaxis_plain_check",
    "yaxis1_plain_format_var": "yaxis1_plain_check",
    "yaxis2_plain_format_var": "yaxis2_plain_check",
    "xtick_major_interval_var": "apply_major_ticker",
    "ytick_major_interval_var": "apply_major_ticker",
    "ytick2_major_interval_var": "apply_major_ticker",
    "spine_top_var": "spine_top_check",
    "spine_bottom_var": "spine_bottom_check",
    "spine_left_var": "spine_left_check",
    "spine_right_var": "spine_right_check",
    "face_color_var": "QColorDialog",
    "fig_color_var": "QColorDialog",
    "legend_show_var": "legend_show_check",
    "legend_loc_var": "legend_loc_combo",
    "x_log_scale_var": "x_log_check",
    "y1_log_scale_var": "y1_log_check",
    "y2_log_scale_var": "y2_log_check",
    "x_invert_var": "set_axis_limits",
    "y1_invert_var": "y1_invert_check",
    "y2_invert_var": "y2_invert_check",
    "enable_smoothing_var": "enable_smoothing_check",
    "smoothing_window_var": "smoothing_window_spin",
    "enable_errorbar_var": "enable_errorbar_check",
    "errorbar_column_var": "errorbar_column_combo",
    "enable_annotation_var": "enable_annotation_check",
    "data_filter_enabled_var": "data_filter_check",
    "filter_min_var": "filter_min_input",
    "filter_max_var": "filter_max_input",
    "filter_column_var": "filter_column_combo",
    "grid_alpha_var": "grid_alpha",
    "grid_linestyle_var": "grid_linestyle",
    "grid_linewidth_var": "grid_linewidth",
    "subplot_mode_var": "subplot_mode_check",
    "rotate_labels_var": "rotate_labels_check",
    "rotation_angle_var": "rotation_angle_spin",
    "export_dpi_var": "export_dpi_spin",
    "colormap_var": "colormap_combo",
    "tight_layout_pad_var": "tight_layout",
}

missing_features = []
for var in v05_vars:
    mapped = var_mapping.get(var)
    if mapped:
        if mapped in curr:
            print(f"  OK     : {var} -> {mapped}")
        else:
            print(f"  MISSING: {var} -> {mapped}")
            missing_features.append((var, mapped))
    else:
        # Check if var name base is in curr
        base = var.replace("_var", "")
        if base in curr:
            print(f"  OK     : {var} -> {base}")
        else:
            print(f"  UNKNOWN: {var}")

print()
# 2. Check plot dispatches
v05_plots = set(re.findall(r'(?:elif|if)\s+plot_type\s*==\s*["\'](\w+)["\']', v05))
curr_plots = set(re.findall(r'(?:elif|if)\s+plot_type\s*==\s*["\'](\w+)["\']', curr))

print("2. Plot Dispatch Check:")
print("  v0.5 Plot Types:", sorted(v05_plots))
print("  Current Plot Types:", sorted(curr_plots))
missing_plots = v05_plots - curr_plots
if missing_plots:
    print("  MISSING PLOT TYPES:", sorted(missing_plots))
else:
    print("  OK     : All v0.5 plot types supported!")

print()
# 3. Check methods
v05_methods = set(re.findall(r'^    def (\w+)', v05, re.MULTILINE))
curr_methods = set(re.findall(r'^    def (\w+)', curr, re.MULTILINE))

print("3. Public/Helper Methods Audit:")
important_methods = [
    "load_data", "plot_graph", "save_settings", "overwrite_save",
    "load_settings", "load_project_file", "export_graph", "export_filtered_data",
    "clear_all", "reset_settings", "add_x_tab", "remove_x_tab", "get_x_tabs_data",
    "on_style_editor_change", "on_combined_series_select", "set_axis_limits"
]

for m in important_methods:
    status = "OK     " if m in curr_methods or m in curr else "MISSING"
    print(f"  {status}: {m}")
