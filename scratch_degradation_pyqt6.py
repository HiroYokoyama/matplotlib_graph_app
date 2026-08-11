import re

orig_main = open('scratch_orig_main.txt', encoding='utf-8').read()
new_main  = open('hygrapher/hygrapher/main.py', encoding='utf-8').read()

print("=== 2D PLOT DISPATCH AUDIT ===")
orig_plots = set(re.findall(r'(?:elif|if)\s+plot_type\s*==\s*["\'](\w+)["\']', orig_main))
new_plots  = set(re.findall(r'(?:elif|if)\s+plot_type\s*==\s*["\'](\w+)["\']', new_main))
print("Original 2D plot types:", sorted(orig_plots))
print("New PyQt6 2D plot types:", sorted(new_plots))
missing_plots = orig_plots - new_plots
if missing_plots:
    print("MISSING 2D PLOTS:", sorted(missing_plots))
else:
    print("All 2D plot types present!")

print()
print("=== 2D FEATURE AUDIT ===")
checks_2d = [
    ("Rotate Labels", "rotate_labels_check"),
    ("Subplot Mode", "subplot_mode_check"),
    ("Line Smoothing", "enable_smoothing_check"),
    ("Error Bars", "enable_errorbar_check"),
    ("Annotations", "enable_annotation_check"),
    ("Non-Destructive Filter", "data_filter_check"),
    ("Custom Colormap", "colormap_combo"),
    ("Export DPI", "export_dpi_spin"),
    ("Axis Limits (X/Y1/Y2)", "xlim_min_input"),
    ("Multi X-Tabs", "add_x_tab"),
    ("Series Styles Editor", "y1_series_styles"),
    ("Project Save/Load (.pmggrp)", "save_project_file"),
    ("Drag and Drop (.pmggrp / .csv)", "dropEvent"),
]

for feature, keyword in checks_2d:
    status = "OK     " if keyword in new_main else "MISSING"
    print(f"  {status}: {feature}")

print()
print("=== 3D FEATURE AUDIT ===")
orig_3d = open('scratch_orig_main_3d.txt', encoding='utf-8').read()
new_3d  = open('hygrapher/hygrapher/main_3d.py', encoding='utf-8').read()
orig_3d_plots = set(re.findall(r'(?:elif|if)\s+plot_type\s*==\s*["\'](\w+)["\']', orig_3d))
new_3d_plots  = set(re.findall(r'(?:elif|if)\s+plot_type\s*==\s*["\'](\w+)["\']', new_3d))
print("Original 3D plot types:", sorted(orig_3d_plots))
print("New PyQt6 3D plot types:", sorted(new_3d_plots))
missing_3d_plots = orig_3d_plots - new_3d_plots
if missing_3d_plots:
    print("MISSING 3D PLOTS:", sorted(missing_3d_plots))
else:
    print("All 3D plot types present!")

checks_3d = [
    ("View Elevation/Azimuth", "elev_spin"),
    ("Mesh Resolution", "resolution_spin"),
    ("Colormap", "colormap_combo"),
    ("Drag and Drop", "dropEvent"),
    ("3D Project Save/Load", "save_project_file"),
]

for feature, keyword in checks_3d:
    status = "OK     " if keyword in new_3d else "MISSING"
    print(f"  {status}: {feature}")
