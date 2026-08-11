new  = open('hygrapher/hygrapher/main.py', encoding='utf-8').read()

checks = [
    ('rotate_labels applied',       'rotate_labels_var.get()'),
    ('xaxis_plain_format applied',  'xaxis_plain_format_var'),
    ('yaxis1_plain_format applied', 'yaxis1_plain_format_var'),
    ('yaxis2_plain_format applied', 'yaxis2_plain_format_var'),
    ('tick font family applied',    'set_fontfamily'),
    ('spine right when no ax2',     'spine_right_var.get()'),
    ('ax2 log scale',               'y2_log_scale_var'),
    ('ax2 invert',                  'y2_invert_var'),
    ('errorbar in plot_series',     'enable_errorbar_var'),
    ('annotation in plot_series',   'enable_annotation_var'),
    ('stem plot dispatch',          'plot_type == "stem"'),
    ('subplot separate legend',     'subplot_mode_var'),
    ('scrollable_canvas update',    'scrollable_canvas.update_idletasks'),
    ('get_font_list',               'get_font_list'),
    ('apply_data_filter method',    'def apply_data_filter'),
    ('legend fontsize from tick_fontsize', 'tick_fontsize_var'),
]
print("=== FEATURE CHECKS IN NEW MAIN.PY ===")
for name, token in checks:
    status = 'OK     ' if token in new else 'MISSING'
    print(f"  {status}: {name}")
