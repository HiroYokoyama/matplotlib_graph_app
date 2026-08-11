import re

orig = open('scratch_orig_main_3d.txt', encoding='utf-8').read()
new  = open('hygrapher/hygrapher/main_3d.py', encoding='utf-8').read()

orig_methods = set(re.findall(r'^    def (\w+)', orig, re.MULTILINE))
new_methods  = set(re.findall(r'^    def (\w+)', new,  re.MULTILINE))

print("=== METHODS IN ORIGINAL NOT IN NEW ===")
for m in sorted(orig_methods - new_methods):
    print(' MISSING:', m)

print()
print("=== METHODS IN NEW NOT IN ORIGINAL (additions) ===")
for m in sorted(new_methods - orig_methods):
    print(' NEW:', m)

print()
orig_vars = set(re.findall(r'self\.(\w+_var)\s*=\s*tk\.\w+Var', orig))
new_vars  = set(re.findall(r'self\.(\w+_var)\s*=\s*tk\.\w+Var', new))
print("=== TK VARS IN ORIGINAL NOT IN NEW ===")
for v in sorted(orig_vars - new_vars):
    print(' MISSING:', v)

print()
print("=== TK VARS IN NEW NOT IN ORIGINAL (additions) ===")
for v in sorted(new_vars - orig_vars):
    print(' NEW:', v)

print()
orig_plots = set(re.findall(r'(?:elif|if)\s+plot_type\s*==\s*["\'](\w+)["\']', orig))
new_plots  = set(re.findall(r'(?:elif|if)\s+plot_type\s*==\s*["\'](\w+)["\']', new))
print("=== PLOT DISPATCH IN ORIGINAL NOT IN NEW ===")
for p in sorted(orig_plots - new_plots):
    print(' MISSING dispatch:', p)

print()
# Check specific features present/absent in new
checks = [
    ('scrollable_canvas update',    'scrollable_canvas.update_idletasks'),
    ('get_font_list',               'get_font_list'),
    ('apply_data_filter method',    'def apply_data_filter'),
    ('apply_major_ticker',          'apply_major_ticker'),
    ('apply_minor_ticker',          'apply_minor_ticker'),
    ('tick font family applied',    'set_fontfamily'),
]
print("=== FEATURE CHECKS IN NEW MAIN_3D.PY ===")
for name, token in checks:
    status = 'OK     ' if token in new else 'MISSING'
    print(f"  {status}: {name}")
