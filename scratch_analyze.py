import re

content = open('scratch_orig_main.txt', encoding='utf-8').read()

# Find all def statements
defs = re.findall(r'^    def (\w+)', content, re.MULTILINE)
print('=== METHODS ===')
for d in defs:
    print(' ', d)
print()

# Find plot types from elif plot_type == "..."
plot_types = re.findall(r'plot_type\s*==\s*["\'](\w+)["\']', content)
print('=== PLOT TYPES (unique) ===')
for p in sorted(set(plot_types)):
    print(' ', p)
print()

# Find tk.Variables created
vars_created = re.findall(r'self\.(\w+_var)\s*=\s*tk\.\w+Var', content)
print('=== TK VARIABLES ===')
for v in sorted(set(vars_created)):
    print(' ', v)
print()

# Find key features
features = [
    ('subplot_mode', 'subplot_mode_var'),
    ('smoothing', 'enable_smoothing_var'),
    ('errorbar', 'enable_errorbar_var'),
    ('annotation', 'enable_annotation_var'),
    ('dnd', 'DND_AVAILABLE'),
    ('filter', 'filter'),
    ('get_data_from_sheet', 'get_data_from_sheet'),
    ('export_filtered_data', 'export_filtered_data'),
    ('open_in_3d_mode', 'open_in_3d_mode'),
]
print('=== KEY FEATURES PRESENT ===')
for name, token in features:
    found = token in content
    print(f'  {name}: {found}')
