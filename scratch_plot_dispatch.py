import re

content = open('hygrapher/hygrapher/main.py', encoding='utf-8').read()

# Find all plot_type dispatch cases (elif/if plot_type == "...")
matches = re.findall(r'(?:elif|if)\s+plot_type\s*==\s*["\'](\w+)["\']', content)
print('plot_type dispatch cases in new main.py:')
for m in sorted(set(matches)):
    print(' ', m)
print()

# Also check what values are in the plot_type combobox
combo_vals = re.findall(r'values=\[([^\]]+)\]', content)
for v in combo_vals:
    if 'line' in v or 'scatter' in v or 'bar' in v:
        print('Combobox values:', v[:200])
        print()

# Find the plot_type combobox setup
idx = content.find('plot_type')
snippet = content[max(0,idx-100):idx+500]
print('Context around plot_type variable:', snippet[:600])
