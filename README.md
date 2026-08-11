# HYGrapher (2D & 3D)

[![HYGrapher CI](https://github.com/HiroYokoyama/matplotlib_graph_app/actions/workflows/ci.yml/badge.svg)](https://github.com/HiroYokoyama/matplotlib_graph_app/actions/workflows/ci.yml)
[![PyPI version](https://badge.fury.io/py/hygrapher.svg)](https://badge.fury.io/py/hygrapher)
[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)

A powerful, cross-platform GUI application for plotting and analyzing CSV/Excel data using Matplotlib, built with Python & Tkinter. Available in both **2D** and **3D** modes.

---

## ✨ Features

- **Cross-Platform Compatibility**: Fully supported on **macOS**, **Linux** (X11/Wayland with native `<Button-4>` / `<Button-5>` mouse wheel scrolling), and **Windows**.
- **Multi X-Axis Tabs**: Configure multiple dataset tabs on a single plot, each with its own X-column and Y1/Y2 series, overlaid seamlessly on one shared graph.
- **Non-Destructive Data Filtering**: Filter data ranges dynamically without corrupting or altering your raw loaded dataset.
- **Integrated Data Editor**: Embedded spreadsheet viewer (`tksheet`) to inspect, edit, and update data live.
- **2D & 3D Visualization Modes**:
  - **2D**: Line, Scatter, Bar, Step, Stem, Area, Pie, Box, Violin, Heatmap, Contour, and Polar plots.
  - **3D**: Surface, Wireframe, Scatter 3D, Line 3D, and Contour 3D plots.
- **Rich Style & Graph Controls**:
  - Independent per-series style controls (Color, Line Style, Line Width, Marker Style, Marker Size, Alpha/Opacity).
  - Minor Ticks (show/hide & interval control for X, Y1, Y2).
  - Customizable colormaps (Viridis, Plasma, Inferno, Coolwarm, Jet, etc.).
  - Export DPI settings (72, 100, 150, 300, 600 DPI).
  - Dedicated Legend font size & location controls.
  - Data smoothing (Moving Average), Error Bars, and Data Point Value Annotations.
- **Project Serialization**: Save and restore complete workspace states including all dataset modifications and multi-X-tab configurations using `.pmggrp` JSON project files.
- **Drag & Drop**: Drag CSV, Excel (`.xlsx`/`.xls`), or project (`.pmggrp`) files directly into the window.

---

## 💾 Installation

### Install from PyPI
```bash
pip install hygrapher
```

### Install with Developer & Testing Dependencies
```bash
git clone https://github.com/HiroYokoyama/matplotlib_graph_app.git
cd matplotlib_graph_app/hygrapher
pip install -e .[dev]
```

---

## 🚀 Usage

### Command Line Interface

Launch 2D mode:
```bash
hygrapher
```

Launch 2D mode with a dataset or project file:
```bash
hygrapher data.csv
hygrapher my_project.pmggrp
```

Launch 3D mode:
```bash
hygrapher-3d
hygrapher-3d data.csv
```

### Python Module Entrypoint
```bash
python -m hygrapher
```

---

## 🏗️ Project Architecture

```
matplotlib_graph_app/
├── .github/
│   └── workflows/
│       └── ci.yml             # GitHub Actions CI matrix (Ubuntu, macOS, Windows)
├── hygrapher/
│   ├── pyproject.toml         # Packaging metadata & entrypoints
│   ├── hygrapher/
│   │   ├── __init__.py        # Package exports
│   │   ├── __main__.py        # CLI entrypoint
│   │   ├── main.py            # 2D Application main window & Multi-X logic
│   │   ├── main_3d.py         # 3D Application main window & camera view controls
│   │   ├── data_manager.py    # Non-destructive data storage & range filter engine
│   │   ├── project_io.py      # .pmggrp JSON project serialization & deserialization
│   │   └── utils.py           # Cross-platform scroll event bindings & ticker math
│   └── tests/
│       ├── conftest.py        # Pytest configuration (Agg backend)
│       ├── test_data_manager.py
│       ├── test_project_io.py
│       ├── test_utils.py
│       ├── test_unit_comprehensive.py
│       └── test_app_headless.py
└── README.md
```

---

## 🧪 Testing & CI/CD

### Run Unit Tests Locally
```bash
cd hygrapher
pytest -v --cov=hygrapher
```

### Build Distribution Package
```bash
python -m build hygrapher
```
This produces `.whl` wheels and `.tar.gz` source archives inside `hygrapher/dist/`.

---

## 📄 License

Apache License 2.0. Created by **Hiromichi Yokoyama**.
