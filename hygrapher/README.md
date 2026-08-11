# HYGrapher (2D & 3D)

[![HYGrapher CI/CD](https://github.com/HiroYokoyama/matplotlib_graph_app/actions/workflows/ci.yml/badge.svg)](https://github.com/HiroYokoyama/matplotlib_graph_app/actions/workflows/ci.yml)
[![PyPI version](https://badge.fury.io/py/hygrapher.svg)](https://badge.fury.io/py/hygrapher)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

A high-performance, cross-platform desktop application for plotting, visual editing, and analyzing CSV/Excel datasets using Matplotlib and **PyQt6**. Available in both **2D** and **3D** visualization modes.

---

## ✨ Key Features

- **PyQt6 Architecture**: Modern OS-native look & feel, dark/light theme support, high performance, and 100% crash-free cross-platform execution on macOS, Linux, and Windows.
- **Multi X-Axis Tabs**: Configure multiple dataset tabs on a single plot, each with its own X-column and Y1/Y2 series, overlaid seamlessly on one shared graph.
- **Non-Destructive Data Filtering**: Filter data ranges dynamically without corrupting or altering your raw loaded dataset.
- **Integrated Data Sheet**: Embedded `QTableWidget` spreadsheet viewer to inspect, edit, and update dataset rows and columns in real-time.
- **2D & 3D Visualization Modes**:
  - **2D**: Line, Scatter, Bar, Step, Stem, Area, Pie, Box, Violin, Heatmap, Contour, and Polar plots.
  - **3D**: Surface, Wireframe, Scatter 3D, Line 3D, and Contour 3D plots.
- **Rich Style & Graph Controls**:
  - Independent per-series style controls (Color, Line Style, Line Width, Marker Style, Marker Size, Alpha/Opacity, Legend Labels).
  - Axis formatting (rotate X-tick labels, plain numeric formatting without scientific notation, custom major tick intervals).
  - Customizable colormaps (`viridis`, `plasma`, `inferno`, `coolwarm`, `jet`, `turbo`, etc.).
  - High-DPI Image Export (72 to 1200 DPI PNG, PDF, SVG).
  - Legend font size, location, and spine toggle controls.
  - Moving Average Line Smoothing, Error Bars, and Data Point Value Annotations.
- **Project Serialization**: Save and restore complete workspace states including all dataset modifications and multi-X-tab configurations using `.pmggrp` JSON project files.
- **Native Drag & Drop**: Drag CSV, TSV, Excel (`.xlsx`/`.xls`), JSON, or project (`.pmggrp`) files directly into the window to open.

---

## 💾 Installation

### Install from PyPI
```bash
pip install hygrapher
```

### Install for Development
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

### Python Entrypoint
```bash
python -m hygrapher
python -m hygrapher.main_3d
```

---

## 🏗️ Architecture & Modules

```
matplotlib_graph_app/
├── .github/
│   └── workflows/
│       └── ci.yml             # Automated CI matrix testing & PyPI tag publishing
├── hygrapher/
│   ├── pyproject.toml         # Packaging metadata & entrypoints
│   ├── hygrapher/
│   │   ├── __init__.py        # Package version & exports
│   │   ├── __main__.py        # CLI entrypoint
│   │   ├── main.py            # 2D Application main window (PyQt6 QMainWindow)
│   │   ├── main_3d.py         # 3D Application main window (mplot3d & QMainWindow)
│   │   ├── data_manager.py    # Non-destructive data storage & range filter engine
│   │   ├── project_io.py      # .pmggrp JSON project serialization & deserialization
│   │   └── utils.py           # Cross-platform ticker math & font manager helpers
│   └── tests/
│       ├── conftest.py        # Pytest configuration & offscreen Qt setup
│       ├── test_app_headless.py # 38 fine-grained PyQt6 application tests
│       ├── test_data_manager.py
│       ├── test_project_io.py
│       ├── test_unit_comprehensive.py
│       └── test_utils.py
└── README.md
```

---

## 🧪 Testing & CI/CD

### Run Unit Tests Locally
```bash
cd hygrapher
pytest -v --cov=hygrapher
```

### Build & Release Workflow
The repository includes automated GitHub Actions CI/CD (`.github/workflows/ci.yml`):
- Runs unit tests and coverage across Linux, macOS, and Windows matrix.
- **Automated PyPI Publishing**: Pushing any tag matching `v*` (e.g. `v0.7.0`) automatically builds wheels and publishes the release package to PyPI.

```bash
# Release a new version:
git tag v0.7.0
git push origin v0.7.0
```

---

## 📄 License

GNU General Public License v3.0 (GPL-3.0). Created by **Hiromichi Yokoyama**.
