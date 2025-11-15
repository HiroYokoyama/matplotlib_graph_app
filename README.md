# HYGrapher (2D & 3D)

**Author**: Hiromichi Yokoyama
**License**: Apache-2.0
**Repo**: `https://github.com/HiroYokoyama/matplotlib_graph_app`

-----

## Overview

**HYGrapher** is a powerful and flexible GUI application built with Python, Tkinter, and Matplotlib for interactively visualizing CSV and Excel data. It provides a user-friendly interface to load data, customize graph aesthetics extensively, and export high-quality plots for presentations or publications.

The application comes in two modes, both included in the package:

  * **`hygrapher` (2D Mode)**: For creating a wide variety of standard 2D plots, including line graphs, scatter plots, bar charts, heatmaps, and contour plots, with dual Y-axis support.
  * **`hygrapher-3d` (3D Mode)**: For generating 3D visualizations like surface plots, wireframes, and 3D scatter plots.

Both modes allow you to save your complete session (data + settings) as a `.pmggrp` project file, letting you resume your work later.

## Features

### General Features

  * **Data Input**: Load data directly from **CSV** or **Excel** files (`.csv`, `.xlsx`, `.xls`).
  * **Drag & Drop**: Supports drag-and-drop for data files and `.pmggrp` project files.
  * **Data Editor**: View and edit your data directly within the app using an integrated `tksheet` table.
  * **Project System**: Save and load your entire workspace (data and all settings) using the custom `.pmggrp` JSON-based project format.
  * **Graph Export**: Export your finalized graph to various formats, including **PNG**, **SVG**, **PDF**, **JPEG**, and more.
  * **Data Export**: Export the current (edited or filtered) data back to a CSV file.
  * **Full Customization**:
      * Set graph titles, axis labels, and font properties (family, size).
      * Control axis limits, tick intervals, and tick direction.
      * Toggle log scale or invert axes.
      * Customize grid, spines, and background colors (both axes and figure).
      * Add and position a legend.
  * **Mode Switching**: Easily switch between 2D and 3D modes from the "File" menu.

-----

### 2D Mode (`hygrapher`)

  * **Plot Types**: Supports a wide range of plots:
      * `line`
      * `scatter`
      * `bar`
      * `step`
      * `stem`
      * `area`
      * `pie`
      * `box`
      * `violin`
      * `heatmap`
      * `contour` (requires `scipy`)
      * `polar`
  * **Dual Y-Axes**: Plot different data series on a left (Y1) and right (Y2) axis.
  * **Per-Series Styling**: Individually customize the style (color, line style, marker, line width, alpha) for every data series on both Y1 and Y2 axes.
  * **Advanced Features**:
      * **Smoothing**: Apply a moving average to line plots.
      * **Error Bars**: Add error bars to your data from a selected data column.
      * **Annotations**: Automatically annotate data points on your graph.
      * **Data Filtering**: Filter the data used for plotting based on a column's value range.
      * **Subplot Mode**: Automatically split Y1 and Y2 data into two separate, vertically stacked subplots.

-----

### 3D Mode (`hygrapher-3d`)

  * **Plot Types**: Create 3D visualizations:
      * `surface` (requires `scipy`)
      * `wireframe` (requires `scipy`)
      * `scatter3d`
      * `line3d`
      * `contour3d` (requires `scipy`)
  * **View Control**: Interactively adjust the 3D **Elevation** and **Azimuth** angles.
  * **Surface/Wireframe Quality**: Control the **Mesh Resolution** for smoother 3D surfaces.
  * **Colormap Selection**: Choose from a comprehensive list of Matplotlib colormaps for surface and contour plots.
  * **Per-Series Styling**: Customize color, alpha, line width, and markers for different Z-axis data series.

## Requirements

  * Python 3
  * `pandas`
  * `matplotlib`
  * `numpy`
  * `tksheet`
  * `scipy`
  * `openpyxl`
  * `tkinterdnd2`

**Compatibility:**

**This package is currently for Windows only.** The application relies on OS-specific components that are not compatible with macOS or Linux at this time.

## Installation

You can install HYGrapher directly from PyPI:

```bash
pip install hygrapher
```

## Usage

### Running the Application

After installation, you can run the application from your command line:

  * **To run the 2D version:**

    ```bash
    hygrapher
    ```

  * **To run the 3D version:**

    ```bash
    hygrapher-3d
    ```

### Loading Data

You can load data in several ways:

1.  **Menu**: Click the **"Load Data"** button or go to `File > Load Data (CSV/Excel)...`.
2.  **Drag & Drop**: Drag a `.csv`, `.xlsx`, `.xls`, or `.pmggrp` file directly onto the application window.
3.  **Command Line**: Pass a file path as an argument when launching the script:
    ```bash
    hygrapher path/to/your/data.csv
    ```
    or
    ```bash
    hygrapher-3d path/to/your/project.pmggrp
    ```

### Basic Workflow

1.  Load your data.
2.  Go to the **"Basic Settings"** tab.
3.  Select your **Plot Type**.
4.  Select the column for the **X-Axis**.
5.  Select one or more columns for the **Y-Axis** (2D) or **Z-Axis** (3D).
6.  Click the **"Plot/Update Graph"** button.
7.  Use the other tabs (**Style**, **Font**, **Axis/Ticks**, etc.) to customize the graph.
8.  Click **"Plot/Update Graph"** again to apply changes.
9.  Save your work as a project (`File > Save Project`) or export the graph (`File > Export Graph...`).

## Known Issues

  * **Windows Only:** As stated in the requirements, the application is currently **not cross-platform**. It will not function correctly on macOS or Linux.
  * **Technical Reason:** The code relies on Windows-specific UI event handling. Specifically, the mouse wheel scrolling uses `event.delta`, which is not available on Linux (which uses `<Button-4>`/`<Button-5>`). Furthermore, the drag-and-drop file path handling is designed to strip Windows-specific `{}` characters, which will fail on other operating systems.

## License

This project is licensed under the **Apache-2.0 License**.
