# -*- coding: utf-8 -*-
"""
hygrapher.utils

Cross-platform helper utilities: CLI argument handling, matplotlib
formatting, and font detection.
"""

import os

import matplotlib.font_manager as fm
import matplotlib.ticker as ticker


def resolve_cli_file(argv):
    """
    Return the first existing file path found among command-line arguments
    (e.g. a file dragged onto the app icon, "Open with...", or
    `hygrapher path/to/data.csv`), skipping flags like ``-x``. Returns
    ``None`` if there isn't one, so `main()` can fall back to an empty window.
    """
    for arg in argv:
        if arg.startswith("-"):
            continue
        if os.path.isfile(arg):
            return arg
    return None


def get_font_list():
    """
    Get system fonts prioritized by cross-platform common fonts.
    """
    try:
        font_list = sorted(list(set(fm.fontManager.get_font_names())))
        common_fonts = [
            "sans-serif",
            "serif",
            "monospace",
            "Arial",
            "Helvetica",
            "Times New Roman",
            "Courier New",
            "DejaVu Sans",
            "Liberation Sans",
            "Yu Gothic",
            "Meiryo",
            "MS Gothic",
        ]
        for f in reversed(common_fonts):
            if f in font_list:
                font_list.remove(f)
                font_list.insert(0, f)
        return font_list
    except Exception as e:
        print(f"Failed to load system fonts: {e}")
        return ["sans-serif", "serif", "monospace", "Arial", "Helvetica"]


def apply_major_ticker(axis, interval_str, is_log_scale):
    """
    Apply MultipleLocator for major ticks if interval_str is a valid positive float
    and the axis is in linear scale.
    """
    if is_log_scale:
        return

    try:
        interval = float(interval_str)
        if interval > 0:
            axis.set_major_locator(ticker.MultipleLocator(interval))
    except (ValueError, TypeError):
        pass


def apply_minor_ticker(axis, show_minor, interval_str, is_log_scale):
    """
    Apply minor ticker settings to an axis.
    """
    if is_log_scale:
        if show_minor:
            axis.set_minor_locator(ticker.LogLocator(subs="auto"))
        return

    if show_minor:
        try:
            interval = float(interval_str)
            if interval > 0:
                axis.set_minor_locator(ticker.MultipleLocator(interval))
            else:
                axis.set_minor_locator(ticker.AutoMinorLocator())
        except (ValueError, TypeError):
            axis.set_minor_locator(ticker.AutoMinorLocator())
    else:
        axis.set_minor_locator(ticker.NullLocator())
