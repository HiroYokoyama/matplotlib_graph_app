# -*- coding: utf-8 -*-
"""
hygrapher.utils

Cross-platform helper utilities for Tkinter GUI events, matplotlib formatting, and font detection.
"""

import sys
import matplotlib.font_manager as fm
import matplotlib.ticker as ticker


def bind_scroll_events(widget, callback_func):
    """
    Bind mouse wheel scrolling to a widget across Windows, macOS, and Linux.

    Windows / macOS: <MouseWheel> with event.delta
    Linux (X11): <Button-4> (scroll up) and <Button-5> (scroll down)
    """
    def _on_mousewheel(event):
        if hasattr(event, 'num') and event.num == 4:
            # Linux scroll up
            callback_func(scroll_units=-1, event=event)
        elif hasattr(event, 'num') and event.num == 5:
            # Linux scroll down
            callback_func(scroll_units=1, event=event)
        elif hasattr(event, 'delta') and event.delta != 0:
            # Windows & macOS
            # On macOS event.delta might be small or negative depending on OS settings
            units = int(-1 * (event.delta / 120))
            if units == 0:
                units = -1 if event.delta > 0 else 1
            callback_func(scroll_units=units, event=event)

    widget.bind("<MouseWheel>", _on_mousewheel)
    widget.bind("<Button-4>", _on_mousewheel)
    widget.bind("<Button-5>", _on_mousewheel)


def bind_mousewheel_recursive(widget, callback_func):
    """
    Recursively bind scroll events to a widget and all of its children.
    """
    bind_scroll_events(widget, callback_func)
    for child in widget.winfo_children():
        bind_mousewheel_recursive(child, callback_func)


def get_font_list():
    """
    Get system fonts prioritized by cross-platform common fonts.
    """
    try:
        font_list = sorted(list(set(fm.fontManager.get_font_names())))
        common_fonts = [
            'sans-serif', 'serif', 'monospace',
            'Arial', 'Helvetica', 'Times New Roman', 'Courier New',
            'DejaVu Sans', 'Liberation Sans', 'Yu Gothic', 'Meiryo', 'MS Gothic'
        ]
        for f in reversed(common_fonts):
            if f in font_list:
                font_list.remove(f)
                font_list.insert(0, f)
        return font_list
    except Exception as e:
        print(f"Failed to load system fonts: {e}")
        return ['sans-serif', 'serif', 'monospace', 'Arial', 'Helvetica']


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
            axis.set_minor_locator(ticker.LogLocator(subs='auto'))
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
