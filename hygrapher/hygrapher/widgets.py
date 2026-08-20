# -*- coding: utf-8 -*-
"""
hygrapher.widgets

Small reusable Qt building blocks shared by the 2D and 3D windows.
"""

from PyQt6.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)


def build_series_picker(title, visible_rows=8):
    """
    A multi-select column picker: Select All / Clear buttons, a list tall
    enough to show `visible_rows` entries, and a live selection counter.

    Returns (container_widget, list_widget).
    """
    box = QGroupBox(title) if title else QWidget()
    layout = QVBoxLayout(box)
    layout.setContentsMargins(6, 6, 6, 6)
    layout.setSpacing(4)

    btn_row = QHBoxLayout()
    btn_row.setSpacing(4)
    select_all_btn = QPushButton("Select All")
    clear_btn = QPushButton("Clear")
    for btn in (select_all_btn, clear_btn):
        btn.setToolTip("Applies to the list below")
        btn_row.addWidget(btn)
    btn_row.addStretch(1)
    count_label = QLabel("0 selected")
    btn_row.addWidget(count_label)
    layout.addLayout(btn_row)

    list_widget = QListWidget()
    list_widget.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
    list_widget.setMinimumHeight(
        (list_widget.fontMetrics().height() + 8) * visible_rows
    )
    list_widget.setAlternatingRowColors(True)
    layout.addWidget(list_widget, stretch=1)

    select_all_btn.clicked.connect(list_widget.selectAll)
    clear_btn.clicked.connect(list_widget.clearSelection)

    def _update_count():
        count_label.setText(f"{len(list_widget.selectedItems())} selected")

    list_widget.itemSelectionChanged.connect(_update_count)
    # model changes (clear/addItem) do not emit itemSelectionChanged
    list_widget.model().rowsRemoved.connect(_update_count)
    list_widget.model().rowsInserted.connect(_update_count)

    return box, list_widget


def wrap_in_scroll_area(widget):
    """
    Put `widget` in a vertically scrolling, frameless viewport.

    The trailing stretch keeps the content packed at the top: without it Qt
    hands the viewport's spare height to the rows, spreading a handful of
    checkboxes over the whole panel.
    """
    container = QWidget()
    container_layout = QVBoxLayout(container)
    container_layout.setContentsMargins(0, 0, 0, 0)
    container_layout.addWidget(widget)
    container_layout.addStretch(1)

    area = QScrollArea()
    area.setWidgetResizable(True)
    area.setFrameShape(QFrame.Shape.NoFrame)
    area.setWidget(container)
    return area
