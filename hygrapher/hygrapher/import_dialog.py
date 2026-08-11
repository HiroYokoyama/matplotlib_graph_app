# -*- coding: utf-8 -*-
"""
hygrapher.import_dialog

A small preview dialog shown when importing a data file, letting the user
pick which row is the real header row (or declare the file has none) for
files that have an extra title/metadata line above the real column headers.
Shared by the 2D and 3D windows.
"""

import os

from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

_HEADER_ROW_COLOR = QColor("#2b5c8f")
_HEADER_ROW_TEXT_COLOR = QColor("#ffffff")


class ImportPreviewDialog(QDialog):
    """
    Shows the first few raw rows of a file and lets the user choose which
    row is the header row (0-based), or declare the file has no header at
    all. After `exec()` returns `QDialog.DialogCode.Accepted`, read
    `header_row` — an int, or `None` for "no header / auto-name columns".
    """

    def __init__(self, file_path, preview_rows, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Import Preview — {os.path.basename(file_path)}")
        self.resize(700, 420)
        self.header_row = 0
        self._preview_rows = preview_rows

        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel(
                "Pick which row holds the column names. Any rows above it "
                "(e.g. a title line) are skipped when the data is loaded."
            )
        )

        self.table = QTableWidget()
        self._populate_table(preview_rows)
        layout.addWidget(self.table)

        control_layout = QHBoxLayout()
        control_layout.addWidget(QLabel("Header row:"))
        self.header_spin = QSpinBox()
        self.header_spin.setRange(0, max(0, len(preview_rows) - 1))
        self.header_spin.setValue(0)
        self.header_spin.valueChanged.connect(self._on_header_row_changed)
        control_layout.addWidget(self.header_spin)

        self.no_header_check = QCheckBox("No header row (auto-name columns)")
        self.no_header_check.toggled.connect(self._on_no_header_toggled)
        control_layout.addWidget(self.no_header_check)
        control_layout.addStretch()
        layout.addLayout(control_layout)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._highlight_header_row(0)

    def _populate_table(self, preview_rows):
        if not preview_rows:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        n_cols = max(len(row) for row in preview_rows)
        self.table.setRowCount(len(preview_rows))
        self.table.setColumnCount(n_cols)
        self.table.setHorizontalHeaderLabels([f"Col {i + 1}" for i in range(n_cols)])
        self.table.setVerticalHeaderLabels([str(i) for i in range(len(preview_rows))])

        for r, row in enumerate(preview_rows):
            for c in range(n_cols):
                value = str(row[c]) if c < len(row) else ""
                self.table.setItem(r, c, QTableWidgetItem(value))

    def _on_header_row_changed(self, value):
        self.header_row = value
        self._highlight_header_row(value)

    def _on_no_header_toggled(self, checked):
        self.header_spin.setEnabled(not checked)
        self.header_row = None if checked else self.header_spin.value()
        self._highlight_header_row(None if checked else self.header_spin.value())

    def _highlight_header_row(self, header_row):
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                if item is None:
                    continue
                if r == header_row:
                    item.setBackground(_HEADER_ROW_COLOR)
                    item.setForeground(_HEADER_ROW_TEXT_COLOR)
                else:
                    item.setBackground(QColor("transparent"))
                    item.setForeground(QColor("black"))
