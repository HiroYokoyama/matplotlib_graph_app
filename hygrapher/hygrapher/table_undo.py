# -*- coding: utf-8 -*-
"""
hygrapher.table_undo

Cell-level undo/redo for the data-sheet QTableWidget, shared by the 2D and
3D windows via QUndoStack. History is kept as a flat sequence of single-cell
edits rather than full-table snapshots, since edits are cheap to replay.
"""

from PyQt6.QtGui import QUndoCommand
from PyQt6.QtWidgets import QTableWidgetItem


class CellEditCommand(QUndoCommand):
    def __init__(self, table, snapshot, row, col, old_text, new_text):
        super().__init__(f"Edit cell ({row + 1}, {col + 1})")
        self.table = table
        self.snapshot = snapshot
        self.row = row
        self.col = col
        self.old_text = old_text
        self.new_text = new_text

    def _apply(self, text):
        self.table.blockSignals(True)
        item = self.table.item(self.row, self.col)
        if item is None:
            item = QTableWidgetItem()
            self.table.setItem(self.row, self.col, item)
        item.setText(text)
        self.table.blockSignals(False)
        self.snapshot[(self.row, self.col)] = text

    def undo(self):
        self._apply(self.old_text)

    def redo(self):
        self._apply(self.new_text)


def snapshot_table(table):
    """Capture every cell's current text, keyed by (row, col)."""
    return {
        (r, c): (table.item(r, c).text() if table.item(r, c) else "")
        for r in range(table.rowCount())
        for c in range(table.columnCount())
    }


def handle_table_item_changed(item, table, snapshot, undo_stack):
    """Call from a QTableWidget.itemChanged handler: pushes a
    CellEditCommand only if the cell's text actually changed relative to
    the snapshot, so bulk repopulation (which also fires itemChanged for
    every cell) doesn't pollute undo history — callers must block the
    table's signals while repopulating and take a fresh snapshot after.
    """
    row, col = item.row(), item.column()
    key = (row, col)
    old_text = snapshot.get(key, "")
    new_text = item.text()
    if old_text == new_text:
        return
    undo_stack.push(CellEditCommand(table, snapshot, row, col, old_text, new_text))
