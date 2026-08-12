# -*- coding: utf-8 -*-
"""
hygrapher.table_undo

Undo/redo for the data-sheet QTableWidget, shared by the 2D and 3D windows
via QUndoStack: cell edits, and row/column insert/delete. History is kept
as a flat sequence of individual commands rather than full-table snapshots,
since each command is cheap to replay.
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


class InsertRowCommand(QUndoCommand):
    def __init__(self, table, snapshot, row, num_cols, on_changed=None):
        super().__init__(f"Insert row {row + 1}")
        self.table = table
        self.snapshot = snapshot
        self.row = row
        self.num_cols = num_cols
        self.on_changed = on_changed

    def _resync(self):
        self.snapshot.clear()
        self.snapshot.update(snapshot_table(self.table))
        if self.on_changed:
            self.on_changed()

    def redo(self):
        self.table.blockSignals(True)
        self.table.insertRow(self.row)
        for c in range(self.num_cols):
            self.table.setItem(self.row, c, QTableWidgetItem(""))
        self.table.blockSignals(False)
        self._resync()

    def undo(self):
        self.table.blockSignals(True)
        self.table.removeRow(self.row)
        self.table.blockSignals(False)
        self._resync()


class DeleteRowCommand(QUndoCommand):
    def __init__(self, table, snapshot, row, row_data, on_changed=None):
        super().__init__(f"Delete row {row + 1}")
        self.table = table
        self.snapshot = snapshot
        self.row = row
        self.row_data = row_data
        self.on_changed = on_changed

    def _resync(self):
        self.snapshot.clear()
        self.snapshot.update(snapshot_table(self.table))
        if self.on_changed:
            self.on_changed()

    def redo(self):
        self.table.blockSignals(True)
        self.table.removeRow(self.row)
        self.table.blockSignals(False)
        self._resync()

    def undo(self):
        self.table.blockSignals(True)
        self.table.insertRow(self.row)
        for c, text in enumerate(self.row_data):
            self.table.setItem(self.row, c, QTableWidgetItem(text))
        self.table.blockSignals(False)
        self._resync()


class InsertColumnCommand(QUndoCommand):
    def __init__(self, table, snapshot, col, header_text, num_rows, on_changed=None):
        super().__init__(f"Insert column '{header_text}'")
        self.table = table
        self.snapshot = snapshot
        self.col = col
        self.header_text = header_text
        self.num_rows = num_rows
        self.on_changed = on_changed

    def _resync(self):
        self.snapshot.clear()
        self.snapshot.update(snapshot_table(self.table))
        if self.on_changed:
            self.on_changed()

    def redo(self):
        self.table.blockSignals(True)
        self.table.insertColumn(self.col)
        self.table.setHorizontalHeaderItem(
            self.col, QTableWidgetItem(self.header_text)
        )
        for r in range(self.num_rows):
            self.table.setItem(r, self.col, QTableWidgetItem(""))
        self.table.blockSignals(False)
        self._resync()

    def undo(self):
        self.table.blockSignals(True)
        self.table.removeColumn(self.col)
        self.table.blockSignals(False)
        self._resync()


class DeleteColumnCommand(QUndoCommand):
    def __init__(self, table, snapshot, col, header_text, col_data, on_changed=None):
        super().__init__(f"Delete column '{header_text}'")
        self.table = table
        self.snapshot = snapshot
        self.col = col
        self.header_text = header_text
        self.col_data = col_data
        self.on_changed = on_changed

    def _resync(self):
        self.snapshot.clear()
        self.snapshot.update(snapshot_table(self.table))
        if self.on_changed:
            self.on_changed()

    def redo(self):
        self.table.blockSignals(True)
        self.table.removeColumn(self.col)
        self.table.blockSignals(False)
        self._resync()

    def undo(self):
        self.table.blockSignals(True)
        self.table.insertColumn(self.col)
        self.table.setHorizontalHeaderItem(
            self.col, QTableWidgetItem(self.header_text)
        )
        for r, text in enumerate(self.col_data):
            self.table.setItem(r, self.col, QTableWidgetItem(text))
        self.table.blockSignals(False)
        self._resync()


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
