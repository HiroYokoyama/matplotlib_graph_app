# -*- coding: utf-8 -*-
import os
os.environ["QT_QPA_PLATFORM"] = "offscreen"

import sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')

import pytest
from unittest.mock import patch

@pytest.fixture(autouse=True)
def mock_qt_dialogs():
    with patch("PyQt6.QtWidgets.QMessageBox.warning") as m_warn, \
         patch("PyQt6.QtWidgets.QMessageBox.critical") as m_crit, \
         patch("PyQt6.QtWidgets.QMessageBox.information") as m_info, \
         patch("PyQt6.QtWidgets.QMessageBox.about") as m_about, \
         patch("PyQt6.QtWidgets.QMessageBox.question") as m_quest:
        yield
