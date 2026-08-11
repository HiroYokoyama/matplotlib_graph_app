# -*- coding: utf-8 -*-
from unittest.mock import MagicMock
from hygrapher.utils import get_font_list, apply_major_ticker, apply_minor_ticker


def test_get_font_list():
    fonts = get_font_list()
    assert isinstance(fonts, list)
    assert len(fonts) > 0


def test_apply_major_ticker():
    mock_axis = MagicMock()
    apply_major_ticker(mock_axis, "5.0", is_log_scale=False)
    assert mock_axis.set_major_locator.called

    mock_axis.reset_mock()
    apply_major_ticker(mock_axis, "invalid", is_log_scale=False)
    assert not mock_axis.set_major_locator.called

    mock_axis.reset_mock()
    apply_major_ticker(mock_axis, "5.0", is_log_scale=True)
    assert not mock_axis.set_major_locator.called


def test_apply_minor_ticker():
    mock_axis = MagicMock()
    apply_minor_ticker(
        mock_axis, show_minor=True, interval_str="1.0", is_log_scale=False
    )
    assert mock_axis.set_minor_locator.called

    mock_axis.reset_mock()
    apply_minor_ticker(
        mock_axis, show_minor=False, interval_str="1.0", is_log_scale=False
    )
    assert mock_axis.set_minor_locator.called

    mock_axis.reset_mock()
    apply_minor_ticker(mock_axis, show_minor=True, interval_str="", is_log_scale=True)
    assert mock_axis.set_minor_locator.called
