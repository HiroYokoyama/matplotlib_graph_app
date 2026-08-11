# -*- coding: utf-8 -*-
from unittest.mock import MagicMock, patch

from importlib.metadata import PackageNotFoundError

from hygrapher.utils import (
    get_app_version,
    get_font_list,
    apply_major_ticker,
    apply_minor_ticker,
    resolve_cli_file,
)


def test_get_app_version_reads_installed_metadata():
    # hygrapher is installed (editable) in this test environment, so this
    # should resolve to a real dotted version string, not the dev fallback.
    version = get_app_version()
    assert version != "0.0.0-dev"
    assert version.count(".") >= 2


def test_get_app_version_falls_back_when_not_installed():
    with patch(
        "hygrapher.utils._pkg_version", side_effect=PackageNotFoundError("hygrapher")
    ):
        assert get_app_version() == "0.0.0-dev"


def test_all_entry_points_agree_on_version():
    """pyproject.toml is the single source of truth; main.py, main_3d.py, and
    the package __init__ must all resolve to the same version rather than
    hardcoding their own copies that can drift out of sync."""
    import hygrapher
    from hygrapher.main import VERSION as version_2d
    from hygrapher.main_3d import VERSION as version_3d

    assert hygrapher.__version__ == get_app_version()
    assert version_2d == get_app_version()
    assert version_3d == get_app_version()


def test_resolve_cli_file_no_args():
    assert resolve_cli_file([]) is None


def test_resolve_cli_file_flags_only():
    assert resolve_cli_file(["--debug", "-v"]) is None


def test_resolve_cli_file_nonexistent_path():
    assert resolve_cli_file(["not_a_real_file.csv"]) is None


def test_resolve_cli_file_finds_existing_file(tmp_path):
    data_file = tmp_path / "data.csv"
    data_file.write_text("A,B\n1,2\n")

    assert resolve_cli_file([str(data_file)]) == str(data_file)


def test_resolve_cli_file_skips_flags_before_path(tmp_path):
    data_file = tmp_path / "data.csv"
    data_file.write_text("A,B\n1,2\n")

    assert resolve_cli_file(["--foo", str(data_file)]) == str(data_file)


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
