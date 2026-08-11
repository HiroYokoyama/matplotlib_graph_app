# -*- coding: utf-8 -*-
"""
HYGrapher Package
"""

from .utils import get_app_version

__version__ = get_app_version()

from .main import GraphApp, main
from .main_3d import main as main_3d

__all__ = ["GraphApp", "main", "main_3d", "__version__"]
