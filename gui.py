"""Compatibility shim for gui module.

This module provides backward compatibility for imports.
All functionality has been moved to src/merge_powerpoint/gui.py
"""

import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent / "src"
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

from merge_powerpoint.gui import MainWindow  # noqa: E402, F401

__all__ = ["MainWindow"]
