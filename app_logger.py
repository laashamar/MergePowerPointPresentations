"""Compatibility shim for app_logger module.

This module provides backward compatibility for imports.
All functionality has been moved to src/merge_powerpoint/app_logger.py
"""

import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent / "src"
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

from merge_powerpoint.app_logger import setup_logging  # noqa: E402, F401

__all__ = ["setup_logging"]
