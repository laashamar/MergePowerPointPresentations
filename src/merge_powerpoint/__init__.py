"""PowerPoint Merger package for merging multiple PowerPoint presentations.

This package provides a graphical user interface and programmatic API for
merging PowerPoint presentations while preserving formatting, animations,
and embedded content.
"""

from merge_powerpoint.app import AppController
from merge_powerpoint.powerpoint_core import PowerPointError, PowerPointMerger

__version__ = "1.0.0"
__all__ = ["AppController", "PowerPointMerger", "PowerPointError"]
