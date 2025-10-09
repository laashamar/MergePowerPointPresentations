"""
Unit tests for the core application logic in app.py.
"""

import pytest
from app import PowerPointMerger


@pytest.fixture
def merger():
    """Returns a fresh instance of PowerPointMerger for each test."""
    return PowerPointMerger()


def test_initialization(merger):
    """Test that the PowerPointMerger initializes with an empty file list."""
    assert merger.files == []
    assert merger.ppt_core is None


def test_add_files(merger):
    """Test adding single and multiple files."""
    merger.add_files(["file1.pptx"])
    assert merger.files == ["file1.pptx"]
    merger.add_files(["file2.pptx", "file3.pptx"])
    assert merger.files == ["file1.pptx", "file2.pptx", "file3.pptx"]


def test_remove_file(merger):
    """Test removing a file by its index."""
    merger.files = ["a.pptx", "b.pptx", "c.pptx"]
    merger.remove_file(1)
    assert merger.files == ["a.pptx", "c.pptx"]
    merger.remove_file(0)
    assert merger.files == ["c.pptx"]


def test_remove_file_invalid_index(merger):
    """Test that removing a file with an invalid index does nothing."""
    merger.files = ["a.pptx", "b.pptx"]
    merger.remove_file(5)  # Index out of bounds
    assert merger.files == ["a.pptx", "b.pptx"]
    merger.remove_file(-1) # Negative index
    assert merger.files == ["a.pptx", "b.pptx"]


def test_move_file_up(merger):
    """Test moving a file up in the list."""
    merger.files = ["a.pptx", "b.pptx", "c.pptx"]
    merger.move_file_up(1)
    assert merger.files == ["b.pptx", "a.pptx", "c.pptx"]
    merger.move_file_up(2)
    assert merger.files == ["b.pptx", "c.pptx", "a.pptx"]


def test_move_file_up_at_top(merger):
    """Test that moving the top file up does nothing."""
    merger.files = ["a.pptx", "b.pptx"]
    merger.move_file_up(0)
    assert merger.files == ["a.pptx", "b.pptx"]


def test_move_file_down(merger):
    """Test moving a file down in the list."""
    merger.files = ["a.pptx", "b.pptx", "c.pptx"]
    merger.move_file_down(0)
    assert merger.files == ["b.pptx", "a.pptx", "c.pptx"]
    merger.move_file_down(0)
    assert merger.files == ["a.pptx", "b.pptx", "c.pptx"]


def test_move_file_down_at_bottom(merger):
    """Test that moving the bottom file down does nothing."""
    merger.files = ["a.pptx", "b.pptx"]
    merger.move_file_down(1)
    assert merger.files == ["a.pptx", "b.pptx"]

