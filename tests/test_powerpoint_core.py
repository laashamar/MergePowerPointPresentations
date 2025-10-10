# tests/test_powerpoint_core.py

"""
Tests for the PowerPoint core functionalities.
"""

import os
import pytest
from unittest.mock import MagicMock, patch, ANY
from powerpoint_core import PowerPointMerger, PowerPointError

# Test data
TEST_FILE_1 = "test1.pptx"
TEST_FILE_2 = "test2.pptx"
OUTPUT_FILE = "merged.pptx"


@pytest.fixture
def merger():
    """Fixture for a PowerPointMerger instance."""
    return PowerPointMerger()


def test_powerpoint_merger_initialization(merger):
    """Test that a PowerPointMerger instance initializes correctly."""
    assert merger is not None
    assert merger.get_files() == []


def test_add_files(merger):
    """Test adding files to the merger."""
    merger.add_files([TEST_FILE_1, TEST_FILE_2])
    assert len(merger.get_files()) == 2
    assert TEST_FILE_1 in merger.get_files()
    assert TEST_FILE_2 in merger.get_files()


def test_add_files_no_duplicates(merger):
    """Test that adding duplicate files doesn't create duplicates."""
    merger.add_files([TEST_FILE_1])
    merger.add_files([TEST_FILE_1])
    assert len(merger.get_files()) == 1


def test_remove_file(merger):
    """Test removing a file from the merger."""
    merger.add_files([TEST_FILE_1, TEST_FILE_2])
    merger.remove_file(TEST_FILE_1)
    assert len(merger.get_files()) == 1
    assert TEST_FILE_1 not in merger.get_files()
    assert TEST_FILE_2 in merger.get_files()


def test_move_file_up(merger):
    """Test moving a file up in the list."""
    merger.add_files([TEST_FILE_1, TEST_FILE_2])
    merger.move_file_up(1)
    files = merger.get_files()
    assert files[0] == TEST_FILE_2
    assert files[1] == TEST_FILE_1


def test_move_file_down(merger):
    """Test moving a file down in the list."""
    merger.add_files([TEST_FILE_1, TEST_FILE_2])
    merger.move_file_down(0)
    files = merger.get_files()
    assert files[0] == TEST_FILE_2
    assert files[1] == TEST_FILE_1


def test_merge_no_files(merger):
    """Test that merging with no files raises PowerPointError."""
    with pytest.raises(PowerPointError, match="No files to merge"):
        merger.merge(OUTPUT_FILE)


def test_merge_with_files(merger):
    """Test merging files with a progress callback."""
    merger.add_files([TEST_FILE_1, TEST_FILE_2])
    progress_callback = MagicMock()
    
    result = merger.merge(OUTPUT_FILE, progress_callback)
    
    assert result is True
    # Verify progress callback was called for each file
    assert progress_callback.call_count == 2
    progress_callback.assert_any_call(1, 2)
    progress_callback.assert_any_call(2, 2)


def test_merge_without_callback(merger):
    """Test merging files without a progress callback."""
    merger.add_files([TEST_FILE_1])
    result = merger.merge(OUTPUT_FILE)
    assert result is True

