from unittest.mock import MagicMock

import pytest

from merge_powerpoint.powerpoint_core import PowerPointError, PowerPointMerger


@pytest.fixture
def merger():
    """Returns a clean PowerPointMerger instance for each test."""
    return PowerPointMerger()


def test_add_files(merger):
    """
    Test that files are added correctly and duplicates are ignored.
    """
    files = ['file1.pptx', 'file2.pptx']
    merger.add_files(files)
    assert merger.get_files() == ['file1.pptx', 'file2.pptx']

    # Test duplicate handling - adding same files again
    merger.add_files(['file1.pptx', 'file3.pptx'])
    assert merger.get_files() == ['file1.pptx', 'file2.pptx', 'file3.pptx']


def test_remove_file(merger):
    """
    Test that a file can be successfully removed from the list.
    """
    merger.add_files(['file1.pptx', 'file2.pptx', 'file3.pptx'])
    merger.remove_file('file2.pptx')
    assert merger.get_files() == ['file1.pptx', 'file3.pptx']

    # Test removing non-existent file (should not raise error)
    merger.remove_file('nonexistent.pptx')
    assert merger.get_files() == ['file1.pptx', 'file3.pptx']


def test_move_file_up(merger):
    """
    Test moving a file up in the list.
    """
    merger.add_files(['file1.pptx', 'file2.pptx', 'file3.pptx'])
    merger.move_file_up(2)  # Move file3 up
    assert merger.get_files() == ['file1.pptx', 'file3.pptx', 'file2.pptx']

    merger.move_file_up(1)  # Move file3 up again
    assert merger.get_files() == ['file3.pptx', 'file1.pptx', 'file2.pptx']

    # Test boundary - can't move first item up
    merger.move_file_up(0)
    assert merger.get_files() == ['file3.pptx', 'file1.pptx', 'file2.pptx']


def test_move_file_down(merger):
    """
    Test moving a file down in the list.
    """
    merger.add_files(['file1.pptx', 'file2.pptx', 'file3.pptx'])
    merger.move_file_down(0)  # Move file1 down
    assert merger.get_files() == ['file2.pptx', 'file1.pptx', 'file3.pptx']

    merger.move_file_down(1)  # Move file1 down again
    assert merger.get_files() == ['file2.pptx', 'file3.pptx', 'file1.pptx']

    # Test boundary - can't move last item down
    merger.move_file_down(2)
    assert merger.get_files() == ['file2.pptx', 'file3.pptx', 'file1.pptx']


def test_get_files(merger):
    """
    Test that the getter method returns the correct list of files.
    """
    assert merger.get_files() == []

    files = ['file1.pptx', 'file2.pptx']
    merger.add_files(files)
    assert merger.get_files() == files


def test_merge_with_no_files(merger):
    """
    Test that calling merge() with no files raises PowerPointError.
    """
    with pytest.raises(PowerPointError, match="No files to merge"):
        merger.merge('output.pptx')


def test_merge_success(merger):
    """
    Test successful merge with progress callback.
    """
    mock_callback = MagicMock()

    files = ['file1.pptx', 'file2.pptx', 'file3.pptx']
    merger.add_files(files)

    result = merger.merge('output.pptx', progress_callback=mock_callback)

    # Assert merge was successful
    assert result is True

    # Assert callback was called correct number of times (once per file)
    assert mock_callback.call_count == 3
    mock_callback.assert_any_call(1, 3)
    mock_callback.assert_any_call(2, 3)
    mock_callback.assert_any_call(3, 3)
