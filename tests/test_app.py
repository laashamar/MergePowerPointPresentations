
import pytest

# Import from the package
from merge_powerpoint.app import AppController
from merge_powerpoint.powerpoint_core import PowerPointMerger


@pytest.fixture
def app_controller():
    """Returns a clean AppController instance for each test."""
    return AppController()


def test_app_controller_initialization():
    """
    Test that the AppController initializes correctly.
    AppController inherits from PowerPointMerger and should have an empty file list.
    """
    controller = AppController()
    # AppController inherits from PowerPointMerger, so it has get_files() method
    assert controller.get_files() == []
    # Verify it's an instance of both AppController and PowerPointMerger
    assert isinstance(controller, AppController)
    assert isinstance(controller, PowerPointMerger)


def test_add_files(app_controller):
    """
    Test adding files using the add_files method from PowerPointMerger.
    """
    files = ['file1.pptx', 'file2.pptx']
    app_controller.add_files(files)

    assert app_controller.get_files() == ['file1.pptx', 'file2.pptx']


def test_add_files_with_duplicates(app_controller):
    """
    Test that duplicate files are not added.
    """
    app_controller.add_files(['file1.pptx', 'file2.pptx'])
    app_controller.add_files(['file1.pptx', 'file3.pptx'])

    # file1.pptx should not be duplicated
    assert app_controller.get_files() == ['file1.pptx', 'file2.pptx', 'file3.pptx']


def test_remove_file(app_controller):
    """
    Test removing a file using the remove_file method from PowerPointMerger.
    """
    app_controller.add_files(['file1.pptx', 'file2.pptx'])
    app_controller.remove_file('file1.pptx')

    assert app_controller.get_files() == ['file2.pptx']


def test_remove_files(app_controller):
    """
    Test removing multiple files using the remove_files method.
    """
    app_controller.add_files(['file1.pptx', 'file2.pptx', 'file3.pptx'])
    app_controller.remove_files(['file1.pptx', 'file3.pptx'])

    assert app_controller.get_files() == ['file2.pptx']


def test_clear_files(app_controller):
    """
    Test clearing all files using the clear_files method.
    """
    app_controller.add_files(['file1.pptx', 'file2.pptx'])
    app_controller.clear_files()

    assert app_controller.get_files() == []


def test_move_file_up(app_controller):
    """
    Test moving a file up in the list.
    """
    app_controller.add_files(['file1.pptx', 'file2.pptx', 'file3.pptx'])
    result = app_controller.move_file_up(2)

    assert result is True
    assert app_controller.get_files() == ['file1.pptx', 'file3.pptx', 'file2.pptx']


def test_move_file_down(app_controller):
    """
    Test moving a file down in the list.
    """
    app_controller.add_files(['file1.pptx', 'file2.pptx', 'file3.pptx'])
    result = app_controller.move_file_down(0)

    assert result is True
    assert app_controller.get_files() == ['file2.pptx', 'file1.pptx', 'file3.pptx']
