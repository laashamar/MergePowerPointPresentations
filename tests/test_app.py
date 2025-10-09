"""
Tests for the AppController.
"""
import pytest
from app import AppController

@pytest.fixture
def controller():
    """Fixture to create an AppController instance for tests."""
    return AppController()

def test_app_controller_initialization(controller):
    """Test that the AppController initializes correctly."""
    assert controller is not None
    # Since AppController inherits from PowerPointMerger, it should have its methods
    assert hasattr(controller, 'add_files')
    assert hasattr(controller, 'merge')
    assert controller.get_files() == []

def test_controller_inherits_merger_functionality(controller):
    """Test that AppController correctly inherits and uses PowerPointMerger methods."""
    files_to_add = ["presentation1.pptx", "presentation2.pptx"]
    controller.add_files(files_to_add)
    assert controller.get_files() == files_to_add

    controller.remove_file("presentation1.pptx")
    assert controller.get_files() == ["presentation2.pptx"]
