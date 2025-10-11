# Test Suite - PowerPoint Presentation Merger

## Overview

This directory contains comprehensive unit tests for the PowerPoint Presentation Merger application. All tests follow PEP8 standards and best practices.

## Test Structure

- `test_app.py` - Tests for the main application orchestration (app.py)
- `test_gui.py` - Tests for GUI components (gui.py)
- `test_logger.py` - Tests for logging infrastructure (logger.py)
- `test_powerpoint_core.py` - Tests for PowerPoint COM automation (powerpoint_core.py)
- `conftest.py` - Pytest configuration and shared fixtures

## Running Tests

### Quick Start

```bash

# Run all tests
pytest

# Run with verbose output
pytest -v

# Run with coverage
pytest --cov=. --cov-report=html

```

### Test Categories

Tests can be run by category using markers:

```bash

# Run only unit tests
pytest -m unit

# Run GUI tests
pytest tests/test_gui.py

```

## Test Statistics

- **Total Tests**: 32
- **Test Coverage**: 72%
- **Success Rate**: 100%

### Coverage by Module

| Module | Coverage | Notes |
|--------|----------|-------|
| app.py | 100% | Full coverage |
| logger.py | 96% | Near complete coverage |
| powerpoint_core.py | 83% | COM automation mocked |
| gui.py | 48% | GUI components partially testable |

## Dependencies

All test dependencies are listed in `requirements-dev.txt`:

```bash

pip install -r requirements-dev.txt

```

## Writing New Tests

Follow the AAA pattern:

1. **Arrange** - Set up test data and mocks
2. **Act** - Execute the code being tested
3. **Assert** - Verify the expected outcome

Example:

```python

def test_example(self):
    """Test that example works correctly."""

    # Arrange
    mock_data = Mock()
    
    # Act
    result = function_to_test(mock_data)
    
    # Assert
    self.assertEqual(result, expected_value)

```

## Mocking Strategy

- **win32com** - Mocked in conftest.py for Windows COM automation
- **tkinter** - Mocked to avoid creating actual GUI windows
- **File system** - Mocked where appropriate for isolation

## Continuous Integration

Tests run automatically on:

- Push to main, develop, tests, or copilot/** branches
- Pull requests to main, develop, or tests branches

See `.github/workflows/ci.yml` for CI/CD configuration.

## Additional Resources

See `Test_Envir.md` in the project root for comprehensive testing documentation.
