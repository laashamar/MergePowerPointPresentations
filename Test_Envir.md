 copilot/implement-tests-structure-cicd
# Test Environment Documentation

## Overview

This document describes the testing structure and CI/CD setup for the PowerPoint Presentation Merger application. The testing infrastructure follows PEP8 standards rigorously and implements best practices for Python testing.

## Test Structure

### Directory Layout

```
MergePowerPointPresentations/
├── src/
│   └── merge_powerpoint/          # Main package
│       ├── __init__.py
│       ├── __main__.py
│       ├── app.py
│       ├── gui.py
│       ├── powerpoint_core.py
│       └── app_logger.py
├── tests/
│   ├── __init__.py
│   ├── conftest.py               # Shared fixtures
│   ├── test_app.py               # Tests for app.py
│   ├── test_gui.py               # Tests for gui.py
│   ├── test_app_logger.py        # Tests for app_logger.py
│   └── test_powerpoint_core.py   # Tests for powerpoint_core.py
├── .github/
│   └── workflows/
│       └── ci.yml                # GitHub Actions CI/CD pipeline
├── pyproject.toml                # Modern Python project config
├── pytest.ini                    # Pytest configuration (if separate)
└── requirements.txt              # Legacy compatibility
```

### Test Categories

Tests are organized using pytest markers:

- `@pytest.mark.unit` - Unit tests for individual functions/classes
- `@pytest.mark.integration` - Integration tests for component interactions
- `@pytest.mark.gui` - GUI-related tests (with mocking)
- `@pytest.mark.slow` - Tests that take longer to execute

## Running Tests

### Install Test Dependencies

```bash
# Install all dependencies including development tools
pip install -e ".[dev]"

# Or install from requirements.txt (legacy)
pip install -r requirements.txt
pip install pytest pytest-qt pytest-cov pytest-mock ruff black
```

### Run All Tests

```bash
# Run all tests with coverage
pytest

# Run with verbose output
pytest -v

# Run with coverage report
pytest --cov=. --cov-report=html
```

### Run Specific Test Categories

```bash
# Run only unit tests
pytest -m unit

# Run only GUI tests
pytest tests/test_gui.py

# Run specific test class
pytest tests/test_app.py::TestPowerPointMergerApp

# Run specific test function
pytest tests/test_app.py::TestPowerPointMergerApp::test_initialization
```

### Run Tests with Different Options

```bash
# Run tests and stop at first failure
pytest -x

# Run tests with detailed output
pytest -vv

# Run tests and show local variables on failure
pytest -l

# Run tests in parallel (requires pytest-xdist)
pytest -n auto
```

## Code Quality and Linting

### PEP8 Compliance

All Python code follows PEP8 standards strictly. Use the following tools to verify:

#### Ruff (Fast Modern Linter)

```bash
# Check all Python files
ruff check src/merge_powerpoint/

# With auto-fix
ruff check --fix src/merge_powerpoint/
```

#### Black (Code Formatter)

```bash
# Check formatting without making changes
black --check src/merge_powerpoint/

# Format code automatically
black src/merge_powerpoint/
```

#### isort (Import Sorting)

```bash
# Check import ordering
isort --check-only src/merge_powerpoint/

# Fix import ordering
isort src/merge_powerpoint/
```

### Running All Quality Checks

```bash
# Run all checks in sequence
black --check src/merge_powerpoint/ && \
isort --check-only src/merge_powerpoint/ && \
ruff check src/merge_powerpoint/
```

## CI/CD Pipeline

### GitHub Actions Workflow

The CI/CD pipeline (`.github/workflows/ci.yml`) runs automatically on:

- Push to `main`, `develop`, `tests`, or any `copilot/**` branches
- Pull requests to `main`, `develop`, or `tests` branches

### Pipeline Jobs

1. **Lint Job**
   - Checks code formatting with Black
   - Verifies import ordering with isort
   - Runs flake8 for PEP8 compliance
   - Runs pylint for code quality

2. **Test Job**
   - Runs all unit tests
   - Generates coverage reports
   - Uploads coverage to Codecov

3. **Test Matrix Job**
   - Tests on Python 3.9, 3.10, 3.11, and 3.12
   - Ensures cross-version compatibility

4. **Code Quality Job**
   - Additional PEP8 compliance checks
   - Complexity analysis

### Viewing CI/CD Results

1. Navigate to the GitHub repository
2. Click on the "Actions" tab
3. Select the workflow run to view details
4. Check individual job logs for any failures

## Configuration Files

### pytest.ini

Configures pytest behavior:
- Test discovery patterns
- Coverage settings
- Output formatting
- Test markers

### .flake8

Configures flake8 linting:
- Maximum line length: 100 characters
- Ignored files and directories
- Ignored error codes (E203, W503 for Black compatibility)

### .pylintrc

Configures pylint:
- Custom scoring thresholds
- Disabled checks for external libraries
- Code complexity limits
- Naming conventions

## Writing New Tests

### Test Structure

Follow this structure when writing new tests:

```python
"""
Unit tests for module_name.py module.

Brief description of what is being tested.
"""
import unittest
from unittest.mock import Mock, patch, MagicMock


class TestClassName(unittest.TestCase):
    """Test cases for ClassName."""

    def setUp(self):
        """Set up test fixtures before each test."""
        # Initialize test objects here
        pass

    def tearDown(self):
        """Clean up after each test."""
        # Clean up resources here
        pass

    def test_specific_behavior(self):
        """Test that specific behavior works as expected."""
        # Arrange
        # Act
        # Assert
        pass
```

### Testing Best Practices

1. **Use Descriptive Test Names**
   - Test names should describe what is being tested
   - Use format: `test_<function>_<scenario>_<expected_result>`

2. **Follow AAA Pattern**
   - **Arrange**: Set up test data and mocks
   - **Act**: Execute the code being tested
   - **Assert**: Verify the expected outcome

3. **Mock External Dependencies**
   - Mock file system operations
   - Mock COM automation (win32com)
   - Mock GUI components (tkinter)

4. **Test Edge Cases**
   - Test with empty inputs
   - Test with invalid inputs
   - Test error handling paths

5. **Keep Tests Independent**
   - Each test should run independently
   - Don't rely on test execution order
   - Clean up resources in tearDown()

## Coverage Goals

- **Target**: 80%+ code coverage
- **Minimum**: 70% code coverage for CI to pass
- **Focus Areas**:
  - All business logic (100%)
  - Error handling paths
  - Input validation

### Mocking Strategy

### Windows-Specific Dependencies

Since the application uses Windows COM automation, tests mock these components:

```python
@patch('powerpoint_core.comtypes.client.CreateObject')
def test_function(self, mock_create_object):
    mock_powerpoint = MagicMock()
    mock_create_object.return_value = mock_powerpoint
    # Test code here
```

### GUI Components

GUI tests mock PySide6 to avoid creating actual windows:

```python
@patch('gui.QApplication')
@patch('gui.MainWindow')
def test_window(self, mock_window, mock_qapp):
    mock_app = MagicMock()
    mock_qapp.return_value = mock_app
    # Test code here
```

## Troubleshooting

### Common Issues

1. **Import Errors**
   - Ensure all dependencies are installed
   - Check Python path includes project root

2. **Test Discovery Issues**
   - Verify test files start with `test_`
   - Ensure test functions start with `test_`
   - Check pytest.ini configuration

3. **Mocking Issues**
   - Verify mock patches target the correct module
   - Use `patch.object()` for class methods
   - Check mock return values are configured

4. **Coverage Reports**
   - Ensure pytest-cov is installed
   - Check coverage configuration in pytest.ini
   - View HTML report in `htmlcov/index.html`

## Continuous Integration

### Pre-commit Checks

Before committing code:

```bash
# Format code
black *.py
isort *.py

# Run linters
flake8 *.py
pylint *.py

# Run tests
pytest -v
```

### Commit Hooks (Optional)

Consider setting up pre-commit hooks:

```bash
# Install pre-commit
pip install pre-commit

# Create .pre-commit-config.yaml
# Run pre-commit install
pre-commit install
```

## Additional Resources

- [pytest documentation](https://docs.pytest.org/)
- [PEP 8 Style Guide](https://pep8.org/)
- [unittest.mock documentation](https://docs.python.org/3/library/unittest.mock.html)
- [GitHub Actions documentation](https://docs.github.com/en/actions)

## Maintenance

### Updating Dependencies

```bash
# Update test dependencies
pip install --upgrade pytest pytest-cov pytest-mock

# Update linting tools
pip install --upgrade flake8 pylint black isort

# Regenerate requirements-dev.txt
pip freeze > requirements-dev.txt
```

### Adding New Tests

1. Create test file in `tests/` directory
2. Follow naming convention: `test_<module>.py`
3. Import module to test
4. Write test classes and functions
5. Run tests locally before committing
6. Verify CI pipeline passes

## Summary

This testing infrastructure provides:

- ✅ Comprehensive unit test coverage
- ✅ PEP8 compliance checking
- ✅ Automated CI/CD with GitHub Actions
- ✅ Code quality monitoring
- ✅ Cross-version Python compatibility testing
- ✅ Coverage reporting

All code changes should maintain or improve test coverage and pass all quality checks before being merged.
.Part 1:

 Implement the Pytest Testing StructureObjective: Create a robust testing structure using pytest that is correctly configured for your project and is runnable from the root directory.1. Create the Test DirectoryIn the project's root directory, create a new folder named tests.Inside the tests/ folder, create an empty file named __init__.py to mark it as a Python package.

2. Create Development Dependencies FileCreate a requirements-dev.txt file in the root directory. This separates testing and linting tools from the application's core dependencies.File: requirements-dev.txt# Development, testing, and linting dependencies
pytest
pytest-cov
flake8
pywin32 # Explicitly add for clarity in the dev environment

3. Create Tests for Core LogicCreate a new file at tests/test_powerpoint_core.py. This test uses a standard import, which pytest handles automatically when run from the project root.File: tests/test_powerpoint_core.py"""Tests for the core PowerPoint merging logic."""

import powerpoint_core

def test_merge_function_exists():
    """Verify that the core merge function can be imported and is callable."""
    # Note: The actual function in powerpoint_core.py is `merge_presentations`
    assert hasattr(powerpoint_core, "merge_presentations"), "merge_presentations function not found"
    assert callable(powerpoint_core.merge_presentations), "merge_presentations is not a callable function"

# Future tests for functionality (e.g., using mock files) can be added here.

4. Create Tests for the GUICreate a new file at tests/test_gui.py. This initial "smoke test" confirms that the main App class can be imported without syntax or dependency errors.File: tests/test_gui.py"""Tests for the GUI application."""

from gui import App

def test_app_can_be_imported():
    """Verify that the App class can be imported without errors."""
    # This test confirms the file is syntactically correct and imports are valid.
    # A full instantiation `app = App()` is avoided here as it may require a
    # running Tkinter event loop, which can be handled in more advanced
    # tests using mocking or specific GUI testing libraries.
    assert App is not None, "The App class could not be imported from gui.py"

5. Configure Pytest 
Create a pytest.ini file in the root directory to define the test paths and file patterns, ensuring pytest discovers your tests correctly.File: pytest.ini[pytest]
testpaths = tests
python_files = test_*.py

6. Update README.md with Test Instructions

Add a clear "Running Tests" section to your README.md file so anyone can run the test suite.Content for README.md## Running Tests

To run the automated tests for this project, first install the necessary dependencies and then execute `pytest` from the project's root directory.

```bash
# 1. Install core application dependencies
pip install -r requirements.txt

# 2. Install development and testing dependencies
pip install -r requirements-dev.txt

# 3. Run the test suite
pytest

---

## Part 2: Implement the GitHub Actions CI Workflow

**Objective:** Automate the building, linting, and testing of your application using a GitHub Actions workflow.

### 1. Create the Workflow File

Create the directory path `.github/workflows/` in your project root if it doesn't exist. Then, create the `python-ci.yml` file inside it with the following content.

**Key Change:** The workflow is configured to run on `windows-latest`. This is critical because your application's core logic depends on `pywin32` and COM automation, which are only available on Windows.

**File: `.github/workflows/python-ci.yml`**
```yaml
name: Python CI Workflow

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]

jobs:
  build-test-lint:
    name: Build, Test, and Lint
    runs-on: windows-latest # CRITICAL: Must be windows-latest for pywin32 COM automation

    steps:
    - name: 🧾 Checkout repository
      uses: actions/checkout@v4

    - name: 🐍 Set up Python
      uses: actions/setup-python@v5
      with:
        python-version: '3.11'

    - name: 📦 Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install -r requirements.txt
        if (Test-Path -Path "requirements-dev.txt") { pip install -r requirements-dev.txt }

    - name: 🧹 Run flake8 linting
      run: |
        # Stop the build if there are critical Python syntax errors or undefined names
        flake8 . --count --select=E9,F63,F7,F82 --show-source --statistics
        # Treat all other issues as warnings and check for complexity and line length
        flake8 . --count --exit-zero --max-complexity=10 --max-line-length=127 --statistics

    - name: 🧪 Run tests with coverage
      run: |
        pytest `
          --cov=powerpoint_core `
          --cov=gui `
          --cov=app `
          --cov-report=xml `
          --cov-report=term-missing `
          --junitxml=pytest-results.xml

    - name: 📊 Upload coverage to Codecov (Optional)
      uses: codecov/codecov-action@v4
      with:
        files: ./coverage.xml
        fail_ci_if_error: true
      # For this to work, a Codecov token must be set in the repository's secrets
      # with the name CODECOV_TOKEN

    - name: 📤 Upload test results
      if: always() # Ensures this step runs even if previous steps fail
      uses: actions/upload-artifact@v4
      with:
        name: pytest-results
        path: pytest-results.xml
 tests
