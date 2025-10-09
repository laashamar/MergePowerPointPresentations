# Test Environment Documentation

## Overview

This document describes the testing structure and CI/CD setup for the PowerPoint Presentation Merger application. The testing infrastructure follows PEP8 standards rigorously and implements best practices for Python testing.

## Test Structure

### Directory Layout

```
MergePowerPointPresentations/
├── tests/
│   ├── __init__.py
│   ├── test_app.py              # Tests for app.py
│   ├── test_gui.py              # Tests for gui.py
│   ├── test_logger.py           # Tests for logger.py
│   └── test_powerpoint_core.py  # Tests for powerpoint_core.py
├── .github/
│   └── workflows/
│       └── ci.yml               # GitHub Actions CI/CD pipeline
├── pytest.ini                   # Pytest configuration
├── .flake8                      # Flake8 linting configuration
├── .pylintrc                    # Pylint configuration
├── requirements.txt             # Production dependencies
└── requirements-dev.txt         # Development dependencies
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
# Install all development dependencies
pip install -r requirements-dev.txt

# Or install only testing tools
pip install pytest pytest-cov pytest-mock
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

#### Flake8

```bash
# Check all Python files
flake8 *.py

# With custom configuration
flake8 *.py --max-line-length=100
```

#### Pylint

```bash
# Check all Python files
pylint *.py

# Check specific module
pylint app.py
```

#### Black (Code Formatter)

```bash
# Check formatting without making changes
black --check *.py

# Format code automatically
black *.py
```

#### isort (Import Sorting)

```bash
# Check import ordering
isort --check-only *.py

# Fix import ordering
isort *.py
```

### Running All Quality Checks

```bash
# Run all checks in sequence
black --check *.py && \
isort --check-only *.py && \
flake8 *.py && \
pylint *.py
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

## Mocking Strategy

### Windows-Specific Dependencies

Since the application uses Windows COM automation, tests mock these components:

```python
@patch('powerpoint_core.win32com.client.Dispatch')
def test_function(self, mock_dispatch):
    mock_powerpoint = MagicMock()
    mock_dispatch.return_value = mock_powerpoint
    # Test code here
```

### GUI Components

GUI tests mock tkinter to avoid creating actual windows:

```python
@patch('gui.tk.Tk')
def test_window(self, mock_tk):
    mock_window = MagicMock()
    mock_tk.return_value = mock_window
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
