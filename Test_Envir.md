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
