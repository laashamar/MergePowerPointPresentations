# Test Suite Instructions

## Overview

This document provides detailed instructions for resolving pytest test failures in the PowerPoint Presentation Merger project. The test suite consists of four main test modules that validate core functionality across different components.

## Objective

Analyze the provided pytest output and apply necessary code corrections to ensure all tests pass successfully.

---

## Test Modules Overview

| Test Module | Purpose | Key Components |
|-------------|---------|----------------|
| `test_gui.py` | GUI component testing | MainWindow initialization and UI elements |
| `test_app.py` | Application logic testing | PowerPointMerger functionality |
| `test_app_logger.py` | Logging system testing | Logger configuration and setup |
| `test_powerpoint_core.py` | Core PowerPoint integration | COM automation and error handling |

---

## Issue Resolution Guide

### 1. GUI Test Failures (`tests/test_gui.py`)

#### GUI Problem Description

All six GUI tests are failing with the following error:

```python
TypeError: MainWindow.__init__() missing 1 required positional argument: 'merger'
```

#### GUI Root Cause

The `main_app` fixture that sets up the `MainWindow` for testing is not passing the required `PowerPointMerger` instance to the constructor.

#### GUI Solution

Modify the `main_app` fixture to properly instantiate and pass the required `PowerPointMerger` object:

```python
# In tests/test_gui.py
from app import PowerPointMerger

@pytest.fixture
def main_app(qtbot):
    app = QApplication.instance() or QApplication([])
    merger = PowerPointMerger()  # Create the merger instance
    window = MainWindow(merger)  # Pass it to the constructor
    qtbot.addWidget(window)
    yield window
    window.close()
```

#### GUI Expected Outcome

- All GUI tests should initialize properly
- MainWindow constructor receives required dependencies
- Tests can properly interact with GUI components

---

### 2. Application Logic Test Failures (`tests/test_app.py`)

#### App Logic Problem Description

Multiple tests are failing due to incorrect attribute references:

- `test_initialization`
- `test_add_files`
- `test_remove_file`
- `test_move_file_up`
- `test_move_file_down`

#### App Logic Root Cause

Tests incorrectly reference `merger.files` attribute, but the actual class attribute is named `merger.file_paths`. This causes methods to operate on empty lists, leading to `IndexError` and incorrect assertions.

#### App Logic Solution

Perform a global find-and-replace operation in `tests/test_app.py`:

**Find:** `merger.files`  
**Replace:** `merger.file_paths`

#### Code Changes Required

Update all test methods to use the correct attribute name:

```python
# Before (incorrect)
assert len(merger.files) == 0

# After (correct)
assert len(merger.file_paths) == 0
```

#### App Logic Expected Outcome

- File manipulation tests work with correct data structure
- List operations (add, remove, move) function properly
- Assertions validate actual application state

---

### 3. Logger Test Failures (`tests/test_app_logger.py`)

#### Logger Problem Description

The `test_setup_logging_configures_correctly` test fails with:

```python
AssertionError: assert 20 == 10
```

#### Logger Root Cause

Test expects logging level to be `10` (DEBUG), but `app_logger.setup_logging` correctly configures it to `logging.INFO` (numerical value `20`).

#### Logger Solution

Update the assertion to expect the correct logging level:

```python
# In tests/test_app_logger.py
import logging

# Inside test_setup_logging_configures_correctly
assert call_args[1]['level'] == logging.INFO
```

#### Logger Expected Outcome

- Logging configuration test validates correct INFO level
- Logger setup matches application requirements
- Test assertions align with actual implementation

---

### 4. PowerPoint Core Test Failures (`tests/test_powerpoint_core.py`)

#### PowerPoint Problem 1: Initialization Failure Test

**Issue:** `test_powerpoint_core_initialization_failure` doesn't raise expected `PowerPointError`

**Root Cause:** Mock only simulates `GetActiveObject` failure, not subsequent `CreateObject` failure

**Solution:** Update mock to simulate both connection attempt failures:

```python
# In test_powerpoint_core_initialization_failure
mock_comtypes_client.GetActiveObject.side_effect = OSError
mock_comtypes_client.CreateObject.side_effect = OSError
```

#### PowerPoint Problem 2: COM Error Handling Test

**Issue:** `test_merge_presentations_handles_error` fails with:

```python
TypeError: COMError() takes exactly 3 arguments (0 given)
```

**Root Cause:** Mock raises `comtypes.COMError` without required constructor arguments

**Solution:** Properly instantiate `COMError` with required parameters:

```python
# In test_merge_presentations_handles_error
mock_powerpoint_app.Presentations.Add.return_value.Slides.InsertFromFile.side_effect = comtypes.COMError(
    "Mock COM Error", -1, "Mock description"
)
```

#### PowerPoint Core Expected Outcome

- Initialization failure scenarios properly tested
- COM error handling validates exception management
- PowerPoint automation edge cases covered

---

## Testing Best Practices

### Running Tests

```bash
# Run all tests
pytest tests/

# Run specific test module
pytest tests/test_gui.py

# Run with verbose output
pytest -v tests/

# Run with coverage report
pytest --cov=. tests/
```

### Test Environment Setup

1. **Virtual Environment**: Ensure tests run in isolated environment
2. **Dependencies**: Install test dependencies from `requirements-dev.txt`
3. **PowerPoint**: Mock COM interactions to avoid requiring PowerPoint installation
4. **Cleanup**: Properly dispose of test fixtures and resources

### Validation Checklist

- [ ] All test modules import correctly
- [ ] Fixtures provide required dependencies
- [ ] Mocks simulate expected behaviors
- [ ] Assertions validate actual vs expected outcomes
- [ ] Error scenarios properly tested
- [ ] Resource cleanup implemented

---

## Troubleshooting

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| Import errors | Missing dependencies | Check `requirements-dev.txt` |
| Fixture failures | Incorrect setup | Verify fixture dependencies |
| Mock failures | Wrong method signatures | Match actual API signatures |
| Assertion errors | Incorrect expectations | Align with implementation |

### Debug Commands

```bash
# Run tests with detailed output
pytest -vv --tb=long tests/

# Run specific failing test
pytest tests/test_gui.py::test_main_window_initialization -v

# Show test coverage gaps
pytest --cov=. --cov-report=html tests/
```

### Additional Resources

- [pytest Documentation](https://docs.pytest.org/)
- [pytest-qt Plugin](https://pytest-qt.readthedocs.io/)
- [Python Mock Library](https://docs.python.org/3/library/unittest.mock.html)

---

## Success Criteria

Upon successful implementation of all fixes:

1. ✅ All tests pass without errors
2. ✅ Test coverage remains comprehensive
3. ✅ Mocks properly simulate dependencies
4. ✅ Error scenarios appropriately handled
5. ✅ Code quality standards maintained

---

## Documentation References

**Last updated**: 2025-10-09

**For technical architecture details**: See [ARCHITECTURE.md](../docs/ARCHITECTURE.md)
