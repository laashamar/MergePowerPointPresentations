# PowerPoint Presentation Merger - Development Instructions

## Overview

This document provides comprehensive development instructions for the PowerPoint Presentation Merger project. It serves as a guide for AI code writers and developers working on this application.

---

## Development Principles

### Core Guidelines

1. **Accurate File Naming**: Always use descriptive and accurate file names to ensure clarity and traceability
2. **PEP8 Compliance**: Follow PEP8 standards rigorously when writing Python code to maintain best practices and readability
3. **Clarification First**: If prompts are unclear, illogical, or confusing, request clarification before proceeding
4. **Best Practices**: Follow established best practices for Markdown documentation and Python development

---

## Project Architecture

### Framework Requirements

- **Primary GUI Framework**: PySide6 (Qt for Python)
- **COM Automation**: comtypes library for PowerPoint integration
- **Testing Framework**: pytest with pytest-qt for GUI testing
- **Code Quality**: flake8, pylint, black for linting and formatting

### Key Components

| Component | Purpose | Technology |
|-----------|---------|------------|
| `gui.py` | User interface | PySide6/Qt |
| `app.py` | Application logic | Python |
| `powerpoint_core.py` | PowerPoint automation | comtypes |
| `app_logger.py` | Logging system | Python logging |

---

## Development Guidelines

### Code Quality Standards

#### Python Code Standards

```python
# Use descriptive variable names
presentation_files = []
merged_output_path = "output.pptx"

# Follow PEP8 formatting
def merge_presentations(input_files, output_path):
    """Merge multiple PowerPoint presentations into one."""
    pass
```

#### Documentation Standards

- Use clear, descriptive commit messages
- Document all public methods and classes
- Include type hints where appropriate
- Maintain comprehensive README documentation

### Error Handling

#### COM Automation Best Practices

```python
import comtypes
import comtypes.client

def robust_powerpoint_operation():
    """Example of robust COM operation with proper cleanup."""
    comtypes.CoInitialize()
    try:
        # PowerPoint operations here
        app = comtypes.client.CreateObject("PowerPoint.Application")
        # ... operations ...
    finally:
        # Always cleanup COM resources
        if 'app' in locals():
            app.Quit()
        comtypes.CoUninitialize()
```

### Threading Considerations

- **COM Initialization**: Each thread must initialize COM separately
- **Resource Cleanup**: Always use try/finally blocks for COM cleanup
- **UI Responsiveness**: Keep long operations in background threads

---

## File Structure

### Project Organization

```text
MergePowerPointPresentations/
├── docs/                    # Documentation
│   ├── ARCHITECTURE.md
│   ├── INSTRUCTIONS.md
│   └── CHANGELOG.md
├── tests/                   # Test suite
│   ├── test_gui.py
│   ├── test_app.py
│   ├── test_app_logger.py
│   └── test_powerpoint_core.py
├── app.py                   # Main application logic
├── gui.py                   # User interface
├── powerpoint_core.py       # PowerPoint automation
├── app_logger.py           # Logging configuration
├── main.py                 # Application entry point
└── requirements.txt        # Dependencies
```

---

## Development Workflow

### Setup Process

1. **Environment Setup**

   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # Windows
   pip install -r requirements.txt
   ```

2. **Development Dependencies**

   ```bash
   pip install pytest pytest-qt pytest-mock
   pip install flake8 pylint black isort
   ```

3. **Testing Setup**

   ```bash
   pytest tests/ -v
   ```

### Quality Assurance

#### Code Formatting

```bash
# Format code with Black
black --line-length=100 *.py

# Sort imports with isort
isort --profile black *.py

# Lint with flake8
flake8 *.py --max-line-length=100

# Advanced linting with pylint
pylint *.py --max-line-length=100
```

#### Testing

```bash
# Run all tests
pytest tests/ -v

# Run with coverage
pytest --cov=. tests/

# Run specific test module
pytest tests/test_gui.py -v
```

---

## Troubleshooting

### Common Issues

| Issue | Symptoms | Solution |
|-------|----------|----------|
| COM Errors | PowerPoint automation fails | Check COM initialization |
| Import Errors | Module not found | Verify virtual environment |
| Test Failures | Pytest errors | Check test dependencies |
| GUI Issues | Interface problems | Verify PySide6 installation |

### Debugging Commands

```bash
# Check Python environment
python --version
pip list

# Verify dependencies
pip check

# Run tests with verbose output
pytest -vv --tb=long tests/

# Check code quality
flake8 --statistics *.py
```

---

## Best Practices

### Code Organization

- **Single Responsibility**: Each module should have a clear, single purpose
- **Dependency Injection**: Pass dependencies through constructors
- **Error Handling**: Use specific exception types and proper cleanup
- **Testing**: Write tests for all public interfaces

### Documentation

- **README**: Keep project overview current and accurate
- **Docstrings**: Document all public methods and classes
- **Comments**: Explain complex logic and business rules
- **Architecture**: Maintain architectural documentation

### Version Control

- **Commit Messages**: Use conventional commit format
- **Branching**: Use feature branches for development
- **Testing**: Ensure all tests pass before merging
- **Documentation**: Update docs with code changes

---

## Resources

### Documentation Links

- [PySide6 Documentation](https://doc.qt.io/qtforpython/)
- [comtypes Documentation](https://pythonhosted.org/comtypes/)
- [pytest Documentation](https://docs.pytest.org/)
- [PEP8 Style Guide](https://pep8.org/)

### Project References

- [ARCHITECTURE.md](ARCHITECTURE.md) - Technical architecture details
- [../tests/INSTRUCTIONS.md](../tests/INSTRUCTIONS.md) - Test suite instructions
- [../README.md](../README.md) - Project overview and setup

---

## Contributing

### Development Process

1. **Setup**: Clone repository and set up development environment
2. **Branch**: Create feature branch from main/develop
3. **Develop**: Write code following established standards
4. **Test**: Ensure all tests pass and add new tests as needed
5. **Document**: Update documentation for any changes
6. **Review**: Submit pull request for code review
7. **Merge**: Merge after approval and testing

### Code Review Checklist

- [ ] Code follows PEP8 standards
- [ ] All tests pass successfully
- [ ] Documentation is updated
- [ ] Error handling is comprehensive
- [ ] COM resources are properly managed
- [ ] No tkinter dependencies remain
- [ ] Type hints are included where appropriate

---

**Last Updated**: 2025-10-09  
**Version**: 2.0  
**Maintainer**: Development Team