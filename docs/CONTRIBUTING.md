# **Contributing to PowerPoint Presentation Merger**

First off, thank you for considering contributing to this project\! Your help is greatly appreciated. Whether you're a developer, a user, or just someone with a good idea, you can make this project better.

## **How to Contribute**

There are many ways to contribute, from writing code and improving documentation to submitting bug reports and feature requests.

### **Reporting Bugs**

If you find a bug, please open an issue on our [GitHub Issues page](https://github.com/laashamar/MergePowerPointPresentations/issues).

When reporting a bug, please include as much detail as possible:

1. **Steps to reproduce** the issue.  
2. **Expected behavior** vs. the actual behavior you observed.  
3. **System information** (e.g., Windows version, Microsoft PowerPoint version).  
4. **Log files** (if you were using run\_with\_logging.py). The log is saved in your Downloads folder.  
5. **Screenshots** or a description of any error messages you received.

### **Suggesting Enhancements**

If you have an idea for a new feature or an improvement to an existing one, feel free to open an issue to discuss it. We track planned enhancements in the [PLANNED_FEATURE_ENHANCEMENTS.md](PLANNED_FEATURE_ENHANCEMENTS.md) file.

## **Your First Code Contribution**

Ready to contribute code? Here’s how to get started.

### **For Developers**

#### **Development Setup**

1. **Clone and Setup**: Get a local copy of the project
   ```bash
   git clone https://github.com/laashamar/MergePowerPointPresentations.git
   cd MergePowerPointPresentations
   ```

2. **Create Virtual Environment**: (Recommended)
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install in Development Mode**: Install the package with all development tools
   ```bash
   pip install -e ".[dev]"
   ```

#### **Making Changes**

1. **Pick an Issue**: Choose an open issue from the [Issues tab](https://github.com/laashamar/MergePowerPointPresentations/issues) that interests you. We recommend starting with issues labeled `good first issue`
2. **Create a Branch**: Create a new branch for your changes
   ```bash
   git checkout -b feature/issue-123
   ```
3. **Implement**: Make your changes in the `src/merge_powerpoint/` directory
4. **Format and Lint**: Ensure code quality
   ```bash
   black src/merge_powerpoint/
   ruff check src/merge_powerpoint/
   ```
5. **Test Thoroughly**: Run tests to ensure nothing breaks
   ```bash
   pytest tests/
   ```
6. **Submit a Pull Request**: Create a PR with a detailed description of your changes

### **For Users**

Not a developer? You can still help\!

1. **Report Bugs**: Help us find and fix issues.  
2. **Request Features**: Let us know what you'd like to see in the application.  
3. **Test Releases**: Help test new versions before they are officially released.  
4. **Provide Feedback**: Share your experience and suggestions for improvement.

## **Coding Standards**

To maintain consistency across the project, we adhere to the following standards:

### **Python Code**

* All Python code must follow the [**PEP 8 style guide**](https://www.python.org/dev/peps/pep-0008/)
* **Use Black for formatting**: All code must be formatted with Black (100 character line length)
* **Pass Ruff linting**: Code must pass Ruff checks with no violations
* Use clear and descriptive variable and function names
* **Include comprehensive docstrings**: Follow PEP 257 for modules, classes, and functions
* Add comments to explain complex or non-obvious parts of the code

### **Package Structure**

* All new Python modules should be added to `src/merge_powerpoint/`
* Follow the established package structure with proper imports
* Update `__init__.py` exports when adding new public APIs
* Maintain backward compatibility with root-level compatibility shims if needed

### **Code Quality Tools**

Before submitting a pull request, ensure your code passes all quality checks:

```bash
# Format code with Black
black src/merge_powerpoint/

# Check linting with Ruff
ruff check src/merge_powerpoint/

# Run tests
pytest tests/
```

### **Markdown Files**

* Use standard Markdown formatting for all documentation (.md files)
* Use headings (\#, \#\#, etc.) to structure documents logically
* Use code blocks with syntax highlighting for code examples
* Keep lines to a reasonable length to improve readability
