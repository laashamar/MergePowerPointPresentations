"""
A simple diagnostic script to verify that the pytestqt module
can be imported by the Python interpreter directly and to inspect
the interpreter's module search path (sys.path).

This version has been corrected to use the proper module name, 'pytestqt'.
"""

import pprint
import sys

# Correct module name based on successful user testing
MODULE_NAME_TO_CHECK = "pytestqt"

print("--- Checking Python Environment ---")
print(f"Python Executable: {sys.executable}")
print("-" * 30)

# Print the module search path for detailed diagnosis.
print("--- Python's Module Search Path (sys.path) ---")
pprint.pprint(sys.path)
print("-" * 30)


try:
    # Attempt to import the module that pytest is failing to find.
    # __import__ is used to import a module from a string variable.
    module = __import__(MODULE_NAME_TO_CHECK)
    print(f"SUCCESS: The '{MODULE_NAME_TO_CHECK}' module was imported successfully.")
    print(f"Module found at: {module.__file__}")

except ImportError as e:
    print(f"FAILURE: The '{MODULE_NAME_TO_CHECK}' module could not be imported.")
    print(f"Error: {e}")

except Exception as e:
    print(f"An unexpected error occurred: {e}")

print("-" * 30)
