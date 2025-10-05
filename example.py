"""
Example usage of the PowerPoint Merger application.

This script provides information about using the PowerPoint Merger.
For normal usage, simply run: python main.py

Note: This application uses COM automation, so sample presentations
should be created manually using PowerPoint or any other PPTX creation tool.
"""


def show_usage_info():
    """Display usage information for the PowerPoint Merger."""
    print("=" * 60)
    print("PowerPoint Merger - Usage Information")
    print("=" * 60)
    print()
    print("To use the PowerPoint Merger application:")
    print()
    print("1. Run the application:")
    print("   python main.py")
    print()
    print("2. Follow the step-by-step workflow:")
    print("   - Step 1: Enter the number of files to merge")
    print("   - Step 2: Select your PowerPoint files")
    print("   - Step 3: Enter the output filename")
    print("   - Step 4: Reorder files using Move Up/Down buttons")
    print()
    print("3. Click 'Create New File' to merge and launch slideshow")
    print()
    print("=" * 60)
    print("Key Features:")
    print("=" * 60)
    print("- Uses COM automation for perfect slide copying")
    print("- Preserves all formatting, animations, and content")
    print("- Automatic slideshow launch after merging")
    print("- Easy file reordering with Move Up/Down buttons")
    print()
    print("Requirements:")
    print("- Windows OS")
    print("- Microsoft PowerPoint installed")
    print("- Python 3.6 or higher")
    print("=" * 60)


if __name__ == "__main__":
    show_usage_info()

