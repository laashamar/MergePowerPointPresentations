"""
Demo script showing the new Phase 3 features.

This script demonstrates how to use the progress callback feature
with the powerpoint_core module.
"""

import logging
import sys
import os

# Add parent directory to path to import powerpoint_core
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


def demo_progress_callback(filename, current_slide, total_slides):
    """
    Example progress callback function.

    Args:
        filename: Name of the file being processed
        current_slide: Current slide number
        total_slides: Total slides in the file
    """
    progress_percent = (current_slide / total_slides) * 100
    print(f"Progress: {filename} - Slide {current_slide}/{total_slides} ({progress_percent:.1f}%)")


def main():
    """Demonstrate Phase 3 features."""
    print("=" * 70)
    print("PowerPoint Merger - Phase 3 Features Demo")
    print("=" * 70)
    print()

    print("This demo shows the new Phase 3 features:")
    print()
    print("1. DRAG-AND-DROP FILE ADDITION")
    print("   - Drag .pptx files onto the application window")
    print("   - Only .pptx files are accepted")
    print("   - Duplicates are automatically prevented")
    print()

    print("2. DRAG-AND-DROP LIST REORDERING")
    print("   - Click and drag file labels to reorder")
    print("   - Files are numbered (1, 2, 3...)")
    print("   - Order changes update immediately")
    print()

    print("3. DYNAMIC STATUS FEEDBACK")
    print("   - Merge runs in separate thread")
    print("   - Real-time progress updates:")
    print('     * During: "Merging \\"filename.pptx\\" (slide X of Y)..."')
    print('     * Success: "Merge Complete!"')
    print('     * Failure: "Error: [message]"')
    print()

    print("4. POST-MERGE ACTIONS")
    print("   - Two buttons appear after successful merge:")
    print("     * 'Open Presentation' - Opens merged file")
    print("     * 'Show in Explorer' - Shows file in file manager")
    print("   - Cross-platform support (Windows, macOS, Linux)")
    print()

    print("=" * 70)
    print("Using the Progress Callback (Example)")
    print("=" * 70)
    print()

    # Note: This is just a demonstration of the API
    # Actual merge requires PowerPoint COM automation (Windows only)
    print("Example callback function:")
    print()
    print("def my_callback(filename, current_slide, total_slides):")
    print("    status = f'Processing {filename} - slide {current_slide}/{total_slides}'")
    print("    print(status)")
    print()
    print("# Use it with merge_presentations:")
    print("success, path, error = powerpoint_core.merge_presentations(")
    print("    file_order=['file1.pptx', 'file2.pptx'],")
    print("    output_filename='merged.pptx',")
    print("    progress_callback=my_callback")
    print(")")
    print()

    print("=" * 70)
    print("Testing callback function:")
    print("=" * 70)
    print()

    # Simulate progress callbacks
    demo_progress_callback("presentation1.pptx", 1, 10)
    demo_progress_callback("presentation1.pptx", 5, 10)
    demo_progress_callback("presentation1.pptx", 10, 10)
    demo_progress_callback("presentation2.pptx", 1, 5)
    demo_progress_callback("presentation2.pptx", 5, 5)

    print()
    print("=" * 70)
    print("For full GUI demo, run: python new_gui/main_gui.py")
    print("(Requires Windows with PowerPoint installed)")
    print("=" * 70)


if __name__ == "__main__":
    main()
