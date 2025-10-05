"""
Core module for PowerPoint presentation merging using COM automation.

This module provides the business logic for merging multiple PowerPoint
presentations using pywin32 COM automation, which ensures perfect copying
of all content, formatting, and animations.
"""

import os
import win32com.client


def merge_presentations(file_order, output_filename):
    """
    Merge multiple PowerPoint presentations into a single file using COM.

    Args:
        file_order: List of file paths in the order they should be merged
        output_filename: Name of the output file (should include .pptx extension)

    Returns:
        tuple: (success: bool, output_path: str, error_message: str or None)
    """
    powerpoint = None
    destination_prs = None
    source_prs = None

    try:
        # Initialize PowerPoint application
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True

        # Create a new presentation
        destination_prs = powerpoint.Presentations.Add()

        # Remove the default blank slide if it exists
        if destination_prs.Slides.Count > 0:
            destination_prs.Slides(1).Delete()

        # Process each source file in order
        for file_path in file_order:
            try:
                # Open source presentation
                source_prs = powerpoint.Presentations.Open(
                    os.path.abspath(file_path),
                    ReadOnly=True,
                    WithWindow=False
                )

                # Copy each slide from source to destination
                for slide_index in range(1, source_prs.Slides.Count + 1):
                    # Copy the slide
                    source_prs.Slides(slide_index).Copy()

                    # Paste it into the destination presentation
                    destination_prs.Slides.Paste()

                # Close source presentation
                source_prs.Close()
                source_prs = None

            except Exception as e:
                if source_prs:
                    source_prs.Close()
                raise Exception(
                    f"Failed to process file {os.path.basename(file_path)}: {str(e)}"
                )

        # Save the merged presentation
        output_path = os.path.abspath(output_filename)
        destination_prs.SaveAs(output_path)

        return True, output_path, None

    except Exception as e:
        return False, "", str(e)

    finally:
        # Clean up: close presentations but keep PowerPoint open for slideshow
        if source_prs:
            try:
                source_prs.Close()
            except:
                pass


def launch_slideshow(output_path):
    """
    Launch PowerPoint slideshow using COM automation.

    Args:
        output_path: Full path to the presentation file

    Returns:
        tuple: (success: bool, error_message: str or None)
    """
    powerpoint = None
    presentation = None

    try:
        # Get PowerPoint application instance
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True

        # Open the presentation
        presentation = powerpoint.Presentations.Open(
            os.path.abspath(output_path),
            WithWindow=True
        )

        # Start the slideshow
        presentation.SlideShowSettings.Run()

        return True, None

    except Exception as e:
        return False, str(e)
