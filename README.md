# **PowerPoint Presentation Merger**

A powerful Python GUI application for merging multiple PowerPoint (.pptx) files into a single presentation with perfect fidelity. Uses COM automation to ensure all formatting, animations, and embedded content are preserved during the merge process.

## **Features**

* **Modern Two-Column GUI**: Intuitive single-window interface with drag-and-drop support
* **Perfect Slide Copying**: COM automation preserves all formatting, animations, and embedded content
* **File Management**: Easy file selection with validation, reordering, and drag-and-drop capabilities
* **Smart File Validation**: Automatic detection of invalid file types, duplicates, and permission issues
* **Flexible Output**: Choose output location and filename
* **Automatic Slideshow**: Launches merged presentation immediately after creation
* **Comprehensive Logging**: Optional live logging for debugging and troubleshooting
* **Error Handling**: Robust error management with clear user feedback
* **Dual Entry Points**: Choose between standard mode or debug mode with live logging

## **System Requirements**

### **Operating System**

* **Windows 7/8/10/11** (COM automation requires Windows)

### **Software Dependencies**

* **Python 3.6 or higher** (for source code execution)  
* **Microsoft PowerPoint** (must be installed and licensed)

### **Hardware Requirements**

* **Memory**: 4GB RAM minimum (8GB recommended for large presentations)  
* **Storage**: 100MB available space  
* **Display**: 1024x768 minimum resolution

## **Installation**

### **Option 1: Pre-built Executable (Recommended for End Users)**

1. Download the latest release from the [Releases page](https://github.com/laashamar/MergePowerPointPresentations/releases).  
2. Extract the archive to your desired location.  
3. Run MergePowerPoint.exe directly.

### **Option 2: From Source Code (For Developers)**

1. **Clone the repository**:  
   git clone \[https://github.com/laashamar/MergePowerPointPresentations.git\](https://github.com/laashamar/MergePowerPointPresentations.git)  
   cd MergePowerPointPresentations

2. **Install dependencies**:  
   pip install \-r requirements.txt

3. **Run the application**:  
   \# Standard mode  
   python main.py

   \# Debug mode with live logging  
   python run\_with\_logging.py

## **Usage Guide**

### **Quick Start**

1. **Launch the application** using one of the methods above.
2. **Add PowerPoint files** by dragging and dropping them or using the "Browse for Files" button.
3. **Arrange the files** in the desired merge order using the up/down arrow buttons.
4. **Select output location** and enter a filename.
5. **Click "Merge Presentations"** to create your merged file.
6. **Enjoy your merged presentation** with automatic slideshow launch.

### **Detailed Workflow**

#### **Adding Files to the Merge Queue**

* **Drag and Drop**: Simply drag PowerPoint files from your file explorer and drop them into the merge queue area.
* **Browse Button**: Click "Browse for Files" to select files using a file dialog.
* **Supported Formats**: Only .pptx and .ppsx files are accepted.
* **Duplicate Prevention**: The application automatically detects and prevents duplicate file additions.
* **Permission Checking**: Files are validated for read access before being added.

#### **Managing the Merge Queue**

* **Reorder Files**: Use the ↑ and ↓ buttons on each file card to change the merge order.
* **Remove Files**: Click the ✕ button on any file card to remove it from the queue.
* **View Full Path**: Hover over a filename to see its complete file path in a tooltip.
* **Clear Queue**: Use the "Clear Queue" button to remove all files at once.

#### **Configuring Output Settings**

* **Output Location**: Click "Browse" to select the folder where the merged file will be saved.
* **Output Filename**: Enter a name for your merged presentation (the .pptx extension is added automatically if omitted).
* **File Overwrite**: If a file with the same name exists, you'll be prompted to confirm overwriting.

#### **Merging Presentations**

* Click **"Merge Presentations"** when you're ready.
* The status label will show progress during the merge operation.
* Upon successful completion:
  * A success message displays the saved file location.
  * The merged presentation automatically opens in slideshow mode.
* If any errors occur, detailed error messages will guide you to resolve them.

## **Application Modes**

### **Standard Mode (main.py)**

* **Best for**: Regular usage and end users  
* **Features**: Clean GUI workflow without logging overhead.  
* **Command**: python main.py

### **Debug Mode (run\_with\_logging.py)**

* **Best for**: Troubleshooting and development  
* **Features**:  
  * Live logging window with real-time status updates.  
  * Detailed error reporting and diagnostics.  
  * Log file saved to Downloads/merge\_powerpoint.log.  
* **Command**: python run\_with\_logging.py

## **Documentation**

* 🏗️ [**ARCHITECTURE.md**](https://www.google.com/search?q=docs/ARCHITECTURE.md) \- Technical architecture and design patterns  
* 📝 [**CHANGELOG.md**](https://www.google.com/search?q=docs/CHANGELOG.md) \- Version history and release notes  
* 🚀 [**PLANNED\_FEATURE\_ENHANCEMENTS.md**](https://www.google.com/search?q=docs/PLANNED_FEATURE_ENHANCEMENTS.md) \- Planned features and roadmap  
* 🤝 [**CONTRIBUTING.md**](https://www.google.com/search?q=docs/CONTRIBUTING.md) \- How to contribute to the project.  
* 📜 [**CODE\_OF\_CONDUCT.md**](https://www.google.com/search?q=docs/CODE_OF_CONDUCT.md) \- Community guidelines.

## **License**

This project is licensed under the MIT License \- see the [LICENSE](https://www.google.com/search?q=LICENSE) file for details.
