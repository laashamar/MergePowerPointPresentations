# **PowerPoint Presentation Merger**

A powerful Python GUI application for merging multiple PowerPoint (.pptx) files into a single presentation with perfect fidelity. Uses COM automation to ensure all formatting, animations, and embedded content are preserved during the merge process.

## **Features**

* **Step-by-Step GUI Workflow**: Intuitive 4-step process for merging presentations  
* **Perfect Slide Copying**: COM automation preserves all formatting, animations, and embedded content  
* **File Management**: Easy file selection with validation and reordering capabilities  
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
2. **Follow the 4-step workflow** described below.  
3. **Enjoy your merged presentation** with automatic slideshow launch.

### **Detailed Workflow**

#### **Step 1: Number of Files**

* Enter the number of PowerPoint files you want to merge.  
* Must be a positive integer.  
* Press **Enter** or click **Next** to continue.

#### **Step 2: File Selection**

* Click **"Add Files from Disk"** to browse for .pptx files.  
* Select exactly the number of files specified in Step 1\.  
* Only PowerPoint (.pptx) files are accepted.  
* Duplicate selections are automatically prevented.

#### **Step 3: Output Filename**

* Enter a name for your merged presentation.  
* The .pptx extension is added automatically if omitted.  
* Press **Enter** or click **Next** to continue.

#### **Step 4: File Ordering**

* Use **"Move Up"** and **"Move Down"** buttons to arrange files.  
* Files will be merged in the order shown (top to bottom).  
* Select a file in the list before using move buttons.  
* Click **"Create New File"** to start the merge process.

#### **Step 5: Merge and Launch**

* The application automatically merges your presentations.  
* Merged presentation is saved to your specified location.  
* Slideshow launches automatically upon completion.

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
