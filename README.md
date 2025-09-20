
# code2word

Generate Word document reports from code files (Python, C, C++, JS, HTML, CSS, PHP) with automatic output capture and frontend screenshots for web files. Includes both CLI and GUI usage.

## Features
- Converts code files into a formatted Word report
- Captures code output/errors for supported languages
- Takes full-page screenshots of HTML/CSS/JS files using Selenium
- Adds student/assignment metadata
- Easy-to-use GUI for file selection, metadata entry, and report generation

## Requirements
- Python 3.10+
- **Google Chrome browser must be installed** (required for HTML/CSS/JS screenshots and proper functioning)
- ChromeDriver (auto-managed by webdriver_manager)
- GCC/G++ (for C/C++ code execution)

### Python Packages
Install dependencies with:

```bash
pip install -r requirements.txt
```

## Usage

### Command Line Interface (CLI)
1. Place your code files in the project directory.
2. Run the script:
	```bash
	python c2w.py <file1> <file2> ...
	```
3. Enter assignment metadata when prompted.
4. The Word report will be saved in the current directory.

**Example:**
```bash
python c2w.py example.py example.html
```

### Graphical User Interface (GUI)
1. Run the GUI application:
   ```bash
   python c2w_APP.py
   ```
2. Fill in your assignment metadata in the form fields.
3. Click "Select Code Files" to choose your code files (supports multiple selection and reordering).
4. Click "Select Output Folder" to choose where the Word report will be saved.
5. Click "Generate Report" to create the Word document. Success/failure messages will be shown in the app.

### Standalone Executable (No Python Required)
1. Download the `Code2word.exe` file from the releases.
2. Double-click to run the GUI application directly.
3. No Python installation or dependencies required!

## Building Executable (For Developers)
To create your own executable:
1. Install PyInstaller: `pip install pyinstaller`
2. Build: `pyinstaller Code2word.spec`
3. Find the executable in the `dist/` folder.## Notes
- For HTML/CSS/JS files, Chrome must be installed and accessible.
- Screenshots are taken in headless mode; adjust `VIEWPORT_HEIGHT` in `c2w.py` if needed.
- For C/C++ files, GCC/G++ must be installed and available in PATH.

## Project Structure
- `c2w.py`: Main script for report generation (CLI)
- `c2w_APP.py`: GUI application for report generation
- `requirements.txt`: Python dependencies
- `INSTRUCTIONS.md`: Step-by-step setup and usage guide
- `.github/`: GitHub Actions workflow for CI
- `.gitignore`: Files and folders excluded from git

## License
MIT License
