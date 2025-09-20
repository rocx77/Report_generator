
# Code2Word

Generate Word document reports from code files (Python, C, C++, Java, JS, HTML, CSS) with automatic output capture, proper formatting, and frontend screenshots for web files. Features both CLI and GUI interfaces for easy use.

## Features
- Converts code files into a professionally formatted Word report
- Captures code output and errors for supported languages
- Takes full-page screenshots of HTML/CSS/JS files using Selenium
- Adds student/assignment metadata with proper formatting
- **Interactive input support**: Prompts users for input values when programs require them (Python, C, C++, Java)
- **Format specifier handling**: Correctly processes format specifiers like %d, %s in output
- **Smart pagination**: Optimizes page breaks to keep related content together
- Easy-to-use GUI for file selection, metadata entry, and report generation
- **Optimized performance**: Fast execution with responsive GUI

## Requirements
- Python 3.10+ (for development only, not needed for executable)
- **Google Chrome browser** (required for HTML/CSS/JS screenshots)
- GCC/G++ compiler (for C/C++ code execution)
- Java JDK (for Java code execution)

### Python Packages (For Development)
If you want to modify or run the source code directly, install dependencies with:

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
4. The executable bundles all necessary components in a single file for easy sharing.

## Building Executable (For Developers)
To create your own executable:
1. Set up a virtual environment: `python -m venv c2wenv`
2. Activate it: `.\c2wenv\Scripts\Activate.ps1` (Windows)
3. Install dependencies: `pip install -r requirements.txt pyinstaller`
4. Build: `python -m PyInstaller Code2word.spec`
5. Find the executable in the `dist/` folder.
6. Only the `Code2word.exe` file is needed for distribution.## Notes
- For HTML/CSS/JS files, Chrome must be installed and accessible.
- Screenshots are taken in headless mode; adjust `VIEWPORT_HEIGHT` in `c2w.py` if needed.
- For C/C++ files, GCC/G++ must be installed and available in PATH.
- For Java files, JDK must be installed and available in PATH.
- **Interactive Programs**: When your code requires user input (using `input()` in Python, `scanf()` in C/C++, or `Scanner` in Java), the tool will prompt you to provide the input values during execution.
- **Format Specifiers**: The tool now correctly handles format specifiers like %d, %s, %f in your code's output.
- **Optimized Layout**: Page breaks are intelligently placed to avoid splitting code blocks unnecessarily.

## Project Structure
- `c2w.py`: Core module for report generation
- `c2w_APP.py`: GUI application for report generation
- `requirements.txt`: Python dependencies
- `INSTRUCTIONS.md`: Step-by-step setup and usage guide
- `Code2word.spec`: PyInstaller specification for building executable
- `dist/Code2word.exe`: Standalone executable (no Python required)
- `chromedriver.exe`: WebDriver for Chrome automation
- `.github/`: GitHub Actions workflow for CI
- `.gitignore`: Files and folders excluded from git

## License
MIT License
