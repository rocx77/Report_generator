# Code2Word

A Python tool to convert code files into a formatted Word report with output, screenshots, and student metadata.

## Quickstart

1. **Clone the repository**
2. **Create a virtual environment** (recommended):
   ```bash
   python -m venv c2wenv
   ```
3. **Activate the environment**:
   - Windows (PowerShell):
     ```bash
     .\c2wenv\Scripts\Activate.ps1
     ```
   - Windows (cmd):
     ```cmd
     .\c2wenv\Scripts\activate.bat
     ```
   - Linux/macOS:
     ```bash
     source c2wenv/bin/activate
     ```
4. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
5. **Run the tool**:
   ```bash
   python c2w.py <file1> <file2> ...
   ```

## Requirements
- Python 3.10+
- Chrome browser (for HTML/CSS/JS screenshots)
- GCC/G++ (for C/C++ code execution)

## Usage
- Enter assignment metadata when prompted.
- The Word report will be saved in the current directory.

## Troubleshooting
- If Chrome is not detected, set `BROWSER_BINARY_LOCATION` in `c2w.py`.
- For C/C++ files, ensure GCC/G++ is installed and in your PATH.

## License
MIT
