# Code2Word - Detailed Instructions

A tool to convert code files into a professionally formatted Word report with output capture, screenshots, and metadata.

## Installation Options

### Option 1: Use the Standalone Executable (Recommended)

1. **Download the executable**:
   - Get `Code2word.exe` from the releases section
   - No Python or other dependencies required

2. **Run the application**:
   - Double-click `Code2word.exe` to launch
   - If Windows SmartScreen appears, click "More info" and then "Run anyway"

### Option 2: Run from Source Code

1. **Clone the repository**:
   ```bash
   git clone https://github.com/rocx77/Report_generator.git
   cd Report_generator
   ```

2. **Create a virtual environment**:
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
   - GUI mode: `python c2w_APP.py`
   - CLI mode: `python c2w.py <file1> <file2> ...`

## Requirements
- Python 3.10+
- Chrome browser (for HTML/CSS/JS screenshots)
- GCC/G++ (for C/C++ code execution)

## Detailed Usage Instructions

### Using the GUI Application

1. **Launch the application**:
   - From executable: Double-click `Code2word.exe`
   - From source: Run `python c2w_APP.py`

2. **Enter Student Information**:
   - Name
   - Registration Number
   - Course Name
   - Assignment Number

3. **Select Code Files**:
   - Click "Browse Files" to select code files
   - Drag and drop files into the application window
   - Supported languages: Python, Java, C, C++, HTML/CSS/JS

4. **Configure Options**:
   - Include screenshots (recommended)
   - Generate code outputs
   - Custom report title

5. **Generate Report**:
   - Click the "Generate Report" button
   - A progress bar will show processing status
   - Report will be saved as "report_<timestamp>.docx"

### Using Command Line Interface

```bash
python c2w.py [-h] [--title TITLE] [--name NAME] [--reg_no REG_NO] [--course COURSE] [--assignment ASSIGNMENT] [--noscreenshot] <file1> <file2> ...
```

**Required arguments**:
- `<file1> <file2> ...`: Path(s) to code file(s)

**Optional arguments**:
- `-h, --help`: Show help message
- `--title TITLE`: Custom report title
- `--name NAME`: Student name
- `--reg_no REG_NO`: Registration number
- `--course COURSE`: Course name
- `--assignment ASSIGNMENT`: Assignment number
- `--noscreenshot`: Disable screenshots (faster processing)

### Features and File Support

| File Type | Code Display | Output Capture | Screenshot |
|-----------|-------------|---------------|-----------|
| Python (.py) | ✅ | ✅ | ❌ |
| Java (.java) | ✅ | ✅ | ❌ |
| C/C++ (.c/.cpp/.h/.hpp) | ✅ | ✅ | ❌ |
| HTML (.html/.htm) | ✅ | ❌ | ✅ |
| CSS (.css) | ✅ | ❌ | ✅* |
| JavaScript (.js) | ✅ | ✅ | ✅* |

*When used with associated HTML file

## Troubleshooting

### Common Issues

1. **ModuleNotFoundError: No module named 'tkinter'**:
   - Install tkinter: `sudo apt-get install python3-tk` (Ubuntu/Debian)
   - For Windows, reinstall Python with "tcl/tk and IDLE" option checked

2. **Chrome not found**:
   - Set custom Chrome path in settings:
     ```python
     # In c2w.py or environment variable
     BROWSER_BINARY_LOCATION = "C:/path/to/chrome.exe"
     ```

3. **Slow Performance**:
   - Disable screenshots with `--noscreenshot` option
   - Process fewer files at once
   - Close resource-intensive applications

4. **Compiler Not Found Error**:
   - Add C/C++ compiler to PATH
   - For Java: Ensure JDK is installed and in PATH

### Building from Source

To build your own executable:
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico c2w_APP.py
```

## License

MIT License

Copyright (c) 2023 

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files.
