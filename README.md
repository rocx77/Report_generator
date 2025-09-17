# code2word

A Python tool to generate Word document reports from code files (Python, C, C++, JS, HTML, CSS, PHP) with automatic output capture and frontend screenshots for web files.

## Features
- Converts code files into a formatted Word report
- Captures code output/errors for supported languages
- Takes full-page screenshots of HTML/CSS/JS files using Selenium
- Adds student/assignment metadata

## Requirements
- Python 3.10+
- **Google Chrome browser must be installed** (required for HTML/CSS/JS screenshots and proper functioning)
- ChromeDriver (auto-managed by webdriver_manager)

### Python Packages
Install dependencies with:

```bash
pip install -r requirements.txt
```

## Usage

1. Place your code files in the project directory.
2. Run the script:

```bash
python c2w.py <file1> <file2> ...
```

3. Enter assignment metadata when prompted.
4. The Word report will be saved in the current directory.

### Example
```bash
python c2w.py example.py example.html
```

## Notes
- For HTML/CSS/JS files, Chrome must be installed and accessible.
- Screenshots are taken in headless mode; adjust `VIEWPORT_HEIGHT` in `c2w.py` if needed.
- For C/C++ files, GCC/G++ must be installed and available in PATH.

## Project Structure
- `c2w.py`: Main script for report generation
- `c2w_APP.py`: (Optional) Application entry or GUI (if present)
- `c2wenv/`: Python virtual environment (recommended)

## License
MIT License
