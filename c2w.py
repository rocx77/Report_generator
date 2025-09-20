import subprocess
import sys
import os
import time
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from typing import List, Dict, Tuple, Any
import re
import tempfile


# --- Configuration for Frontend Screenshots ---
BROWSER_BINARY_LOCATION = None
VIEWPORT_HEIGHT = 1394


def run_code(file_path: str) -> Tuple[str | None, str | None, str | None]:
    """
    Executes a given code file, handles various output types (text, images),
    and captures its output and errors.
    Returns a tuple of (stdout, stderr, image_path).
    """
    ext = os.path.splitext(file_path)[1].lower()
    abs_file_path = os.path.abspath(file_path)
    temp_image_path = None
    exe_file = None
    command = []

    if ext == '.py':
        with open(abs_file_path, 'r', encoding='utf-8') as f:
            code_content = f.read()

        if re.search(r'import\s+(matplotlib\.pyplot|seaborn)', code_content):
            temp_dir = tempfile.gettempdir()
            temp_file_path = os.path.join(temp_dir, "temp_plot_script.py")
            temp_image_path = os.path.join(temp_dir, "matplotlib_output.png")
            
            # Use forward slashes for the path to avoid SyntaxError on Windows
            save_path = temp_image_path.replace('\\', '/')

            injected_code = (
                "import matplotlib\n"
                "matplotlib.use('Agg')\n"
                "import matplotlib.pyplot as plt\n"
                + code_content.replace("plt.show()", "") + "\n"
                + f"plt.savefig('{save_path}')\n"
            )
            
            with open(temp_file_path, "w", encoding='utf-8') as temp_f:
                temp_f.write(injected_code)
            
            command = ['python', temp_file_path]
            result = subprocess.run(command, capture_output=True, text=True)
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
            return result.stdout, result.stderr, temp_image_path
        else:
            command = ['python', abs_file_path]

    elif ext == '.c':
        exe_file = os.path.join(os.getcwd(), 'temp_program.exe')
        subprocess.run(['gcc', abs_file_path, '-o', exe_file], check=True)
        command = [exe_file]
    elif ext == '.cpp':
        exe_file = os.path.join(os.getcwd(), 'temp_program.exe')
        subprocess.run(['g++', abs_file_path, '-o', exe_file], check=True)
        command = [exe_file]
    elif ext == '.java':
        class_file = os.path.splitext(os.path.basename(abs_file_path))[0]
        subprocess.run(['javac', abs_file_path], check=True)
        command = ['java', '-cp', os.path.dirname(abs_file_path), class_file]
    elif ext == '.js':
        command = ['node', abs_file_path]
    elif ext == '.php':
        command = ['php', abs_file_path]
    elif ext in ['.html', '.css']:
        return None, None, None
    else:
        raise Exception(f"Unsupported file type: {ext}")

    result = subprocess.run(command, capture_output=True, text=True)

    if exe_file and ext in ['.c', '.cpp'] and os.path.exists(exe_file):
        os.remove(exe_file)

    return result.stdout, result.stderr, None


def generate_full_page_screenshots(file_path: str, viewport_height: int = 1394) -> List[str]:
    """
    Generates multiple screenshots of a webpage by scrolling,
    ensuring they do not overlap, and returns a list of screenshot paths.
    """
    options = webdriver.ChromeOptions()
    if BROWSER_BINARY_LOCATION:
        options.binary_location = BROWSER_BINARY_LOCATION

    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-gpu')

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    abs_path = os.path.abspath(file_path)
    file_url = 'file:///' + abs_path.replace('\\', '/')

    driver.set_window_size(1920, viewport_height)
    driver.get(file_url)
    time.sleep(2)

    screenshot_paths = []
    page_height = driver.execute_script("return document.body.scrollHeight")
    current_scroll_position = 0
    screenshot_counter = 0

    while current_scroll_position < page_height:
        # Save temporary screenshot to the system's temp directory
        temp_dir = tempfile.gettempdir()
        screenshot_name = os.path.join(temp_dir, f"{os.path.splitext(os.path.basename(file_path))[0]}_screenshot_{screenshot_counter:02d}.png")
        
        driver.save_screenshot(screenshot_name)
        screenshot_paths.append(screenshot_name)
        screenshot_counter += 1

        current_scroll_position += viewport_height
        driver.execute_script(f"window.scrollTo(0, {current_scroll_position});")
        time.sleep(0.5)

        page_height = driver.execute_script("return document.body.scrollHeight")
        
        if current_scroll_position >= page_height and screenshot_counter > 1:
            break
        elif current_scroll_position + viewport_height > page_height and current_scroll_position < page_height:
            driver.execute_script(f"window.scrollTo(0, {page_height});")
            time.sleep(0.5)
            screenshot_name = os.path.join(temp_dir, f"{os.path.splitext(os.path.basename(file_path))[0]}_screenshot_{screenshot_counter:02d}.png")
            driver.save_screenshot(screenshot_name)
            screenshot_paths.append(screenshot_name)
            break

    driver.quit()
    return screenshot_paths


def add_code_block(doc: Any, code_text: str):
    """
    Adds a formatted code block to the Word document.
    """
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    cell = table.cell(0, 0)

    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), 'D9D9D9')
    cell._tc.get_or_add_tcPr().append(shading_elm)

    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(code_text)
    run.font.name = 'Courier New'
    run.font.size = Pt(10.5)


def set_fixed_cell_width(cell, width_in_inches: float):
    """
    Sets a fixed width for a table cell.
    """
    cell.width = Inches(width_in_inches)
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:type'), 'dxa')
    tcW.set(qn('w:w'), str(int(width_in_inches * 1440)))
    tcPr.append(tcW)


def create_student_info_table(doc: Any, metadata: Dict[str, str]):
    """
    Creates and populates the student information table.
    """
    table = doc.add_table(rows=4, cols=2)
    table.style = 'Table Grid'

    for row in table.rows:
        set_fixed_cell_width(row.cells[0], 1.5)
        set_fixed_cell_width(row.cells[1], 1.5)

    table.cell(0, 0).text = 'Name'
    table.cell(0, 1).text = metadata['name']
    table.cell(1, 0).text = 'Registration Number'
    table.cell(1, 1).text = metadata['reg_no']
    table.cell(2, 0).text = 'Semester'
    table.cell(2, 1).text = metadata['semester']
    table.cell(3, 0).text = 'Group'
    table.cell(3, 1).text = metadata['group']

    table.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT


def create_word_doc(files: List[str], metadata: Dict[str, str], output_dir: str = "."):
    """
    Main function to generate the Word document report.
    """
    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    files = [os.path.abspath(f) for f in files]
    output_dir = os.path.abspath(output_dir)

    doc = Document()
    section = doc.sections[0]

    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    heading = doc.add_heading(metadata['subject'] + f" Experiment {metadata['experiment_no']}", level=1)
    heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = heading.runs[0]
    run.font.size = Pt(28)
    run.font.bold = True

    for i, file_path in enumerate(files):
        if not os.path.isfile(file_path):
            print(f"Skipping: {file_path} (file not found)")
            continue

        ext = os.path.splitext(file_path)[1].lower()

        with open(file_path, 'r', encoding='utf-8') as f:
            code_text = f.read()

        if i > 0:
            doc.add_page_break()

        doc.add_heading(f'File: {os.path.basename(file_path)}', level=2)
        doc.add_heading('Source Code:', level=3)
        add_code_block(doc, code_text)

        if ext in ['.html', '.css', '.js']:
            try:
                screenshot_paths = generate_full_page_screenshots(file_path)
                doc.add_heading('Frontend Preview:', level=3)
                if screenshot_paths:
                    for s_path in screenshot_paths:
                        doc.add_picture(s_path, width=Inches(6))
                        os.remove(s_path)
                else:
                    doc.add_paragraph("[No screenshots generated]")
            except Exception as e:
                doc.add_heading('Frontend Error:', level=3)
                doc.add_paragraph(str(e))
        else:
            try:
                output_text, error_text, image_path = run_code(file_path)
                
                doc.add_heading('Output:', level=3)
                doc.add_paragraph(output_text or "[No Output]")

                if image_path and os.path.exists(image_path):
                    doc.add_heading('Image Output:', level=3)
                    doc.add_picture(image_path, width=Inches(6))
                    os.remove(image_path)
                
                if error_text:
                    doc.add_heading('Errors:', level=3)
                    doc.add_paragraph(error_text)
            except Exception as e:
                doc.add_heading('Execution Error:', level=3)
                doc.add_paragraph(str(e))

    doc.add_heading('Submitted By:-', level=2)
    create_student_info_table(doc, metadata)

    name_prefix = metadata['name'].split()[0].replace(' ', '_')
    output_filename = f"{name_prefix}_{metadata['subject']}_{metadata['experiment_no']}.docx"
    output_docx = os.path.join(output_dir, output_filename)
    
    doc.save(output_docx)
    print(f"\n✅ Final report saved as {output_docx}")


def main():

    print("-" * 40)
    print("📝 Enter Assignment Info:")
    metadata = {
        'name': input("Enter your name: "),
        'subject': input("Enter subject name: "),
        'reg_no': input("Enter registration number: "),
        'group': input("Enter group: "),
        'semester': input("Enter semester: "),
        'experiment_no': input("Enter experiment number: ")
    }

    if len(sys.argv) < 2:
        print("\nUsage: python code_to_word_report.py <file1> <file2> ...")
        return

    files = [os.path.abspath(f) for f in sys.argv[1:]]
    create_word_doc(files, metadata)


if __name__ == '__main__':
    main()