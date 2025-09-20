# Build script for creating executable
# Run this with: python build_exe.py

import PyInstaller.__main__
import os

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

PyInstaller.__main__.run([
    '--name=Code2word',
    '--onefile',
    '--windowed',  # No console window
    '--add-data=c2w.py;.',  # Include the c2w module
    '--hidden-import=customtkinter',
    '--hidden-import=selenium',
    '--hidden-import=webdriver_manager',
    '--hidden-import=docx',
    '--hidden-import=PIL',
    '--hidden-import=lxml',
    '--distpath=dist',
    '--workpath=build',
    '--specpath=.',
    'c2w_APP.py'
])

print("\n✅ Executable created successfully!")
print("📁 Location: dist/Code2word.exe")
print("🚀 You can now distribute this single .exe file!")