#!/usr/bin/env python3
"""
Test script to validate the interactive input system integration
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import c2w

def test_gui_callback(prompt):
    """Test GUI callback that simulates user input"""
    print(f"GUI would show dialog with prompt: {prompt}")
    # Return simulated user input based on the prompt
    if "name" in prompt.lower():
        return "John Doe"
    elif "age" in prompt.lower():
        return "25"
    elif "first number" in prompt.lower():
        return "10"
    elif "second number" in prompt.lower():
        return "20"
    else:
        return "test input"

def test_interactive_system():
    """Test the interactive input system"""
    print("Testing interactive input system...")
    
    # Test with Python file
    test_file = "test_interactive.py"
    if os.path.exists(test_file):
        print(f"\nTesting with {test_file}:")
        output, error, image = c2w.run_code(test_file, test_gui_callback)
        print("Output:", output)
        if error:
            print("Error:", error)
    
    # Test document generation
    metadata = {
        'name': 'Test User',
        'subject': 'Computer Science',
        'reg_no': 'CS123',
        'group': 'A',
        'semester': '3',
        'experiment_no': '1'
    }
    
    files = ["test_interactive.py"]
    if os.path.exists("test_interactive.c"):
        files.append("test_interactive.c")
    
    print(f"\nGenerating Word document with files: {files}")
    try:
        c2w.create_word_doc(files, metadata, ".", test_gui_callback)
        print("✅ Document generation successful!")
    except Exception as e:
        print(f"❌ Document generation failed: {e}")

if __name__ == "__main__":
    test_interactive_system()