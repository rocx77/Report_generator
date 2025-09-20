import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import os
import c2w
from typing import List, Dict, Tuple, Any


class InputDialog(ctk.CTkToplevel):
    def __init__(self, parent, prompt):
        super().__init__(parent)
        
        self.title("Input Required")
        self.geometry("400x150")
        self.transient(parent)
        self.grab_set()
        
        # Center the dialog on parent window
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (400 // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (150 // 2)
        self.geometry(f"400x150+{x}+{y}")
        
        self.result = None
        self.protocol("WM_DELETE_WINDOW", self.cancel_clicked)  # Handle window close
        
        # Create prompt label
        prompt_label = ctk.CTkLabel(self, text=prompt, wraplength=350)
        prompt_label.pack(pady=10, padx=20)
        
        # Create input entry
        self.entry = ctk.CTkEntry(self, width=350)
        self.entry.pack(pady=10, padx=20)
        self.entry.focus()
        
        # Set a reasonable timeout to prevent hanging
        self.after(30000, self.timeout)  # 30 second timeout
        
        # Create buttons frame
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=10)
        
        ok_button = ctk.CTkButton(button_frame, text="OK", command=self.ok_clicked)
        ok_button.pack(side="left", padx=5)
        
        cancel_button = ctk.CTkButton(button_frame, text="Cancel", command=self.cancel_clicked)
        cancel_button.pack(side="left", padx=5)
        
        # Bind Enter key to OK
        self.bind('<Return>', lambda e: self.ok_clicked())
        self.entry.bind('<Return>', lambda e: self.ok_clicked())
        
    def ok_clicked(self):
        self.result = self.entry.get()
        self.destroy()
        
    def cancel_clicked(self):
        self.result = None
        self.destroy()
        
    def timeout(self):
        """Handle the timeout case to prevent hanging"""
        if self.winfo_exists():
            messagebox.showinfo("Timeout", "Input dialog timed out. Using default value.")
            self.result = ""
            self.destroy()
        self.destroy()


class CodeReportApp(ctk.CTk):
    def __init__(self):
        # Set appearance mode before init for better startup performance
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        super().__init__()

        self.title("Code to Word Report Generator")
        
        # Configure the main grid for responsiveness
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Set an initial size but allow it to be dynamic
        self.geometry("600x750")
        
        # Set initial variables
        self.metadata = {}
        self.files = []
        self.output_path = ""
        
        # Optimize rendering
        self.update_idletasks()

        # Main frame to hold all content, and make it resize with the window
        main_frame = ctk.CTkFrame(self)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Create input fields in a dedicated frame
        input_frame = self.create_input_fields(main_frame)
        input_frame.grid(row=0, column=0, pady=(0, 10), sticky="ew")

        # File selection and reorder section
        file_section_frame = ctk.CTkFrame(main_frame)
        file_section_frame.grid(row=1, column=0, pady=10, sticky="ew")
        file_section_frame.grid_columnconfigure(0, weight=1)

        # File selection button
        ctk.CTkButton(file_section_frame, text="Select Code Files", command=self.select_files).grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # Scrollable frame for reordering files, placed in a new row and made to expand
        self.file_list_frame = ctk.CTkScrollableFrame(main_frame, label_text="Reorder Files")
        self.file_list_frame.grid(row=2, column=0, pady=10, sticky="nsew", rowspan=2)
        self.file_list_frame.grid_columnconfigure(0, weight=1)
        
        # Output folder button
        ctk.CTkButton(main_frame, text="Select Output Folder", command=self.select_output_path).grid(row=4, column=0, pady=10, padx=10, sticky="ew")
        
        # Generate button
        ctk.CTkButton(main_frame, text="Generate Report", command=self.generate_report).grid(row=5, column=0, pady=20, padx=10, sticky="ew")

        self.status_label = ctk.CTkLabel(main_frame, text="")
        self.status_label.grid(row=6, column=0, pady=10, sticky="ew")

    def create_input_fields(self, parent):
        frame = ctk.CTkFrame(parent)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)

        self.entries = {}
        field_map = {
            'Name': 'name',
            'Subject': 'subject',
            'Registration Number': 'reg_no',
            'Group': 'group',
            'Semester': 'semester',
            'Experiment Number': 'experiment_no'
        }
        for i, (label_text, key_name) in enumerate(field_map.items()):
            label = ctk.CTkLabel(frame, text=label_text, anchor="w")
            label.grid(row=i, column=0, padx=10, pady=5, sticky="w")
            
            entry = ctk.CTkEntry(frame)
            entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
            self.entries[key_name] = entry
        
        return frame

    def update_file_list_display(self):
        # Clear existing widgets from the scrollable frame
        for widget in self.file_list_frame.winfo_children():
            widget.destroy()

        if not self.files:
            self.file_list_frame.configure(label_text="No files selected.")
            return

        self.file_list_frame.configure(label_text=f"Reorder {len(self.files)} files:")
        for i, file_path in enumerate(self.files):
            file_name = os.path.basename(file_path)
            
            row_frame = ctk.CTkFrame(self.file_list_frame)
            row_frame.grid(row=i, column=0, pady=2, padx=5, sticky="ew")
            row_frame.grid_columnconfigure(0, weight=1)

            file_label = ctk.CTkLabel(row_frame, text=f"{i+1}. {file_name}", anchor="w")
            file_label.grid(row=0, column=0, padx=5, pady=2, sticky="ew")

            btn_up = ctk.CTkButton(row_frame, text="▲", width=40, command=lambda idx=i: self.move_file_up(idx))
            btn_up.grid(row=0, column=1, padx=2, pady=2)
            if i == 0:
                btn_up.configure(state="disabled")

            btn_down = ctk.CTkButton(row_frame, text="▼", width=40, command=lambda idx=i: self.move_file_down(idx))
            btn_down.grid(row=0, column=2, padx=2, pady=2)
            if i == len(self.files) - 1:
                btn_down.configure(state="disabled")

    def move_file_up(self, index):
        if index > 0:
            self.files[index], self.files[index - 1] = self.files[index - 1], self.files[index]
            self.update_file_list_display()

    def move_file_down(self, index):
        if index < len(self.files) - 1:
            self.files[index], self.files[index + 1] = self.files[index + 1], self.files[index]
            self.update_file_list_display()

    def select_files(self):
        selected_files = filedialog.askopenfilenames(
            title="Select Code Files",
            filetypes=[
                ("All Supported Files", "*.py *.c *.cpp *.java *.html *.css *.js *.php"),
                ("Python Files", "*.py"),
                ("C/C++ Files", "*.c *.cpp"),
                ("Java Files", "*.java"),
                ("HTML/CSS Files", "*.html *.css"),
                ("JavaScript Files", "*.js"),
                ("PHP Files", "*.php"),
                ("All Files", "*.*")
            ]
        )
        if selected_files:
            self.files = list(selected_files)
            self.status_label.configure(text=f"Selected {len(self.files)} files. Reorder below.")
            self.update_file_list_display()

    def select_output_path(self):
        self.output_path = filedialog.askdirectory(title="Select Output Folder")
        self.status_label.configure(text=f"Output path set to: {self.output_path}")

    def gui_input_callback(self, file_name=None, prompts=None):
        """
        GUI callback function for collecting user input during code execution.
        This method is called when interactive programs need user input.
        
        Args:
            file_name (str): The name of the file being executed (optional)
            prompts (list): A list of prompts to show to the user (optional)
        """
        if prompts and isinstance(prompts, list):
            # Handle multiple prompts for C/C++ files
            results = []
            for prompt in prompts:
                dialog = InputDialog(self, prompt)
                self.wait_window(dialog)
                if dialog.result is None:  # User canceled
                    return None
                results.append(dialog.result)
            return results
        else:
            # Handle single prompt for Python/other files
            dialog = InputDialog(self, prompts if prompts else "Enter input:")
            self.wait_window(dialog)
            return dialog.result

    def generate_report(self):
        if not self.files or not self.output_path:
            messagebox.showerror("Error", "Please select files and output folder.")
            return

        metadata = {key: entry.get() for key, entry in self.entries.items()}
        
        # Create a progress indicator
        progress_window = ctk.CTkToplevel(self)
        progress_window.title("Generating Report")
        progress_window.geometry("400x150")
        progress_window.transient(self)
        progress_window.grab_set()
        
        # Center the progress window
        progress_window.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (400 // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (150 // 2)
        progress_window.geometry(f"400x150+{x}+{y}")
        
        # Add a label
        progress_label = ctk.CTkLabel(progress_window, text="Generating report...\nPlease wait, this may take a few moments.", font=("Arial", 14))
        progress_label.pack(pady=20)
        
        # Add a progress bar
        progress_bar = ctk.CTkProgressBar(progress_window, width=300)
        progress_bar.pack(pady=10)
        progress_bar.start()
        
        # Use a background thread for processing to keep GUI responsive
        import threading
        
        def process_report():
            try:
                output_file = c2w.create_word_doc(self.files, metadata, self.output_path, gui_input_callback=self.gui_input_callback)
                
                # Update UI on the main thread
                self.after(100, lambda: self.report_complete(progress_window, output_file))
            except Exception as e:
                # Handle error on the main thread
                self.after(100, lambda: self.report_error(progress_window, str(e)))
        
        # Start processing in background
        threading.Thread(target=process_report, daemon=True).start()
    
    def report_complete(self, progress_window, output_file):
        """Called when report generation is complete"""
        progress_window.destroy()
        messagebox.showinfo("Success", f"Report generated successfully!\nSaved to: {output_file}")
        self.status_label.configure(text="✅ Report generated successfully.")
        
        # Offer to open the file
        if messagebox.askyesno("Open File", "Would you like to open the generated report?"):
            import subprocess
            import platform
            
            if platform.system() == 'Windows':
                subprocess.Popen(['start', '', output_file], shell=True)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.Popen(['open', output_file])
            else:  # Linux
                subprocess.Popen(['xdg-open', output_file])
    
    def report_error(self, progress_window, error_message):
        """Called when report generation encounters an error"""
        progress_window.destroy()
        messagebox.showerror("Error", f"Failed to generate report:\n{error_message}")
        self.status_label.configure(text="❌ Error generating report.")


if __name__ == '__main__':
    # Optimize startup
    import sys
    import os
    
    # Disable error reporting for faster startup
    if hasattr(sys, 'frozen'):
        # We're running in a PyInstaller bundle
        os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
    
    # Disable unnecessary debug logging
    import logging
    logging.basicConfig(level=logging.ERROR)
    
    # Initialize the app
    app = CodeReportApp()
    
    # Schedule a garbage collection after startup
    def post_startup():
        import gc
        gc.collect()
    app.after(1000, post_startup)
    
    # Start the app
    app.mainloop()