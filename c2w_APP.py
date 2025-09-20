import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import c2w
from typing import List, Dict, Tuple, Any


class CodeReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Code to Word Report Generator")
        
        # Configure the main grid for responsiveness
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Set an initial size but allow it to be dynamic
        self.geometry("600x700")

        ctk.set_appearance_mode("dark")

        self.metadata = {}
        self.files = []
        self.output_path = ""

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
                ("All Supported Files", "*.py *.c *.cpp *.html *.css *.js *.php"),
                ("Python Files", "*.py"),
                ("C/C++ Files", "*.c *.cpp"),
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

    def generate_report(self):
        if not self.files or not self.output_path:
            messagebox.showerror("Error", "Please select files and output folder.")
            return

        metadata = {key: entry.get() for key, entry in self.entries.items()}

        try:
            c2w.create_word_doc(self.files, metadata, self.output_path)
            messagebox.showinfo("Success", "Report generated successfully!")
            self.status_label.configure(text="✅ Report generated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report:\n{str(e)}")


if __name__ == '__main__':
    app = CodeReportApp()
    app.mainloop()