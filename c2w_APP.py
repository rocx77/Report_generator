import customtkinter as ctk  # type: ignore
from tkinter import filedialog, messagebox
import os
import c2w


class CodeReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Code to Word Report Generator")
        self.geometry("600x950+300+20")    # you can make your window adjustments here

        ctk.set_appearance_mode("dark")

        self.metadata = {}
        self.files = []
        self.output_path = ""

        # Main frame to hold everything
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Metadata inputs
        self.create_input_fields(main_frame)

        # File selection buttons
        file_frame = ctk.CTkFrame(main_frame)
        file_frame.pack(pady=10, padx=10, fill="x")
        ctk.CTkButton(file_frame, text="Select Code Files", command=self.select_files).pack(pady=10, padx=10)

        # Scrollable frame for reordering files
        self.file_list_frame = ctk.CTkScrollableFrame(main_frame, label_text="Reorder Files (Drag Up/Down)")
        self.file_list_frame.pack(pady=10, padx=10, fill="both", expand=True)

        # Output folder button
        ctk.CTkButton(main_frame, text="Select Output Folder", command=self.select_output_path).pack(pady=10, padx=10)

        # Generate button
        ctk.CTkButton(main_frame, text="Generate Report", command=self.generate_report).pack(pady=20, padx=10)

        self.status_label = ctk.CTkLabel(main_frame, text="")
        self.status_label.pack(pady=10)

    def create_input_fields(self, parent):
        frame = ctk.CTkFrame(parent)
        frame.pack(pady=20, padx=20, fill="x")

        self.entries = {}
        field_map = {
            'Name': 'name',
            'Subject': 'subject',
            'Registration Number': 'reg_no',
            'Group': 'group',
            'Semester': 'semester',
            'Experiment Number': 'experiment_no'
        }
        for label_text, key_name in field_map.items():
            label = ctk.CTkLabel(frame, text=label_text)
            label.pack(anchor="w", padx=10)
            entry = ctk.CTkEntry(frame, width=500)
            entry.pack(pady=5, padx=10)
            self.entries[key_name] = entry

    def update_file_list_display(self):
        # Clear existing widgets from the scrollable frame
        for widget in self.file_list_frame.winfo_children():
            widget.destroy()

        # Populate the frame with files and buttons
        if not self.files:
            self.file_list_frame.configure(label_text="No files selected.")
            return

        self.file_list_frame.configure(label_text=f"Reorder {len(self.files)} files:")
        for i, file_path in enumerate(self.files):
            file_name = os.path.basename(file_path)
            
            # Create a sub-frame for each file row
            row_frame = ctk.CTkFrame(self.file_list_frame)
            row_frame.pack(pady=2, padx=5, fill="x")

            # File Name Label
            file_label = ctk.CTkLabel(row_frame, text=f"{i+1}. {file_name}", anchor="w")
            file_label.pack(side="left", fill="x", expand=True, padx=5)

            # Down Button
            btn_down = ctk.CTkButton(row_frame, text="▼", width=40, command=lambda idx=i: self.move_file_down(idx))
            btn_down.pack(side="right", padx=2)
            if i == len(self.files) - 1:
                btn_down.configure(state="disabled")

            # Up Button
            btn_up = ctk.CTkButton(row_frame, text="▲", width=40, command=lambda idx=i: self.move_file_up(idx))
            btn_up.pack(side="right", padx=2)
            if i == 0:
                btn_up.configure(state="disabled")

    def move_file_up(self, index):
        if index > 0:
            self.files[index], self.files[index - 1] = self.files[index - 1], self.files[index]
            self.update_file_list_display()

    def move_file_down(self, index):
        if index < len(self.files) - 1:
            self.files[index], self.files[index + 1] = self.files[index + 1], self.files[index]
            self.update_file_list_display()

    def select_files(self):
        """
        Handles file selection via a button click.
        """
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
        """
        Handles output folder selection.
        """
        self.output_path = filedialog.askdirectory(title="Select Output Folder")
        self.status_label.configure(text=f"Output path set to: {self.output_path}")

    def generate_report(self):
        """
        Generates the final report.
        """
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