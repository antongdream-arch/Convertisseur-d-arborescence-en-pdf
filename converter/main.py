import threading
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox
from config import *
from converters import txt_converter, docx_converter, xlsx_converter

ctk.set_appearance_mode(THEME_MODE)
ctk.set_default_color_theme(THEME_COLOR)


class ConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(APP_DIMENSIONS)

        try:
            self.state('zoomed')
        except Exception:
            pass

        self.input_folder = None

        self.title_label = ctk.CTkLabel(self, text="Universal PDF Converter", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(pady=(20, 10), fill="x")

        self.selection_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.selection_frame.pack(pady=10, padx=20, fill="x")

        self.path_entry = ctk.CTkEntry(self.selection_frame, placeholder_text="Select a folder...")
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.browse_button = ctk.CTkButton(self.selection_frame, text="Browse", command=self.choose_folder)
        self.browse_button.pack(side="right")

        self.convert_button = ctk.CTkButton(
            self, text="Start Conversion", font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=BUTTON_COLOR, hover_color=HOVER_COLOR, height=40, state="disabled",
            command=self.start_conversion
        )
        self.convert_button.pack(pady=15, padx=100, fill="x")

        self.console_label = ctk.CTkLabel(self, text="Operation Log:", font=ctk.CTkFont(weight="bold"))
        self.console_label.pack(anchor="w", padx=20)

        self.console = ctk.CTkTextbox(self, height=100, state="disabled")
        self.console.pack(pady=(5, 20), padx=20, fill="both", expand=True)

    def log(self, message):
        self.console.configure(state="normal")
        self.console.insert("end", message + "\n")
        self.console.see("end")
        self.console.configure(state="disabled")

    def choose_folder(self):
        folder = filedialog.askdirectory(title="Select source folder")
        if folder:
            self.input_folder = Path(folder)
            self.path_entry.delete(0, "end")
            self.path_entry.insert(0, str(self.input_folder))
            self.convert_button.configure(state="normal")
            self.log(f"Folder selected: {self.input_folder}")

    def start_conversion(self):
        if not self.input_folder or not self.input_folder.exists():
            messagebox.showerror("Error", "Please select a valid folder.")
            return

        self.convert_button.configure(state="disabled")
        self.browse_button.configure(state="disabled")
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.configure(state="disabled")

        thread = threading.Thread(target=self.conversion_process)
        thread.start()

    def conversion_process(self):
        try:
            output_folder = self.input_folder.with_name(self.input_folder.name + "_PDF")
            self.log("--- CONVERSION STARTED ---")

            files_found = 0

            for file in self.input_folder.rglob("*"):
                if file.is_file() and file.stat().st_size > 0:
                    if file.suffix.lower() == ".txt":
                        txt_converter(file, self.input_folder, output_folder, self.log)
                        files_found += 1
                    elif file.suffix.lower() == ".docx":
                        docx_converter(file, self.input_folder, output_folder, self.log)
                        files_found += 1
                    elif file.suffix.lower() == ".xlsx":
                        xlsx_converter(file, self.input_folder, output_folder, self.log)
                        files_found += 1

            if files_found == 0:
                self.log("No compatible files found.")
            else:
                self.log(f"\n--- DONE! {files_found} file(s) processed. ---")

        except Exception as e:
            self.log(f"An error occurred: {e}")

        finally:
            self.convert_button.configure(state="normal")
            self.browse_button.configure(state="normal")


if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()