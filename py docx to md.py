import pypandoc
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class DocxToMdConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Word to Markdown Converter")
        self.root.geometry("500x300")
        self.root.resizable(False, False)

        # Main Container
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(expand=True, fill="both")

        # Title Label
        self.title_label = ttk.Label(
            self.main_frame, 
            text="DOCX → Markdown", 
            font=("Helvetica", 16, "bold")
        )
        self.title_label.pack(pady=(0, 20))

        # Instructions
        self.inst_label = ttk.Label(
            self.main_frame, 
            text="Choose whether to convert a single file or a folder of files."
        )
        self.inst_label.pack(pady=(0, 20))

        # Button Container
        self.btn_frame = ttk.Frame(self.main_frame)
        self.btn_frame.pack(pady=10)

        # File Button
        self.file_btn = ttk.Button(
            self.btn_frame, 
            text="Convert Single File", 
            command=self.convert_single_file
        )
        self.file_btn.grid(row=0, column=0, padx=10)

        # Folder Button
        self.folder_btn = ttk.Button(
            self.btn_frame, 
            text="Convert Entire Folder", 
            command=self.convert_folder
        )
        self.folder_btn.grid(row=0, column=1, padx=10)

        # Status Label
        self.status_label = ttk.Label(
            self.main_frame, 
            text="Ready", 
            foreground="gray"
        )
        self.status_label.pack(pady=(20, 0))

    def update_status(self, text, color="black"):
        self.status_label.config(text=text, foreground=color)
        self.root.update_idletasks()

    def perform_conversion(self, source_path):
        """Helper to handle the actual pypandoc call"""
        dest_path = os.path.splitext(source_path)[0] + '.md'
        pypandoc.convert_file(source_path, 'gfm', outputfile=dest_path)

    def convert_single_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word files", "*.docx")]
        )
        
        if file_path:
            try:
                self.update_status("Converting...", "blue")
                self.perform_conversion(file_path)
                self.update_status("Success! File converted.", "green")
                messagebox.showinfo("Done", f"Converted:\n{os.path.basename(file_path)}")
            except Exception as e:
                self.update_status("Error occurred", "red")
                messagebox.showerror("Error", f"Failed to convert:\n{str(e)}")

    def convert_folder(self):
        folder_path = filedialog.askdirectory(title="Select Folder containing .docx files")
        
        if folder_path:
            files = [f for f in os.listdir(folder_path) 
                     if f.endswith(".docx") and not f.startswith("~$")]
            
            if not files:
                messagebox.showwarning("No Files", "No .docx files found in this folder.")
                return

            try:
                count = 0
                for filename in files:
                    full_path = os.path.join(folder_path, filename)
                    self.update_status(f"Converting {filename}...", "blue")
                    self.perform_conversion(full_path)
                    count += 1
                
                self.update_status(f"Finished! {count} files converted.", "green")
                messagebox.showinfo("Done", f"Successfully converted {count} files!")
            except Exception as e:
                self.update_status("Error occurred", "red")
                messagebox.showerror("Error", f"An error occurred during batch conversion:\n{str(e)}")

if __name__ == "__main__":
    # Ensure Pandoc is installed
    try:
        # This just checks if pypandoc can find the pandoc binary
        pypandoc.get_pandoc_version()
    except OSError:
        # If Pandoc isn't installed on the system, we warn the user immediately
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Pandoc Missing", 
                             "Pandoc is not installed on this system.\n\nPlease install Pandoc from https://pandoc.org/installing.html")
        exit()

    root = tk.Tk()
    app = DocxToMdConverter(root)
    root.mainloop()