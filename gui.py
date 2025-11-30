import tkinter as tk
from tkinter import filedialog, messagebox
import os
import threading
from main import process_batch

class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to XML Converter")
        self.root.geometry("600x350")
        self.root.resizable(False, False)

        # Title
        title_label = tk.Label(root, text="Excel to XML Converter", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=15)

        # File Selection
        self.file_frame = tk.Frame(root)
        self.file_frame.pack(pady=5, padx=20, fill=tk.X)
        
        tk.Label(self.file_frame, text="Input Excel Files:").pack(anchor=tk.W)
        
        self.file_paths_var = tk.StringVar()
        self.file_entry = tk.Entry(self.file_frame, textvariable=self.file_paths_var, state='readonly')
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.browse_files_btn = tk.Button(self.file_frame, text="Select Files", command=self.browse_files)
        self.browse_files_btn.pack(side=tk.LEFT)

        # Output Directory Selection
        self.output_frame = tk.Frame(root)
        self.output_frame.pack(pady=5, padx=20, fill=tk.X)
        
        tk.Label(self.output_frame, text="Output Folder (Optional):").pack(anchor=tk.W)
        
        self.output_dir_var = tk.StringVar()
        self.output_entry = tk.Entry(self.output_frame, textvariable=self.output_dir_var, state='readonly')
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.browse_output_btn = tk.Button(self.output_frame, text="Select Folder", command=self.browse_output_folder)
        self.browse_output_btn.pack(side=tk.LEFT)

        # Convert Button
        self.convert_btn = tk.Button(root, text="Convert All to XML", command=self.start_conversion, bg="#4CAF50", fg="white", font=("Helvetica", 12, "bold"))
        self.convert_btn.pack(pady=20)

        # Status
        self.status_label = tk.Label(root, text="", fg="gray")
        self.status_label.pack(pady=5)

    def browse_files(self):
        filenames = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filenames:
            self.file_paths_var.set("; ".join(filenames))
            count = len(filenames)
            self.status_label.config(text=f"{count} file(s) selected", fg="blue")

    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir_var.set(folder)

    def start_conversion(self):
        file_paths_str = self.file_paths_var.get()
        if not file_paths_str:
            messagebox.showwarning("Warning", "Please select at least one Excel file.")
            return

        file_paths = file_paths_str.split("; ")
        output_dir = self.output_dir_var.get() or None
        
        self.convert_btn.config(state=tk.DISABLED, text="Converting...")
        self.status_label.config(text="Processing...", fg="orange")

        # Run in separate thread to avoid freezing GUI
        threading.Thread(target=self.run_conversion, args=(file_paths, output_dir)).start()

    def run_conversion(self, file_paths, output_dir):
        try:
            results = process_batch(file_paths, output_dir)
            self.root.after(0, lambda: self.conversion_finished(results))
        except Exception as e:
            self.root.after(0, lambda: self.conversion_error(str(e)))

    def conversion_finished(self, results):
        self.convert_btn.config(state=tk.NORMAL, text="Convert All to XML")
        
        success_count = len(results["success"])
        failed_count = len(results["failed"])
        
        if failed_count == 0:
            self.status_label.config(text=f"Success! Converted {success_count} files.", fg="green")
            messagebox.showinfo("Success", f"Successfully converted {success_count} files!")
        else:
            self.status_label.config(text=f"Completed with errors. Success: {success_count}, Failed: {failed_count}", fg="orange")
            error_msg = "\n".join([f"{os.path.basename(f)}: {e}" for f, e in results["failed"]])
            messagebox.showwarning("Partial Success", f"Converted: {success_count}\nFailed: {failed_count}\n\nErrors:\n{error_msg}")

    def conversion_error(self, error_msg):
        self.convert_btn.config(state=tk.NORMAL, text="Convert All to XML")
        self.status_label.config(text="Error occurred", fg="red")
        messagebox.showerror("Error", f"An unexpected error occurred:\n{error_msg}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()
