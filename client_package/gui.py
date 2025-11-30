import tkinter as tk
from tkinter import filedialog, messagebox
import os
import threading
from main import convert_excel_to_xml

class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to XML Converter")
        self.root.geometry("500x250")
        self.root.resizable(False, False)

        # Title
        title_label = tk.Label(root, text="Excel to XML Converter", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=20)

        # File Selection
        self.file_frame = tk.Frame(root)
        self.file_frame.pack(pady=10)

        self.file_path_var = tk.StringVar()
        self.entry = tk.Entry(self.file_frame, textvariable=self.file_path_var, width=40, state='readonly')
        self.entry.pack(side=tk.LEFT, padx=5)

        self.browse_btn = tk.Button(self.file_frame, text="Select Excel File", command=self.browse_file)
        self.browse_btn.pack(side=tk.LEFT, padx=5)

        # Convert Button
        self.convert_btn = tk.Button(root, text="Convert to XML", command=self.start_conversion, bg="#4CAF50", fg="white", font=("Helvetica", 12, "bold"))
        self.convert_btn.pack(pady=20)

        # Status
        self.status_label = tk.Label(root, text="", fg="gray")
        self.status_label.pack(pady=5)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filename:
            self.file_path_var.set(filename)
            self.status_label.config(text="File selected", fg="blue")

    def start_conversion(self):
        input_path = self.file_path_var.get()
        if not input_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return

        output_path = os.path.splitext(input_path)[0] + ".xml"
        
        self.convert_btn.config(state=tk.DISABLED, text="Converting...")
        self.status_label.config(text="Converting...", fg="orange")

        # Run in separate thread to avoid freezing GUI
        threading.Thread(target=self.run_conversion, args=(input_path, output_path)).start()

    def run_conversion(self, input_path, output_path):
        try:
            convert_excel_to_xml(input_path, output_path)
            self.root.after(0, lambda: self.conversion_success(output_path))
        except Exception as e:
            self.root.after(0, lambda: self.conversion_error(str(e)))

    def conversion_success(self, output_path):
        self.convert_btn.config(state=tk.NORMAL, text="Convert to XML")
        self.status_label.config(text=f"Success! Saved to {os.path.basename(output_path)}", fg="green")
        messagebox.showinfo("Success", f"Conversion completed successfully!\n\nFile saved at:\n{output_path}")

    def conversion_error(self, error_msg):
        self.convert_btn.config(state=tk.NORMAL, text="Convert to XML")
        self.status_label.config(text="Error occurred", fg="red")
        messagebox.showerror("Error", f"An error occurred during conversion:\n{error_msg}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()
