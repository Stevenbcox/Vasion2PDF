import os
import threading
import tkinter as tk
from ttkthemes import ThemedTk
from tkinter import filedialog, ttk, Menu
from tkinter.filedialog import askopenfilename
from main import main


class GUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Vasion 2 PDF")
        self.is_running = True

        # Create a menu bar
        menu_bar = Menu(self.master)
        self.master.config(menu=menu_bar)

        # Create labels and entry widgets
        self.input_file_label = tk.Label(master, text="Input File:")
        self.input_file_label.grid(row=0, column=0, sticky=tk.E, padx=5, pady=(5, 0))

        self.input_file_var = tk.StringVar()
        self.input_file_entry = ttk.Entry(master, textvariable=self.input_file_var, width=30)
        self.input_file_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=(5, 0))

        self.browse_input_button = ttk.Button(master, text="Browse", command=self.browse_input)
        self.browse_input_button.grid(row=0, column=2, padx=10, pady=(5, 0), sticky=tk.W)

        self.output_folder_label = tk.Label(master, text="Output Folder:")
        self.output_folder_label.grid(row=1, column=0, sticky=tk.E, padx=5, pady=5)

        self.output_folder_var = tk.StringVar()
        self.output_folder_entry = ttk.Entry(master, textvariable=self.output_folder_var, width=30)
        self.output_folder_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        self.browse_output_button = ttk.Button(master, text="Browse", command=self.browse_output)
        self.browse_output_button.grid(row=1, column=2, padx=10, pady=5, sticky=tk.W)

        # Create label for processing status
        self.processing_status_var = tk.StringVar()
        self.processing_status_label = tk.Label(master, textvariable=self.processing_status_var, fg="grey")
        self.processing_status_label.grid(row=2, columnspan=2, padx=10, pady=10)

        # Create submit button
        self.generate_button = ttk.Button(master, text="Submit", command=self.generate_excel)
        self.generate_button.grid(row=2, columnspan=3, padx=10, pady=10, sticky=tk.E)

        # Bind the close event to set the is_running flag to False
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)

    def set_processing_status(self, status):
        self.processing_status_var.set(status)

    def browse_input(self):
        file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.input_file_var.set(file_path)

    def browse_output(self):
        folder_path = filedialog.askdirectory()
        self.output_folder_var.set(folder_path)

    def open_file_explorer(self, folder_path):
        os.startfile(folder_path)

    def generate_excel(self):
        input_file = self.input_file_var.get()
        output_folder = self.output_folder_var.get()

        if not input_file or not output_folder:
            self.set_processing_status("Please select both input file and output folder.")
            return

        try:
            threading.Thread(target=self.main_threaded,
                             args=(input_file, os.path.abspath(self.output_folder_var.get()))).start()
            self.generate_button.config(state=tk.DISABLED)
            self.set_processing_status("Processing...")

        except Exception as e:
            print("Error:", e)
            self.set_processing_status("Error during processing")
            self.generate_button.config(state=tk.NORMAL)

    def main_threaded(self, input_file, output_folder):
        try:
            main(input_file, output_folder, self)
            self.set_processing_status("File created successfully.")
            os.startfile(output_folder)
        except Exception as e:
            print("Error during processing:", e)
            self.set_processing_status("Error during processing")
        finally:
            self.generate_button.config(state=tk.NORMAL)

    def on_close(self):
        self.is_running = False
        self.master.destroy()


if __name__ == "__main__":
    themed = ThemedTk(theme='plastik')
    app = GUI(themed)
    themed.mainloop()
