import os
import sys
import time
import json
import logging
import threading
import traceback
import pandas as pd
import tkinter as tk
import multiprocessing
import concurrent.futures
from openai import OpenAI
from dotenv import load_dotenv
from multiprocessing import Queue, Pool
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk, filedialog, messagebox


class JSONFormatter(logging.Formatter):
    def format(self, record):
        log_record = {
            'time': self.formatTime(record, self.datefmt),
            'level': record.levelname,
            'logger': record.name,
            'message': record.getMessage(),
        }
        return json.dumps(log_record)

def setup_json_logging():
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    console_handler = logging.StreamHandler(sys.stdout)
    
    console_handler.setLevel(logging.INFO)
    
    json_formatter = JSONFormatter()
    console_handler.setFormatter(json_formatter)
    logger.addHandler(console_handler)
    openai_logger = logging.getLogger('openai')
    openai_logger.setLevel(logging.DEBUG)
    openai_logger.propagate = True

load_dotenv(dotenv_path='./config/.env')
api_key = os.getenv('OPENAI_API_KEY')
client = OpenAI(api_key=api_key)


class ExcelAnalyzerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Analyzer Made For Rumi")
        self.geometry("600x500")
        self.create_widgets()
        self.queue = Queue()

    def create_widgets(self):
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", background="#8EC5FC")
        style.configure("TLabel", font=("Helvetica", 10), foreground="#4B4B4B")

        # Instruction Label
        ttk.Label(self, text="Enter instructions for analysis before selecting a file:").pack(pady=10)

        # Instruction Entry Box with Scroll
        self.instruction_entry = ScrolledText(self, height=10, width=70, wrap=tk.WORD)
        self.instruction_entry.insert(tk.INSERT, "E.g., 'Analyze sentiment of comments'...")
        self.instruction_entry.pack(pady=10)
        
        # Model Selection Label
        ttk.Label(self, text="Select AI model for analysis:").pack(pady=5)

        # Model Selection Combobox
        self.model_var = tk.StringVar()
        self.model_var.set("gpt-4o")  # Set a default value
        self.model_selection = ttk.Combobox(self, textvariable=self.model_var, values=["gpt-4o", "gpt-4o-mini", "gpt-4-turbo"], state="readonly")
        self.model_selection.pack(pady=5)

        # File Selection Button
        ttk.Button(self, text="Select Data File", command=self.select_file).pack(pady=15)

        # Selected File Label
        self.file_label = ttk.Label(self, text="No file selected", foreground="gray")
        self.file_label.pack(pady=5)

        # Progress Bar
        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.progress.pack(pady=10)

        # Status Label
        self.status_label = ttk.Label(self, text="", font=("Helvetica", 10, "italic"))
        self.status_label.pack(pady=5)

    def select_file(self):
        file_path = filedialog.askopenfilename(initialdir='/data', filetypes=[("Excel and CSV files", "*.xlsx *.xls *.csv")])
        if file_path:
            instructions = self.instruction_entry.get("1.0", tk.END).strip()
            if instructions:
                self.analyze_file(file_path, instructions)
            else:
                messagebox.showwarning("Warning", "Please provide instructions for analysis.")

    def analyze_file(self, file_path, instructions):
        try:
            logging.info(f"Analyzing file: {file_path}")
            self.update_status("Reading data file...")
            self.update()

            df = self.read_data_file(file_path)
            if df is None or df.empty:
                return

            columns_to_analyze = self.select_columns(df.columns)
            if not columns_to_analyze:
                messagebox.showwarning("Warning", "No columns were selected for analysis.")
                self.update_status("")
                return

            input_texts = df[columns_to_analyze].astype(str).agg(' '.join, axis=1).tolist()
            total = len(input_texts)

            self.reset_progress()
            self.update_status("Analyzing...")
            self.update()

            with Pool(processes=multiprocessing.cpu_count()) as pool:
                manager = multiprocessing.Manager()
                return_dict = manager.dict()
                
                jobs = [
                    pool.apply_async(self.call_openai_api, args=(idx, instructions, text, return_dict, self.model_var.get()))
                    for idx, text in enumerate(input_texts)
                ]

                while any(not job.ready() for job in jobs):
                    completed = sum(1 for job in jobs if job.ready())
                    progress_percent = (completed / total) * 100
                    self.update_progress(progress_percent)
                    self.update()
                    time.sleep(0.1)

            responses = [return_dict[i] for i in range(total)]
            df['Analysis'] = responses
            output_path = self.get_output_path(file_path)
            self.save_output_file(df, output_path)

            messagebox.showinfo("Success", f"File has been analyzed and saved to {output_path}")
            self.update_status("Analysis complete.")
            logging.info("Analysis complete.")
            self.reset_progress()

        except Exception as e:
            logging.error(f"An unexpected error occurred: {e}", exc_info=True)
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")
            self.update_status("")
            self.reset_progress()

    def read_data_file(self, file_path):
        try:
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path, engine='openpyxl')
            elif file_path.endswith('.xls'):
                df = pd.read_excel(file_path, engine='xlrd')
            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                messagebox.showerror("Error", "Unsupported file format. Please select an Excel or CSV file.")
                self.update_status("")
                return None
            if df.empty:
                messagebox.showwarning("Warning", "The selected file is empty.")
                self.update_status("")
                return None
            return df
        except Exception as e:
            logging.error(f"Failed to read the data file: {e}", exc_info=True)
            messagebox.showerror("Error", f"Failed to read the data file:\n{e}")
            self.update_status("")
            return None

    def select_columns(self, columns):
        popup = tk.Toplevel(self)
        popup.title("Selector")
        popup.geometry("250x450")
        popup.transient(self)
        popup.grab_set()

        ttk.Label(popup, text="Select columns for analysis:").pack(pady=10)

        column_vars = []
        for column in columns:
            var = tk.BooleanVar()
            ttk.Checkbutton(popup, text=column, variable=var).pack(anchor=tk.W)
            column_vars.append((column, var))

        selected_columns = []

        def select_and_close():
            for col, var in column_vars:
                if var.get():
                    selected_columns.append(col)
            if not selected_columns:
                messagebox.showwarning("Warning", "Please select at least one column.")
            else:
                popup.destroy()

        ttk.Button(popup, text="Select", command=select_and_close).pack(pady=10)
        self.center_window(popup)
        popup.wait_window()
        return selected_columns

    def center_window(self, window):
        window.update_idletasks()
        window_width = window.winfo_width()
        window_height = window.winfo_height()

        main_x = self.winfo_x()
        main_y = self.winfo_y()
        main_width = self.winfo_width()
        main_height = self.winfo_height()

        pos_x = main_x + (main_width // 2) - (window_width // 2)
        pos_y = main_y + (main_height // 2) - (window_height // 2)

        window.geometry(f"+{pos_x}+{pos_y}")

    @staticmethod
    def call_openai_api(idx, instructions, text, return_dict, model_name):
        try:
            response = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "user", "content": f"{instructions}\n\nText: {text}"}],
                max_tokens=100
            )
            reply = response.choices[0].message.content.strip()
            return_dict[idx] = reply
        except Exception as e:
            logging.error(f"OpenAI API error at index {idx}: {e}", exc_info=True)
            return_dict[idx] = f"Error: {e}"

    def update_status(self, message):
        self.status_label.config(text=message)

    def update_progress(self, value):
        self.progress['value'] = value

    def reset_progress(self):
        self.progress['value'] = 0

    def get_output_path(self, file_path):
        directory, filename = os.path.split(file_path)
        name, ext = os.path.splitext(filename)
        new_filename = f"{name}_analyzed{ext}"
        return os.path.join(directory, new_filename)

    def save_output_file(self, df, output_path):
        try:
            if output_path.endswith('.xlsx') or output_path.endswith('.xls'):
                df.to_excel(output_path, index=False)
            elif output_path.endswith('.csv'):
                df.to_csv(output_path, index=False)
            else:
                pass
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save the output file: {e}")
            self.update_status("")

if __name__ == "__main__":
    app = ExcelAnalyzerApp()
    setup_json_logging()
    app.mainloop()