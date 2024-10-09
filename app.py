import os
import time
import logging
import threading
import pandas as pd
import tkinter as tk
import multiprocessing
import concurrent.futures
from openai import OpenAI
from multiprocessing import Queue, Pool
from tkinter import ttk, filedialog, messagebox


# review both text columns and categorize if each row represents an incident or a ticket considering the ITIL framework. use the words ticker and incident only. do not explain yourself

logging.basicConfig(level=logging.INFO)
MAX_INPUT_TOKENS = 2048

client = OpenAI(
    api_key = "",
)

class ExcelAnalyzerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Analyzer")
        self.geometry("600x500")
        self.create_widgets()
        self.queue = Queue()
        self.pool = 10

    def create_widgets(self):
        
        ttk.Label(self, text="Enter instructions for analysis:").pack(pady=5)
        
        self.instruction_entry = tk.Text(self, height=5, width=70)
        self.instruction_entry.pack(pady=5)
        
        ttk.Button(self, text="Select Data File", command=self.select_file).pack(pady=10)
        
        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.progress.pack(pady=10)
        self.status_label = ttk.Label(self, text="")
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
            self.pool = Pool()
            manager = multiprocessing.Manager()
            return_dict = manager.dict()
            jobs = []

            for idx, text in enumerate(input_texts):
                text = text[:MAX_INPUT_TOKENS]
                prompt = f"{instructions}\n\nText: {text}"
                job = self.pool.apply_async(self.call_openai_api, args=(idx, prompt, return_dict))
                jobs.append(job)

            while any(not job.ready() for job in jobs):
                completed = sum(1 for job in jobs if job.ready())
                progress_percent = (completed / total) * 100
                self.update_progress(progress_percent)
                self.update()
                time.sleep(0.1)
            time.sleep(5)
            self.pool.close()
            self.pool.join()

            responses = [return_dict[i] for i in range(total)]
            df['Analysis'] = responses
            output_path = self.get_output_path(file_path)
            self.save_output_file(df, output_path)

            messagebox.showinfo("Success", f"File has been analyzed and saved to {output_path}")
            self.update_status("Analysis complete.")
            logging.info("Analysis complete.")
            self.reset_progress()

        except Exception as e:
            logging.error(f"An unexpected error occurred: {e}")
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")
            self.update_status("")
            self.reset_progress()

    def read_data_file(self, file_path):
        file_ext = os.path.splitext(file_path)[1].lower()
        try:
            if file_ext == '.xlsx':
                df = pd.read_excel(file_path, engine='openpyxl')
            elif file_ext == '.xls':
                df = pd.read_excel(file_path, engine='xlrd')
            elif file_ext == '.csv':
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
    def call_openai_api(idx, prompt, return_dict):
        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=100
            )
            reply = response.choices[0].message.content.strip()
            return_dict[idx] = reply
        except Exception as e:
            logging.error(f"OpenAI API error at index {idx}: {e}")
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
        file_ext = os.path.splitext(output_path)[1].lower()
        try:
            if file_ext in ['.xlsx', '.xls']:
                df.to_excel(output_path, index=False)
            elif file_ext == '.csv':
                df.to_csv(output_path, index=False)
            else:
                pass
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save the output file:\n{e}")
            self.update_status("")

if __name__ == "__main__":
    app = ExcelAnalyzerApp()
    app.mainloop()
