import os
import sys
import time
import json
import logging
import traceback
import pandas as pd
import tkinter as tk
import multiprocessing
from openai import OpenAI
from dotenv import load_dotenv
from multiprocessing import Queue, Pool
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk, filedialog, messagebox

# review both text columns and categorize if each row represents an incident or a service request considering the ITIL framework. use the words incident and service request only. do not explain yoursel.
# Compare the names in the two lists and reply with "True" if they match or "False" if they do not. For each pair, provide the result in the format "name1 - name2: result".
# For the given name, determine if it exists in the provided list of names. Respond with 'True' if it does, and 'False' if it does not. Provide the output in the format 'name: True/False'.

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

MAX_INPUT_TOKENS = 2048

load_dotenv(dotenv_path='./config/.env')
api_key = os.getenv('OPENAI_API_KEY')
client = OpenAI(api_key=api_key)


class DataProcessor:
    def __init__(self, df):
        self.df = df

    def select_data(self, selected_columns):
        if selected_columns:
            return self.df[selected_columns]
        return self.df

    def prepare_prompts(self, instructions, selected_data):
        prompts = []
        data_records = selected_data.to_dict(orient='records')
        for idx, record in enumerate(data_records):
            try:
                prompt = instructions.format(**record)
            except KeyError as e:
                missing_key = str(e).strip("'")
                messagebox.showerror("Error", f"Column '{missing_key}' not found in the selected data.")
                return []
            prompts.append(prompt)
        return prompts


class OpenAIAPIClient:
    def __init__(self, model_name):
        self.model_name = model_name

    def get_response(self, prompt):
        try:
            response = client.chat.completions.create(
                model=self.model_name,
                messages=[{"role": "user", "content": prompt}],
                temperature=0
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            logging.error(f"OpenAI API error: {e}")
            return f"Error: {e}"


class ExcelAnalyzerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Analyzer Made For Rumi")
        self.geometry("600x500")
        self.create_widgets()
        self.queue = Queue()
        self.pool = Pool(processes=multiprocessing.cpu_count())

    def on_template_selected(self, event):
        selected_template = self.template_var.get()
        self.instruction_entry.delete("1.0", tk.END)
        self.instruction_entry.insert(tk.INSERT, selected_template)

    def create_widgets(self):
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", background="#8EC5FC")
        style.configure("TLabel", font=("Helvetica", 10), foreground="#4B4B4B")

        # Instruction Label
        ttk.Label(self, text="Enter instructions for analysis before selecting a file:").pack(pady=10)

        # Instruction Entry Box with Scroll
        self.instruction_entry = ScrolledText(self, height=10, width=70, wrap=tk.WORD)
        self.instruction_entry.insert(tk.INSERT, "Please provide analysis instructions using placeholders for column names in curly braces. E.g., 'For each entry, determine if {Column1} relates to {Column2}.'")
        self.instruction_entry.pack(pady=10)

        # Model Selection Label
        ttk.Label(self, text="Select AI model for analysis:").pack(pady=5)

        # Model Selection Combobox
        self.model_var = tk.StringVar()
        self.model_var.set("gpt-4o-mini")
        self.model_selection = ttk.Combobox(self, textvariable=self.model_var, values=["gpt-4o", "gpt-4o-mini", "gpt-4-turbo"], state="readonly")
        self.model_selection.pack(pady=5)

        # Template Selection Combobox
        ttk.Label(self, text="Or select an analysis template:").pack(pady=5)
        self.template_var = tk.StringVar()
        self.template_var.set("Select a template")
        templates = [
            "Categorize entries based on {Column1}",
            "Compare {Column1} with {Column2} and note differences",
            "Summarize the information in {Column1}"
        ]
        self.template_selection = ttk.Combobox(self, textvariable=self.template_var, values=templates, state="readonly")
        self.template_selection.pack(pady=5)
        self.template_selection.bind("<<ComboboxSelected>>", self.on_template_selected)

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

            # Selecting columns
            columns_to_analyze = self.select_columns(df.columns)
            if not columns_to_analyze:
                messagebox.showwarning("Warning", "No columns were selected for analysis.")
                self.update_status("")
                return

            # Initialize DataProcessor
            processor = DataProcessor(df)
            selected_data = processor.select_data(columns_to_analyze)

            # Prompt preparation using the user's instructions
            instructions_template = self.instruction_entry.get("1.0", tk.END).strip()
            prompts = processor.prepare_prompts(instructions_template, selected_data)
            if not prompts:
                self.update_status("No prompts to process.")
                return
            # Calling the API
            self.call_api_and_process_responses(prompts, df, file_path)

        except Exception as e:
            logging.error(f"An unexpected error occurred: {e}")
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")
            self.update_status("")
            self.reset_progress()

    def call_api_and_process_responses(self, prompts, df, file_path):
        total = len(prompts)
        self.reset_progress()
        self.update_status("Analyzing...")
        self.update()
        self.pool = Pool()

        manager = multiprocessing.Manager()
        return_dict = manager.dict()
        jobs = []

        model_name = self.model_var.get()

        for idx, prompt in enumerate(prompts):
            prompt = prompt[:MAX_INPUT_TOKENS]
            job = self.pool.apply_async(
                self.call_openai_api,
                args=(idx, prompt, return_dict, model_name)
            )
            jobs.append(job)

        # Progress tracking
        while any(not job.ready() for job in jobs):
            completed = sum(1 for job in jobs if job.ready())
            progress_percent = (completed / total) * 100
            self.update_progress(progress_percent)
            self.update()
            time.sleep(0.1)

        self.pool.close()
        self.pool.join()

        # Collect responses
        responses = [return_dict.get(i, 'Error') for i in range(total)]
        df['Analysis'] = responses

        output_path = self.get_output_path(file_path)
        self.save_output_file(df, output_path)

        messagebox.showinfo("Success", f"File has been analyzed and saved to {output_path}")
        self.update_status("Analysis complete.")
        logging.info("Analysis complete.")
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
                error_message = traceback.format_exc()
                logging.error(f"Failed to read the data file:\n{error_message}")
                messagebox.showerror("Error", f"Failed to read the data file:\n{e}")
                self.update_status("")
                return None

    def select_columns(self, columns):
        popup = tk.Toplevel(self)
        popup.title("Select Columns")
        popup.geometry("300x400")
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
    def call_openai_api(idx, prompt, return_dict, model_name):
        api_client = OpenAIAPIClient(model_name)
        reply = api_client.get_response(prompt)
        return_dict[idx] = reply
    
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
    setup_json_logging()
    app.mainloop()
