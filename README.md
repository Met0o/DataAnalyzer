# DataAnalyzerApp

## Overview

I developed this app for my wife. 

DataAnalyzerApp is a Python-based graphical application that allows for customizable instruction-based analysis using OpenAI's Model API. It can handle data files in Excel and CSV formats. The app provides a simple user interface for specifying analysis instructions, selecting data columns, and executing custom NLP analysis.

The application is built using the `tkinter` library for UI and leverages `multiprocessing` to speed up the analysis by using allocating available CPU cores.

## Features
- Customizable Analysis Instructions: 
   - Input your own instructions for the AI model to perform on the data.
- Select Data Files: 
   - Analyze data from Excel (.xlsx, .xls) or CSV (.csv) files.
- Analysis Modes: 
   - Choose between row-wise analysis and column-wise analysis.
- Row Analysis: 
   - Analyze data row by row, suitable for tasks like sentiment analysis or data classification.
- Column Analysis: 
   - Compare data between two columns, even when the data is not aligned by rows, suitable for tasks like data lookup or cross-referencing.
- Select Specific Columns: 
   - Choose which columns in your data file to include in the analysis.
- Model Selection: 
   - Choose from several OpenAI GPT models for analysis.
- Concurrent Processing: 
   - Supports concurrent processing for faster analysis by utilizing multiple CPU cores.
- Progress Monitoring: 
   - Progress bar and status updates to inform users during long-running analyses.
- Error Handling and Logging: 
   - Provides informative error messages and logs errors for debugging purposes.

## Prerequisites
- Python 3.10 or higher
  - `pandas`
  - `openpyxl`
  - `openai`
  - `xlrd`

## Installation
1. Clone the repository:
   ```sh
   git clone https://github.com/Met0o/DataAnalyzer.git
   cd DataAnalyzerApp
   ```
2. Install the required dependencies:
   ```sh
   pip install pandas openpyxl xlrd openai
   ```

## Usage
1. Run the application:
   ```sh
   python app.py
   ```
2. Enter the instructions for the type of analysis you wish to perform in the text box.
   - In the text box labeled "Enter instructions for analysis before selecting a file", input the instructions for the type of analysis you wish to perform.
   Example: "Analyze the sentiment of the following text and classify it as Positive, Negative, or Neutral."
   Example: "Compare the names in the two lists and reply with 'True' if they match or 'False' if they do not."
3. Click the "Select Data File" button to choose an Excel or CSV file.
4. A pop-up window will allow you to select the columns for analysis.
5. Select Analysis Type: Choose the analysis mode by selecting either "Row Analysis" or "Column Analysis":
   - Row Analysis: The model will process each row individually.
   - Column Analysis: The model will compare data across columns, useful for comparing data that is not aligned by rows.
5. Pick one of the avaialble models to perform the analysis.
6. The analysis progress will be displayed using the progress bar.
7. Once complete, the analyzed file will be saved with a suffix `_analyzed` in the original file's directory.

## Configuration
- **CPU Cores**: The app automatically detects the number of available CPU cores using `multiprocessing.cpu_count()`, allowing the workload to be distributed efficiently.
- **Output File**: The analyzed data is saved in the same format as the original file, with `_analyzed` appended to the filename.
- **Analysis Instructions**: Provide specific instructions in the text box for how the analysis should be conducted (e.g., "Perform sentiment analysis on the comments").
- **Model Selection**: Pick one of 3 models, GPT-4o being most capable, followed by GPT-4o-mini, and GPT-4-Turbo.

## Examples
Row Analysis Example

1. Instructions:
   ```sh
   Analyze the following customer feedback and classify it as Positive, Negative, or Neutral sentiment.
   ```
2. Data:
   - A CSV file with a column named 'Feedback' containing customer comments.
3. Process:
   - Select "Row Analysis" mode.
   - Pick your column/s for the analysis.
   - The model will analyze each row for the selected columns individually and provide a sentiment classification.

Column Analysis Example

1. Instructions:
   ```sh
   Compare the names in the two lists and reply with "True" if they match or "False" if they do not. For each pair, provide the result in the format "name1 - name2: result".
   ```
2. Data:
   - An Excel file with two columns: 'Name 1' and 'Name 2'.
3. Process:
   - Select "Column Analysis" mode.
   - Pick your columns for the analysis.
   - The model will compare the names in the two columns and output whether they match.

## Screenshots
*Main User Interface*

![Main UI](UI.png)

## Notes
- Make sure you have an active API key for OpenAI to use the application.

## Troubleshooting
- **Unsupported File Format**: If you get an error about an unsupported file, ensure you select either `.xlsx`, `.xls`, or `.csv` files.
- **OpenAI API Errors**: If there are issues connecting to the OpenAI API, verify your API key and network connection.
- **Empty File Warning**: If the selected file is empty, the app will notify you. Make sure the file contains valid data.
