from flask import Flask, request, render_template, send_file
import os
import pandas as pd
import logging
import time
from openai import OpenAI
from multiprocessing import Pool, Manager


client = OpenAI(
    api_key = "sk-Qrho3uAETH1KoJsptvsQrZfcq917vf5NLr9FVIemxKT3BlbkFJkAS0NJNHcTN2f421pfDPXC9yFr_6K9Ul4q8wjEQQMA",
)

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    instructions = request.form['instructions']
    file = request.files['datafile']

    if not instructions or not file:
        return "Please provide analysis instructions and a data file", 400

    file_path = os.path.join('uploads', file.filename)
    file.save(file_path)

    # Read the data file
    if file.filename.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(file_path)
    elif file.filename.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        return "Unsupported file format. Please upload an Excel or CSV file.", 400
    
    # Analyze the file using OpenAI API
    columns_to_analyze = df.columns.tolist()
    input_texts = df[columns_to_analyze].astype(str).agg(' '.join, axis=1).tolist()
    total = len(input_texts)

    pool = Pool()
    manager = Manager()
    return_dict = manager.dict()
    jobs = []

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

    for idx, text in enumerate(input_texts):
        prompt = f"{instructions}\n\nText: {text}"
        job = pool.apply_async(call_openai_api, args=(idx, prompt, return_dict))
        jobs.append(job)

    pool.close()
    pool.join()

    # Ensure all keys exist in return_dict
    responses = [return_dict.get(i, "Error: No response") for i in range(total)]
    df['Analysis'] = responses

    # Save the output file
    output_path = os.path.join('outputs', f"{os.path.splitext(file.filename)[0]}_analyzed.csv")
    df.to_csv(output_path, index=False)

    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    if not os.path.exists('outputs'):
        os.makedirs('outputs')
    app.run(host='0.0.0.0', port=5000, debug=True)