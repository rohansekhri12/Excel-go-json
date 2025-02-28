"""
Excel to JSON Converter Script

📌 Instructions:
1️⃣ Upload an Excel file that follows this structure:
   --------------------------------------------------------
   | Tower Name | Floor Number | Company Name(s)         |
   --------------------------------------------------------
   | Tower A    | 1           | Company X, Company Y    |
   | Tower A    | 2           | Company Z               |
   | Tower B    | 1           | Company A, Company B    |
   --------------------------------------------------------

2️⃣ Ensure the column names are exactly:
   - "Tower Name"
   - "Floor Number"
   - "Company Name(s)"

3️⃣ The "Company Name(s)" column should have comma-separated values if multiple companies occupy the same floor.

4️⃣ Once uploaded, the script will convert the data into a structured JSON format.

🚀 Happy converting!
"""

import gradio as gr
from google.colab import files
import pandas as pd
import json
import os
from difflib import get_close_matches

def format_tower_name(tower):
    tower = tower.strip()
    if not tower.lower().startswith("tower-"):
        return f"Tower-{tower}"
    return tower

def find_similar_columns(actual_columns, required_columns):
    suggestions = {}
    for required_col in required_columns:
        close_matches = get_close_matches(required_col, actual_columns, n=1, cutoff=0.7)
        if close_matches:
            suggestions[required_col] = close_matches[0]
    return suggestions

def process_excel_data(file):
    df = pd.read_excel(file.name)
    required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
    actual_columns = df.columns.tolist()
    
    missing_columns = [col for col in required_columns if col not in actual_columns]
    if missing_columns:
        suggestions = find_similar_columns(actual_columns, missing_columns)
        error_message = f"Missing required columns: {', '.join(missing_columns)}"
        if suggestions:
            error_message += "\nSuggested corrections: " + ", ".join([f"{k} → {v}" for k, v in suggestions.items()])
        return None, error_message
    
    towers_dict = {}
    for _, row in df.iterrows():
        tower = format_tower_name(str(row["Tower Name"]))
        floor = str(row["Floor Number"])
        companies = [company.strip() for company in str(row["Company Name(s)"]).split(',')]

        if tower not in towers_dict:
            towers_dict[tower] = {}
        
        if floor in towers_dict[tower]:
            towers_dict[tower][floor].extend(companies)
        else:
            towers_dict[tower][floor] = companies
    
    towers_list = [{"data": [{"name": list(set(companies)), "floor": floor} for floor, companies in floors.items()], "tower": tower} for tower, floors in towers_dict.items()]
    return towers_list, None

def convert_and_download(file, filename):
    if not file:
        return "No file uploaded. Please upload an Excel file.", None
    
    json_data, error_message = process_excel_data(file)
    if error_message:
        return error_message, None
    
    json_filename = os.path.splitext(filename)[0] + ".json"
    with open(json_filename, "w", encoding="utf-8") as json_file:
        json.dump(json_data, json_file, indent=2)
    
    return f"✅ Data successfully saved as {json_filename}. Download it below:", json_filename

def download_json(json_filename):
    return files.download(json_filename)

with gr.Blocks() as ui:
    gr.Markdown("""<h1 style='text-align: center;'>Excel to JSON Converter</h1>""")
    gr.Image("CIPIS logo.jpg", elem_id="logo", width=200)
    gr.Markdown("""<p style='text-align: center;'>Powered by CIPIS</p>""")
    
    with gr.Row():
        file_input = gr.File(label="Upload Excel File")
        filename_input = gr.Textbox(label="Enter JSON filename (without extension)", value="converted_data")
    
    convert_button = gr.Button("Convert to JSON")
    output_text = gr.Textbox(label="Status", interactive=False)
    download_button = gr.Button("Download JSON", visible=False)
    
    def handle_conversion(file, filename):
        message, json_filename = convert_and_download(file, filename)
        download_button.visible = bool(json_filename)
        return message, json_filename
    
    convert_button.click(handle_conversion, inputs=[file_input, filename_input], outputs=[output_text, download_button])
    download_button.click(download_json, inputs=[], outputs=[])

ui.launch()
