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
import pandas as pd
import json
import os
import io
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
    df = pd.read_excel(file)
    
    required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
    actual_columns = df.columns.tolist()
    
    missing_columns = [col for col in required_columns if col not in actual_columns]
    
    if missing_columns:
        suggestions = find_similar_columns(actual_columns, missing_columns)
        error_message = "🚨 Error: Missing required columns: " + ", ".join(missing_columns)
        if suggestions:
            for missing_col, suggestion in suggestions.items():
                error_message += f"\n🔍 Did you mean '{suggestion}' instead of '{missing_col}'?"
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
    return towers_list, "✅ JSON generated successfully!"

def convert_to_json(file, filename):
    if not file:
        return None, "🚨 No file uploaded!"
    
    json_data, message = process_excel_data(file)
    if json_data is None:
        return None, message
    
    json_filename = os.path.splitext(filename)[0] + ".json"
    json_bytes = io.BytesIO()
    json_bytes.write(json.dumps(json_data, indent=2).encode('utf-8'))
    json_bytes.seek(0)
    
    return json_bytes, json_filename

def ui():
    with gr.Blocks() as app:
        gr.Markdown("""
        # 📂 Excel to JSON Converter
        🚀 **Convert your structured Excel data into JSON format effortlessly!**
        
        ### 📌 Instructions:
        1️⃣ Upload an **Excel file** that follows this structure:
        - Columns: `Tower Name`, `Floor Number`, `Company Name(s)` (comma-separated if multiple)
        
        2️⃣ Click **Convert to JSON**.
        3️⃣ Download the **automatically generated JSON file**.
        
        **Powered by CIPIS** 🏢
        """)
        
        with gr.Row():
            file_input = gr.File(label="📂 Upload Excel File")
            preview_box = gr.Dataframe(label="📜 File Preview", interactive=False)
        
        filename_input = gr.Textbox(label="✏️ Optional: Change JSON Filename (without extension)", placeholder="Default: filename.json")
        convert_button = gr.Button("⚡ Convert to JSON")
        output_message = gr.Markdown()
        json_download = gr.File(label="⬇️ Download JSON File", interactive=False)
        
        def update_preview(file):
            if file is None:
                return None
            df = pd.read_excel(file)
            return df.head()
        
        file_input.change(update_preview, inputs=[file_input], outputs=[preview_box])
        convert_button.click(convert_to_json, inputs=[file_input, filename_input], outputs=[json_download, output_message])
        
        gr.Markdown("""
        ---
        🚀 **Created with ❤️ by CIPIS**
        """)
    return app

app = ui()
app.launch()
