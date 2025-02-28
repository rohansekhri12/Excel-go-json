import gradio as gr
import pandas as pd
import json
import os
from difflib import get_close_matches

# Ensure Gradio is installed in Colab
def install_dependencies():
    try:
        import gradio
    except ImportError:
        os.system("pip install gradio")
install_dependencies()

# Function to process Excel data
def process_excel_data(file):
    df = pd.read_excel(file.name)
    required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
    actual_columns = df.columns.tolist()
    
    # Check for missing or incorrect columns
    missing_columns = [col for col in required_columns if col not in actual_columns]
    if missing_columns:
        suggestions = {col: get_close_matches(col, actual_columns, n=1, cutoff=0.7) for col in missing_columns}
        return f"❌ Error: Missing required columns: {', '.join(missing_columns)}", None
    
    # Convert data to JSON format
    towers_dict = {}
    for _, row in df.iterrows():
        tower = f"Tower-{row['Tower Name'].strip()}"
        floor = str(row['Floor Number'])
        companies = [c.strip() for c in str(row['Company Name(s)']).split(',')]
        if tower not in towers_dict:
            towers_dict[tower] = {}
        if floor in towers_dict[tower]:
            towers_dict[tower][floor].extend(companies)
        else:
            towers_dict[tower][floor] = companies
    
    towers_list = [{"data": [{"name": list(set(companies)), "floor": floor} for floor, companies in floors.items()], "tower": tower} for tower, floors in towers_dict.items()]
    
    # Save to JSON
    json_filename = file.name.replace('.xlsx', '.json')
    with open(json_filename, 'w', encoding='utf-8') as json_file:
        json.dump(towers_list, json_file, indent=2)
    
    return "✅ JSON conversion successful! Click below to download.", json_filename

# UI Components
def interface(file):
    if file is None:
        return "⚠️ Please upload an Excel file.", None
    return process_excel_data(file)

with gr.Blocks() as app:
    gr.Markdown("""
    # 🏢 Excel to JSON Converter
    ### 📌 Instructions:
    1️⃣ Upload an **Excel file** with the following columns:
    - **Tower Name**
    - **Floor Number**
    - **Company Name(s)** (comma-separated values)
    
    2️⃣ The system will process your file and provide a **JSON download link**.
    
    3️⃣ If any errors occur (e.g., missing column names), they will be displayed below.
    
    ---
    #### Powered by CIPIS  
    ![CIPIS Logo](./CIPIS_logo.jpg)
    
    ---
    """)
    
    file_input = gr.File(label="📂 Upload Excel File", type="file")
    output_text = gr.Textbox(label="🔔 Status", interactive=False)
    download_button = gr.DownloadButton("⬇️ Download JSON", interactive=False)
    
    file_input.change(interface, inputs=file_input, outputs=[output_text, download_button])

app.launch()
