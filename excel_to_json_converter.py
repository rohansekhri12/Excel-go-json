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

"""
Excel to JSON Converter with Gradio UI

📌 Features:
- Easy-to-use UI for uploading Excel files.
- Error handling for missing or incorrect column names.
- Displays suggested corrections if column names are wrong.
- Allows users to rename the JSON file while ensuring `.json` extension.
- Automatically downloads the JSON file after conversion.
- Includes a logo and "Powered by CIPIS" branding.

🚀 Built for non-coders: Just click 'Run' and use the interface!
"""

# Ensure Gradio is installed
try:
    import gradio as gr
except ModuleNotFoundError:
    import os
    os.system("pip install gradio")
    import gradio as gr

import pandas as pd
import json
import os
from difflib import get_close_matches

# Function to format tower names
def format_tower_name(tower):
    tower = tower.strip()
    if not tower.lower().startswith("tower-"):
        return f"Tower-{tower}"
    return tower

# Function to suggest similar column names
def find_similar_columns(actual_columns, required_columns):
    suggestions = {}
    for required_col in required_columns:
        close_matches = get_close_matches(required_col, actual_columns, n=1, cutoff=0.7)
        if close_matches:
            suggestions[required_col] = close_matches[0]
    return suggestions

# Function to process Excel file
def process_excel_file(file):
    try:
        df = pd.read_excel(file)
    except Exception as e:
        return f"⚠️ Error reading the Excel file: {e}", None, None

    required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
    actual_columns = df.columns.tolist()

    missing_columns = [col for col in required_columns if col not in actual_columns]
    if missing_columns:
        suggestions = find_similar_columns(actual_columns, missing_columns)
        error_message = f"🚨 Missing required columns: {', '.join(missing_columns)}"
        
        if suggestions:
            suggestion_texts = [f"🔍 Did you mean '{suggestion}' instead of '{missing_col}'?" for missing_col, suggestion in suggestions.items()]
            error_message += "\n" + "\n".join(suggestion_texts)

        return error_message, None, None

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

    return None, towers_list, os.path.splitext(file.name)[0] + ".json"

# Function to save JSON and provide download link
def convert_and_download(file, json_filename):
    error_message, json_data, default_json_filename = process_excel_file(file)

    if error_message:
        return error_message, None

    json_filename = os.path.splitext(json_filename)[0] + ".json"

    with open(json_filename, 'w', encoding='utf-8') as json_file:
        json.dump(json_data, json_file, indent=2)

    return f"✅ JSON successfully saved as '{json_filename}'!", json_filename

# Gradio UI
def gradio_interface(file, json_filename):
    if file is None:
        return "❌ Please upload an Excel file.", None

    return convert_and_download(file, json_filename)

# Create the Gradio app
with gr.Blocks() as app:
    gr.Markdown("## 📝 Excel to JSON Converter")
    gr.Markdown("### 📂 Upload your Excel file and get structured JSON output!")
    gr.Markdown("📌 **Instructions:**")
    gr.Markdown("""
    1. Upload an Excel file that follows this structure:
       ```
       | Tower Name | Floor Number | Company Name(s)         |
       |-----------|--------------|-------------------------|
       | Tower A    | 1           | Company X, Company Y    |
       | Tower A    | 2           | Company Z               |
       | Tower B    | 1           | Company A, Company B    |
       ```
    2. Ensure column names are exactly:
       - "Tower Name"
       - "Floor Number"
       - "Company Name(s)"
    3. You can rename the output JSON file but **the extension will always remain .json**.
    4. If an error occurs, suggestions will be provided.
    """)

    gr.Image("CIPIS logo.jpg", elem_id="logo", label="Powered by CIPIS", show_label=False, type="filepath")

    file_input = gr.File(label="Upload Excel File")
    json_name_input = gr.Textbox(label="JSON Filename (Optional, Default will be used if left blank)")

    output_text = gr.Textbox(label="Status", interactive=False)
    json_download = gr.File(label="Download JSON")

    submit_button = gr.Button("Convert to JSON")

    submit_button.click(gradio_interface, inputs=[file_input, json_name_input], outputs=[output_text, json_download])

# Launch the Gradio app
app.launch()

