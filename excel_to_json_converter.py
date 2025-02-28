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

# 📦 Auto-install dependencies if missing
try:
    import gradio as gr
except ModuleNotFoundError:
    print("📦 Installing Gradio, please wait...")
    !pip install gradio --quiet
    import gradio as gr

import pandas as pd
import json
import os
from difflib import get_close_matches

# 🔹 Function to format tower names
def format_tower_name(tower):
    tower = tower.strip()
    if not tower.lower().startswith("tower-"):
        return f"Tower-{tower}"
    return tower

# 🔍 Function to find similar column names if incorrect
def find_similar_columns(actual_columns, required_columns):
    suggestions = {}
    for required_col in required_columns:
        close_matches = get_close_matches(required_col, actual_columns, n=1, cutoff=0.7)
        if close_matches:
            suggestions[required_col] = close_matches[0]
    return suggestions

# 🚀 Process the uploaded Excel file
def process_excel(file):
    try:
        df = pd.read_excel(file.name)
    except Exception as e:
        return None, f"❌ Error reading file: {str(e)}"

    required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
    actual_columns = df.columns.tolist()
    
    missing_columns = [col for col in required_columns if col not in actual_columns]
    column_suggestions = find_similar_columns(actual_columns, missing_columns)
    
    if missing_columns:
        error_msg = f"🚨 Missing columns: {', '.join(missing_columns)}\n"
        if column_suggestions:
            for missing_col, suggestion in column_suggestions.items():
                error_msg += f"🔍 Did you mean '{suggestion}' instead of '{missing_col}'?\n"
        return None, error_msg + "❌ Please correct column names and re-upload the file."

    # Convert Excel data to structured JSON format
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

# 💾 Save JSON and return download link
def convert_excel_to_json(file, json_filename):
    if not file:
        return "❌ No file uploaded.", None

    data, error = process_excel(file)
    if error:
        return error, None

    # Ensure .json extension
    json_filename = os.path.splitext(json_filename)[0] + ".json"

    with open(json_filename, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, indent=2)

    return f"✅ JSON file '{json_filename}' created successfully!", json_filename

# 🎨 Gradio UI
def ui(file, json_filename):
    message, json_path = convert_excel_to_json(file, json_filename)
    return message, json_path if json_path else None

with gr.Blocks() as app:
    gr.Markdown("## 📂 Excel to JSON Converter")
    gr.Markdown(
        """
        🚀 **Instructions:**  
        1️⃣ Upload an Excel file with these **exact** column names:  
        - "Tower Name"  
        - "Floor Number"  
        - "Company Name(s)" (comma-separated if multiple companies)  
        2️⃣ Enter a name for the JSON output file **(default extension will be .json)**  
        3️⃣ Click **Convert** to generate the JSON  
        4️⃣ Download the JSON file when ready  
        """
    )

    with gr.Row():
        file_input = gr.File(label="📂 Upload Excel File", type="file")
        json_filename_input = gr.Textbox(label="💾 JSON Filename (without .json)", placeholder="output")

    convert_button = gr.Button("🚀 Convert")
    output_message = gr.Textbox(label="📝 Status", interactive=False)
    download_button = gr.File(label="⬇️ Download JSON", interactive=False)

    convert_button.click(ui, inputs=[file_input, json_filename_input], outputs=[output_message, download_button])

    # 🔹 Display logo
    gr.Image("CIPIS logo.jpg", label="Powered by CIPIS", show_label=True)

# 🔥 Launch the Gradio app
app.launch()

