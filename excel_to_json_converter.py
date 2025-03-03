# """
# Excel to JSON Converter Script

# 📌 Instructions:
# 1️⃣ Upload an Excel file that follows this structure:
#    --------------------------------------------------------
#    | Tower Name | Floor Number | Company Name(s)         |
#    --------------------------------------------------------
#    | Tower A    | 1           | Company X, Company Y    |
#    | Tower A    | 2           | Company Z               |
#    | Tower B    | 1           | Company A, Company B    |
#    --------------------------------------------------------

# 2️⃣ Ensure the column names are exactly:
#    - "Tower Name"
#    - "Floor Number"
#    - "Company Name(s)"

# 3️⃣ The "Company Name(s)" column should have comma-separated values if multiple companies occupy the same floor.

# 4️⃣ Once uploaded, the script will convert the data into a structured JSON format.

# 🚀 Happy converting!
# """

# from google.colab import files
# import pandas as pd
# import json
# import os
# from difflib import get_close_matches

# def select_excel_file():
#     uploaded = files.upload()
#     for filename in uploaded.keys():
#         print(f"✅ File '{filename}' uploaded successfully.")
#         return filename
#     return None

# def format_tower_name(tower):
#     tower = tower.strip()
#     if not tower.lower().startswith("tower-"):
#         return f"Tower-{tower}"
#     return tower

# def find_similar_columns(actual_columns, required_columns):
#     suggestions = {}
#     for required_col in required_columns:
#         close_matches = get_close_matches(required_col, actual_columns, n=1, cutoff=0.7)
#         if close_matches:
#             suggestions[required_col] = close_matches[0]
#     return suggestions

# def process_excel_data(file_path):
#     if not os.path.exists(file_path):
#         print("🚨 Error: File does not exist.")
#         return None

#     try:
#         df = pd.read_excel(file_path)
#     except Exception as e:
#         print(f"⚠️ Error reading the Excel file: {e}")
#         return None

#     required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
#     actual_columns = df.columns.tolist()
    
#     missing_columns = [col for col in required_columns if col not in actual_columns]
    
#     if missing_columns:
#         print(f"🚨 Error: Missing required columns: {', '.join(missing_columns)}")
        
#         suggestions = find_similar_columns(actual_columns, missing_columns)
#         if suggestions:
#             for missing_col, suggestion in suggestions.items():
#                 print(f"🔍 Did you mean '{suggestion}' instead of '{missing_col}'?")
        
#         print("❌ Please correct the column names and re-upload the file.")
#         return None

#     towers_dict = {}
#     for _, row in df.iterrows():
#         tower = format_tower_name(str(row["Tower Name"]))
#         floor = str(row["Floor Number"])
#         companies = [company.strip() for company in str(row["Company Name(s)"]).split(',')]

#         if tower not in towers_dict:
#             towers_dict[tower] = {}

#         if floor in towers_dict[tower]:
#             towers_dict[tower][floor].extend(companies)
#         else:
#             towers_dict[tower][floor] = companies

#     towers_list = [{"data": [{"name": list(set(companies)), "floor": floor} for floor, companies in floors.items()], "tower": tower} for tower, floors in towers_dict.items()]

#     return towers_list

# def save_to_json(data, excel_filename):
#     default_json_filename = os.path.splitext(excel_filename)[0] + ".json"

#     # Clearer instructions
#     print("\n💾 File Naming Instructions:")
#     print("✅ Default filename is already provided below.")
#     print("✏️ If you want to change it, please modify only the name. The extension will always remain .json.")
#     print("--------------------------------------------------------")
    
#     # Simulating an input box appearing below instructions
#     print(f"Enter filename (or press Enter to use '{default_json_filename}'): ")
#     user_filename = input().strip()  # Asking for input in the next line

#     json_filename = os.path.splitext(user_filename)[0] + ".json" if user_filename else default_json_filename

#     with open(json_filename, 'w', encoding='utf-8') as json_file:
#         json.dump(data, json_file, indent=2)

#     print(f"\n✅ Data successfully saved to '{json_filename}'.")
#     files.download(json_filename)
#     print("⬇️ Download started!")

# def main():
#     print("📂 Upload an Excel file to process.")

#     file_name = select_excel_file()

#     if not file_name:
#         print("❌ No file selected. Exiting...")
#         return

#     print(f"⚙️ Processing file: {file_name}")

#     data = process_excel_data(file_name)

#     if data is not None:
#         save_to_json(data, file_name)

# # Run the script
# main()





import gradio as gr
import pandas as pd
import json
import os
from google.colab import files
def process_excel(file_path):
    try:
        df = pd.read_excel(file_path, usecols=[0, 1, 2])  # Process only the first three columns
        df.columns = ["Tower Name", "Floor Number", "Company Name(s)"]  

        data_dict = []
        tower_dict = {}
        
        for _, row in df.iterrows():
            tower = f"Tower-{str(row['Tower Name']).strip()}"
            floor = str(row["Floor Number"]).strip()
            companies = [c.strip() for c in str(row["Company Name(s)"]).split(',')]
            
            if tower not in tower_dict:
                tower_dict[tower] = {}
                data_dict.append({"tower": tower, "data": []})
            
            if floor in tower_dict[tower]:
                tower_dict[tower][floor].extend(companies)
            else:
                tower_dict[tower][floor] = companies
                for entry in data_dict:
                    if entry["tower"] == tower:
                        entry["data"].append({"floor": floor, "name": companies})
                        break
        
        return data_dict, None
    except Exception as e:
        return None, f"⚠️ Error processing file: {str(e)}"

def handle_conversion(file_path, filename):
    if not file_path:
        return None, None, None, "❌ No file uploaded."
    
    json_data, error = process_excel(file_path)
    if error:
        return None, None, None, error

    json_filename = f"{filename.strip()}.json" if filename.strip() else "converted_data.json"
    json_path = os.path.join(os.getcwd(), json_filename)

    with open(json_path, 'w', encoding='utf-8') as json_file:
        json.dump(json_data, json_file, indent=2)

    json_string = json.dumps(json_data, indent=2)

    # 🚀 Auto-download in Google Colab
    try:
        files.download(json_path)  # Works in Google Colab
    except:
        pass  # Ignore if not in Colab

    return json_string, json_path, json_path, None

def ui():

    with gr.Blocks() as demo:
        # Load Logo
        logo_path = "logo.png"  # Local path
        github_logo_url = "https://github.com/rohansekhri12/Excel-go-json/blob/main/CIPIS_logo.jpg"  # Update with your GitHub raw link

        # Try displaying local logo first; fallback to GitHub URL
        logo_html = f'<img src="{logo_path}" alt="Logo" width="200">' if os.path.exists(logo_path) else f'<img src="{github_logo_url}" alt="Logo" width="200">'

        

    
        gr.HTML("""
        <div style="text-align:center;">
            <h1>📊 Excel to JSON Converter</h1>
            <h2> powered by CIPIS <h2>
            <p>Convert structured Excel files into JSON format easily.</p>
        </div>
        """)

        with gr.Row():
            with gr.Column():
                gr.HTML("""
                <div style="padding:10px; border:1px solid #ccc; border-radius:5px;">
                    <h3>✅ Instructions:</h3>
                    <ul>
                        <li>Ensure Excel contains only the first three columns: <b>Tower Name, Floor Number, Company Name(s)</b>.</li>
                        <li>Company names must be <b>comma-separated</b> if multiple companies share the same floor.</li>
                        <li>You can rename the JSON output (optional).</li>
                        <li>Click <b>Convert to JSON</b>, and download the converted file.</li>
                    </ul>
                </div>
                """)

        with gr.Row():
            file_input = gr.File(label="📎 Upload Your Excel File", type="filepath")
            filename_input = gr.Textbox(label="📝 Optional: Rename JSON File (without extension)")
        
        convert_button = gr.Button("🚀 Convert to JSON")

        with gr.Row():
            with gr.Column():
                excel_preview = gr.Dataframe(label="🐑 Excel Preview (before conversion)")
            with gr.Column():
                json_preview = gr.Textbox(label="📜 Converted JSON Output", lines=20)

        with gr.Row():
            download_button = gr.File(label="👥 Download JSON File", interactive=False)
        error_msg = gr.Textbox(label="⚠️ Error Messages", interactive=False, visible=False)

        file_input.change(lambda f: (pd.read_excel(f, usecols=[0, 1, 2]) if f else None), inputs=[file_input], outputs=[excel_preview])

        convert_button.click(
            handle_conversion,
            inputs=[file_input, filename_input],
            outputs=[json_preview, download_button, download_button, error_msg]
        )

    return demo

demo = ui()
demo.launch()
