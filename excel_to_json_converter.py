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

from google.colab import files
import pandas as pd
import json
import os
from difflib import get_close_matches  # To suggest correct column names

def select_excel_file():
    """Allows user to upload an Excel file in Google Colab."""
    uploaded = files.upload()
    for filename in uploaded.keys():
        print(f"✅ File '{filename}' uploaded successfully.")
        return filename  # Return the name of the uploaded file
    return None

def format_tower_name(tower):
    """Ensures the tower name follows 'Tower-X' format but avoids duplication."""
    tower = tower.strip()
    if not tower.lower().startswith("tower-"):
        return f"Tower-{tower}"  # Add prefix only if missing
    return tower  # Return as-is if already formatted

def find_similar_columns(actual_columns, required_columns):
    """Suggests correct column names for missing ones."""
    suggestions = {}
    for required_col in required_columns:
        close_matches = get_close_matches(required_col, actual_columns, n=1, cutoff=0.7)
        if close_matches:
            suggestions[required_col] = close_matches[0]
    return suggestions

def process_excel_data(file_path):
    """Reads the Excel file and processes it into JSON format."""
    if not os.path.exists(file_path):
        print("🚨 Error: File does not exist.")
        return None

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"⚠️ Error reading the Excel file: {e}")
        return None

    required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
    actual_columns = df.columns.tolist()
    
    # Find missing columns
    missing_columns = [col for col in required_columns if col not in actual_columns]
    
    if missing_columns:
        print(f"🚨 Error: Missing required columns: {', '.join(missing_columns)}")
        
        # Suggest corrections
        suggestions = find_similar_columns(actual_columns, missing_columns)
        if suggestions:
            for missing_col, suggestion in suggestions.items():
                print(f"🔍 Did you mean '{suggestion}' instead of '{missing_col}'?")
        
        print("❌ Please correct the column names and re-upload the file.")
        return None

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

    return towers_list

def save_to_json(data, excel_filename):
    """Saves the processed data to a JSON file and allows the user to download it."""
    default_json_filename = os.path.splitext(excel_filename)[0] + ".json"

    # Clearer instructions
    print("\n💾 File Naming Instructions:")
    print("✅ Default filename is already provided below.")
    print("✏️ If you want to change it, please modify only the name. The extension will always remain .json.")
    print("--------------------------------------------------------")

    # Simulated placeholder with a clear message
    user_filename = input(f"Enter filename (or press Enter to use '{default_json_filename}'): ").strip()

    # Ensure .json extension is added
    json_filename = os.path.splitext(user_filename)[0] + ".json" if user_filename else default_json_filename

    with open(json_filename, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, indent=2)

    print(f"\n✅ Data successfully saved to '{json_filename}'.")
    files.download(json_filename)
    print("⬇️ Download started!")

def main():
    """Main function to execute the script."""
    print("📂 Upload an Excel file to process.")

    file_name = select_excel_file()

    if not file_name:
        print("❌ No file selected. Exiting...")
        return

    print(f"⚙️ Processing file: {file_name}")

    data = process_excel_data(file_name)

    if data is not None:
        save_to_json(data, file_name)

# Run the script
main()

