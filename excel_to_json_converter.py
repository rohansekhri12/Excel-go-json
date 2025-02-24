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

def select_excel_file():
    """Allows user to upload an Excel file in Google Colab."""
    uploaded = files.upload()  # Opens file upload dialog
    for filename in uploaded.keys():
        print(f"✅ File '{filename}' uploaded successfully.")
        return filename  # Return the name of the uploaded file
    return None

def format_tower_name(tower):
    """Formats the tower name to ensure it follows 'Tower-X' format."""
    return f"Tower-{tower}".replace(" ", "")

def process_excel_data(file_path):
    """Reads the Excel file and processes it into JSON format."""

    # Check if file exists
    if not os.path.exists(file_path):
        print("🚨 Error: File does not exist.")
        return None

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"⚠️ Error reading the Excel file: {e}")
        return None

    required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
    for col in required_columns:
        if col not in df.columns:
            print(f"🚨 Error: Missing required column '{col}' in the Excel file.")
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

def save_to_json(data):
    """Saves the processed data to a JSON file and allows the user to download it."""
    json_filename = "tower_data.json"

    with open(json_filename, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, indent=2)

    print(f"✅ Data successfully saved to '{json_filename}'.")

    # Provide download link
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
        save_to_json(data)

# Run the script
main()
