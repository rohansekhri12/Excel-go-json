import pandas as pd
import json
import os
from google.colab import files

def upload_file():
    """Handles file selection in Colab."""
    try:
        uploaded = files.upload()  # Opens file selection dialog in Colab
        file_name = list(uploaded.keys())[0]  # Get uploaded file name
        return file_name  # Return the path of uploaded file
    except Exception as e:
        print(f"File upload error: {e}")
        return None

def process_excel_data(file_path):
    """Reads the Excel file and processes it into JSON format."""
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

    required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
    for col in required_columns:
        if col not in df.columns:
            print(f"Error: Missing required column '{col}' in the Excel file.")
            return None

    towers_dict = {}
    for _, row in df.iterrows():
        tower = f"Tower-{str(row['Tower Name']).strip()}"
        floor = str(row["Floor Number"]).strip()
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
    """Save JSON output and allow user to download it."""
    json_filename = "converted_data.json"
    with open(json_filename, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, indent=2)
    files.download(json_filename)  # Automatically download JSON file

def main():
    """Main function to run script."""
    print("📂 Select an Excel file OR enter the file path manually.")
    
    file_path = upload_file()
    
    if not file_path:
        print("No file uploaded. Exiting.")
        return
    
    print(f"Processing file: {file_path}")
    data = process_excel_data(file_path)

    if data:
        save_to_json(data)

if __name__ == "__main__":
    main()
