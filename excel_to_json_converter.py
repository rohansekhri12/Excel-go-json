import pandas as pd
import json
import os
from google.colab import files  # Import for handling file uploads & downloads

# Function to upload Excel file
def upload_file():
    uploaded = files.upload()  # Opens file upload dialog in Colab
    for filename in uploaded.keys():
        return filename  # Return the uploaded file name

# Function to process Excel data and convert it to JSON
def process_excel_data(file_path):
    """Reads the Excel file and converts it into the required JSON format."""
    
    # Read Excel file
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return None

    # Ensure required columns exist
    required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
    for col in required_columns:
        if col not in df.columns:
            print(f"Error: Missing required column '{col}' in the Excel file.")
            return None

    # Dictionary to store processed data
    towers_dict = {}

    for _, row in df.iterrows():
        tower = f"Tower-{str(row['Tower Name']).replace(' ', '')}"
        floor = str(row["Floor Number"])
        companies = [company.strip() for company in str(row["Company Name(s)"]).split(',')]

        if tower not in towers_dict:
            towers_dict[tower] = {}

        if floor in towers_dict[tower]:
            towers_dict[tower][floor].extend(companies)
        else:
            towers_dict[tower][floor] = companies

    # Convert to required JSON structure
    towers_list = [{"data": [{"name": list(set(companies)), "floor": floor} for floor, companies in floors.items()], "tower": tower} for tower, floors in towers_dict.items()]
    
    return towers_list

# Function to save and download JSON file
def save_to_json(data):
    """Saves the processed data as JSON and prompts user to download."""
    json_filename = "converted_data.json"
    with open(json_filename, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, indent=2)
    
    print(f"✅ Data successfully saved as {json_filename}")
    files.download(json_filename)  # Auto-downloads JSON file

# Main function to execute script in Colab
def main():
    print("📂 Upload your Excel file")
    excel_file = upload_file()

    if not excel_file:
        print("❌ No file uploaded. Exiting...")
        return

    print(f"🔄 Processing file: {excel_file}")

    # Process Excel data
    json_data = process_excel_data(excel_file)

    if json_data:
        # Save and download JSON file
        save_to_json(json_data)

# Run the main function
main()
