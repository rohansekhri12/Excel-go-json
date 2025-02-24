
# import pandas as pd
# import json
# import os
# import tkinter as tk
# from tkinter import filedialog

# def select_excel_file():
#     """Opens a file dialog to select the Excel file."""
#     root = tk.Tk()
#     root.withdraw()  # Hide the main window
#     file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
#     return file_path

# def format_tower_name(tower):
#     """Formats the tower name to ensure it follows 'Tower-X' format."""
#     return f"Tower-{tower}".replace(" ", "")  # Ensuring no spaces

# def process_excel_data(file_path):
#     """Reads the Excel file and processes it into the required JSON format where 'data' comes first, then 'tower'."""
    
#     # Check if file exists
#     if not os.path.exists(file_path):
#         print("Error: File does not exist.")
#         return None
    
#     try:
#         # Read the Excel file
#         df = pd.read_excel(file_path)
#     except Exception as e:
#         print(f"Error reading the Excel file: {e}")
#         return None

#     # Ensure the required columns exist
#     required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
#     for col in required_columns:
#         if col not in df.columns:
#             print(f"Error: Missing required column '{col}' in the Excel file.")
#             return None

#     # Initialize a dictionary to store towers
#     towers_dict = {}

#     for _, row in df.iterrows():
#         tower = format_tower_name(str(row["Tower Name"]))  # Formatting the tower name
#         floor = str(row["Floor Number"])  # Convert to string to maintain consistency
#         companies = [company.strip() for company in str(row["Company Name(s)"]).split(',')]

#         # If tower not in dictionary, initialize it
#         if tower not in towers_dict:
#             towers_dict[tower] = {}

#         # If floor already exists, append the company names
#         if floor in towers_dict[tower]:
#             towers_dict[tower][floor].extend(companies)
#         else:
#             towers_dict[tower][floor] = companies

#     # Convert dictionary to required JSON format
#     towers_list = []
#     for tower, floors in towers_dict.items():
#         data_list = [{"name": list(set(companies)), "floor": floor} for floor, companies in floors.items()]
#         towers_list.append({"data": data_list, "tower": tower})  # **Ensuring 'data' comes before 'tower'**

#     return towers_list

# def save_to_json(data):
#     """Asks the user for a JSON filename and saves the processed data."""
#     json_filename = input("Enter the name of the JSON file (without extension): ").strip()
    
#     if not json_filename:
#         json_filename = "tower_data_from_excel"

#     json_filename = f"{json_filename}.json"  # Ensure it has a .json extension

#     with open(json_filename, 'w', encoding='utf-8') as json_file:
#         json.dump(data, json_file, indent=2)
    
#     print(f"Data successfully saved to {json_filename}")

# def main():
#     """Main function to execute the script."""
#     print("Select the Excel file or manually enter the path.")
    
#     # Ask the user to select a file using a file dialog
#     file_path = select_excel_file()
    
#     if not file_path:
#         print("No file selected. Exiting...")
#         return

#     print(f"Processing file: {file_path}")

#     # Process the Excel data
#     data = process_excel_data(file_path)

#     if data is not None:
#         # Ask user for JSON filename and save data
#         save_to_json(data)

# if __name__ == "__main__":
#     main()

    
from google.colab import files
import pandas as pd
import json
import os

def select_excel_file():
    """Allows user to upload an Excel file in Google Colab."""
    uploaded = files.upload()  # Opens file upload dialog
    for filename in uploaded.keys():
        print(f"File {filename} uploaded successfully.")
        return filename  # Return the name of the uploaded file

def format_tower_name(tower):
    """Formats the tower name to ensure it follows 'Tower-X' format."""
    return f"Tower-{tower}".replace(" ", "")

def process_excel_data(file_path):
    """Reads the Excel file and processes it into JSON format."""
    
    # Check if file exists
    if not os.path.exists(file_path):
        print("Error: File does not exist.")
        return None

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return None

    required_columns = ["Tower Name", "Floor Number", "Company Name(s)"]
    for col in required_columns:
        if col not in df.columns:
            print(f"Error: Missing required column '{col}' in the Excel file.")
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

    print(f"Data successfully saved to {json_filename}")

    # Provide download link
    files.download(json_filename)
    print("Download started!")

def main():
    """Main function to execute the script."""
    print("Upload an Excel file to process.")
    
    file_name = select_excel_file()
    
    if not file_name:
        print("No file selected. Exiting...")
        return

    print(f"Processing file: {file_name}")

    data = process_excel_data(file_name)

    if data is not None:
        save_to_json(data)

if __name__ == "__main__":
    main()
