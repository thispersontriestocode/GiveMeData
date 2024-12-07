import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

# Function to extract specific columns from a single CSV file
def extract_columns_from_csv(file_path, encodings):
    last_exception = None
    required_columns = ['Name', 'Device Type', 'Module Name', 'Vendor Name', 'Serial Number']

    for encoding in encodings:
        for header_option in [2, 0]:  # Try both header=2 and header=0
            try:
                # Load the CSV file with the specified header option
                df = pd.read_csv(file_path, header=header_option, encoding=encoding)

                # Check if the required columns exist in the DataFrame
                if all(column in df.columns for column in required_columns):
                    return df[required_columns]  # Extract and return the required columns
                else:
                    print(f"Columns not found in header={header_option} for file: {file_path}")

            except Exception as e:
                last_exception = e  # Store the last exception for debugging
                print(f"Failed with header={header_option} and encoding={encoding} for file {file_path}: {e}")

    print(f"Error processing file {file_path}: {last_exception}")
    return None

# Function to save the DataFrame to an Excel file
def save_to_excel(df, output_file_path):
    try:
        df.to_excel(output_file_path, index=False)  # Save as Excel file
        print(f"Data saved to {output_file_path}")
    except Exception as e:
        print(f"Error saving file {output_file_path}: {e}")

# Function to select a single CSV file
def select_file(encodings):
    file_path = filedialog.askopenfilename(title="Select a CSV file", filetypes=[("CSV files", "*.csv")])
    if file_path:
        extracted_data = extract_columns_from_csv(file_path, encodings)
        if extracted_data is not None:
            output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                             filetypes=[("Excel files", "*.xlsx")])
            if output_file_path:
                save_to_excel(extracted_data, output_file_path)  # Call save_to_excel to save the extracted data

# Function to select a folder and process all CSV files in it
def select_folder(encodings):
    folder_path = filedialog.askdirectory(title="Select a Folder")
    if folder_path:
        output_folder = filedialog.askdirectory(title="Select Output Folder")  # Ask for output folder
        if output_folder:
            for filename in os.listdir(folder_path):
                if filename.endswith('.csv'):
                    file_path = os.path.join(folder_path, filename)
                    extracted_data = extract_columns_from_csv(file_path, encodings)
                    if extracted_data is not None:
                        # Create output file path
                        output_file_name = os.path.splitext(filename)[0] + '.xlsx'  # Keep the original filename
                        output_file_path = os.path.join(output_folder, output_file_name)
                        save_to_excel(extracted_data, output_file_path)  # Save each extracted data to its own file

# Create the main window
root = tk.Tk()
root.title("CSV Column Extractor")

# List of encodings to try
encodings_to_try = ['utf-16', 'utf-8', 'iso-8859-1', 'windows-1252']

# Create buttons for file and folder selection
btn_file = tk.Button(root, text="Select CSV File", command=lambda: select_file(encodings_to_try))
btn_file.pack(pady=10)

btn_folder = tk.Button(root, text="Select Folder of CSV Files", command=lambda: select_folder(encodings_to_try))
btn_folder.pack(pady=10)

# Start the GUI event loop
root.mainloop()