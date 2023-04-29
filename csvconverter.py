import os
import sys
import subprocess

def install_dependencies():
    dependencies = ['pandas', 'openpyxl']
    for dependency in dependencies:
        try:
            __import__(dependency)
        except ImportError:
            print(f"{dependency} not found, trying to install automatically...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", dependency])
                print(f"{dependency} installed successfully.")
            except subprocess.CalledProcessError as e:
                print(f"Error installing {dependency}, please install manually.")
                sys.exit(1)

install_dependencies()

import pandas as pd

def remove_quotes(s):
    return s.strip('"')

def xlsx_to_csv(input_dir, output_dir):
    if not input_dir.endswith(os.path.sep):
        input_dir += os.path.sep

    if not output_dir.endswith(os.path.sep):
        output_dir += os.path.sep

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.lower().endswith('.xlsx'):
                input_filepath = os.path.join(root, file)
                print(f"Processing {input_filepath}")
                
                output_filename = os.path.splitext(file)[0] + '.csv'
                output_filepath = os.path.join(output_dir, output_filename)
                
                # Read the first worksheet of the Excel file
                df = pd.read_excel(input_filepath, sheet_name=0).fillna('')

                # Sanitize the data by replacing problematic characters
                df = df.applymap(lambda x: str(x).replace('"', '""').replace('\\', '\\\\').replace("'", "''"))

                # Write the CSV file
                df.to_csv(output_filepath, index=False, quoting=1)  # quoting=1 means to quote all non-numeric values

if __name__ == '__main__':
    input_directory = remove_quotes(input("Enter the input directory: "))
    output_directory = remove_quotes(input("Enter the output directory: "))
    xlsx_to_csv(input_directory, output_directory)
