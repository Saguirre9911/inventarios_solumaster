import os

import openpyxl
import pandas as pd

# Specify the path to the Excel file
file_path = "docs/"


# Get the list of files in the docs folder
files = os.listdir(file_path)
# Find the Excel file in the list
for file in files:
    if file.endswith(".XLS"):
        file_path += file
        break
else:
    print("No Excel file found in the docs folder.")
    # Add appropriate error handling or exit the program if needed

# Reading the CSV file with tab as a delimiter and setting the encoding
# UTF-8 is a good starting point; adjust if you know the specific encoding
workbook = pd.read_csv(file_path, delimiter="\t", encoding="latin-1", skiprows=2)

# Check the first few rows to ensure it reads correctly
print(workbook.head())

# workbook = pd.read_excel(file_path, engine="openpyxl")

# Especifica la ruta y el nombre del archivo de salida
output_file_path = "docs/Excel_enviar.xlsx"

# Guardar el DataFrame en un archivo .xlsx con un nombre de hoja específico
with pd.ExcelWriter(output_file_path, engine="openpyxl") as writer:
    workbook.to_excel(writer, index=False, sheet_name="Excel_enviar")
