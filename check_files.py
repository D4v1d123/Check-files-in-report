# Author: David Alejandro Guerrero Amador.
# Check the files that are in the report but not in the files folder.
import os
import pandas
from numpy import isnan

# The content is taken from a spreadsheet and the file IDs mentioned in report are 
# extracted.
def extract_id_files_in_report(path, worksheet):
    data_frame = pandas.read_excel(path, worksheet)
    id_files_report = {}
    
    for _, row_values in data_frame.iterrows():
        if not(isnan(row_values["ID"])):
            id = int(row_values["ID"])
            id_files_report.add(id)
        
    return id_files_report

# The IDs are extracted from the name of each file to be update.  
def extract_id_files(path):
    files = os.listdir(path)
    id_files = {}
    
    for file in files:
        id = int(file.split()[0])
        id_files.add(id)
    
    return id_files


report_path = r"D:\Users\User\Documents\BACKUP 04-03-2024\David\Reportes\Report.xlsx"
files_path = r"D:\Users\User\Documents\BACKUP 04-03-2024\David\Salud.SIS\SUBIR ARCHIVOS"

id_files_report = extract_id_files_in_report(report_path, "Hoja1")
id_files = extract_id_files(files_path)

files_in_report = sorted(list(id_files_report.difference(id_files)))
files_in_folder = sorted(list(id_files.difference(id_files_report)))

print(f"Files that are in the report but not in the files folder to upload:\n{files_in_report}\n")
print(f"Files that are in the files folder to upload but not in the report:\n{files_in_folder}")
