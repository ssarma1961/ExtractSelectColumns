"""Program to extract desired columns from a csv/excel file and sort on a chosen column"""


# Pre-requisite: Install 'openpyxl' modules to interact with Excel files

import pandas as pd
import re
import sys


# Make sure 'openpyxl' is installed
fn = input("Path and File name to extract the desired columns from xlsx/csv file: ")
def get_file_extension(fn):
    match = re.search(r'\.([a-zA-Z0-9]+)$', fn)
    if match:
        return match.group(1)  # Returns the extension without the dot
    else:
        return None

# Example usage

extension = get_file_extension(fn)
if extension:
    print(f"The file extension is: {extension}")
else:
    print("No valid file extension found.")
    sys.exit(0)
if extension=="xlsx" :
    sn= input("Input sheet name to read")
    try:
        df = pd.read_excel(io=fn,engine='openpyxl',sheet_name=sn)
        #print(df)
    except FileNotFoundError:
        print("Error: The specified file was not found. Please check the file path and try again.")
        sys.exit(0)
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        sys.exit(0)
elif extension=="csv":
    print ("input data is csv file")
    df = pd.read_csv(fn)
    df.to_excel('file.xlsx', index=False)  # convert input csv to xlsx format
else:
    sys.exit(0)

# Get all column values
listofcolumns = []
selectedcolumns = []
for column in df.columns:
    print("Here are all the columns available:", column)
    listofcolumns.append(column)

print(listofcolumns)

for cols in listofcolumns:
    selected = input(f"Extract this column {cols}? ")
    if selected.upper() == "Y":  # Ensure the comparison is case-insensitive
        selectedcolumns.append(cols)

print(selectedcolumns)
df2 = df[selectedcolumns]  # new dataframe

sc = input("Key in column title to sort on: ")
df2 = df2.sort_values(sc)

outfile = input("Output file name including extension xlsx: ")
excel_file = pd.ExcelWriter(outfile, engine='openpyxl')
if extension=="xlsx":
    df2.to_excel(excel_writer=excel_file,sheet_name=sn,columns=selectedcolumns, index=False)
else:
    df2.to_excel(excel_writer=excel_file,columns=selectedcolumns, index=False)  # if input file is csv
excel_file.save()  # Save the file """