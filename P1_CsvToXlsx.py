from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd
import P0_XmlToCsv as P0

# Update CSV
P0.UpdateCSV()

# Call CSV
df = pd.read_csv("new.csv")

# Append values to Main Excel
filename = "Reading_Ratings.xlsx"

workbook = load_workbook(filename) 
  
# Get the first sheet 
sheet = workbook.worksheets[0] 

# Grab all tables on the sheet
tables = []

for table in sheet.tables.values():
    tables.append(table)

# Grab loc of table of interest (TOI)
toi = tables[0]

# Adjust table size to match csv | (+1) to value to prevent overwriting of Header Row
Rows_of_Dataset = len(df)
toi.ref = f"A1:F{Rows_of_Dataset + 1}"

workbook.save(filename)

# Create a list to store the values 
titles = [] 
  
# Iterate through columns 
for column in sheet.iter_cols(): 
    # Get the value of the first cell in the 
    # column (the cell with the column name) 
    column_name = column[0].value 
    # Check if the column is the "Name" column 
    if column_name == "Title": 
        # Iterate over the cells in the column 
        for cell in column: 
            # Add the value of the cell to the list 
            titles.append(cell.value) 

for index, title in enumerate(df['Title']):
    if title not in titles:
        print("tr")