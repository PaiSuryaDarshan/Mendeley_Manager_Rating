import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl

# Parse XML into Element Tree
# EndNote XML V8
tree = ET.parse("export.xml")
root = tree.getroot()

# Root - Test
print(root[0].tag)
print(len(root[0]))


# List collecting data
cols = ["S.No", "Year", "Title", "Author", "DOI"]
rows = []

# Loop over Root Children till core of interest is reached
for record in root[0]:
    row = []
    Sno = ""
    year = ""
    title = "" 
    author = "" 
    doi = ""
    for items in record:
        if items.tag == 'electronic-resource-num':
            doi = 'https://doi.org/' + items.text
        for titles in items:
            if titles.tag == 'title':
                title = titles.text
                pass

            if titles.tag == 'year':
                year = titles.text
                pass

            for title_data in titles:
                if title_data.tag == 'author':
                    author = title_data.text.strip()
                    break
    row.append(Sno)
    row.append(year)
    row.append(title)
    row.append(author)
    row.append(doi)

    rows.append(row)

df = pd.DataFrame(rows, columns=cols)
df.to_csv("new.csv", index=False)
print(df)

# Append values to Main Excel
filename = "Reading_Ratings.xlsx"

workbook = openpyxl.load_workbook(filename) 
  
# Get the first sheet 
sheet = workbook.worksheets[0] 

# Create a list to store the values 
names = [] 
  
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
            names.append(cell.value) 

for index, title in enumerate(df['Title']):
    if title in names:
        pass
    else:
# ! Fix this - Last step, adding non-duplicate items to table... ERRO: df.values returns 'munpy.ndarray' instead of 'list'
        print(index)
        new_value = df.values[index]
        print(new_value)
        sheet.append(new_value)