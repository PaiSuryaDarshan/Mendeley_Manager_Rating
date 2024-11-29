import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl

def UpdateCSV():
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

    if __name__ == "__main__":
        print(df)

UpdateCSV()