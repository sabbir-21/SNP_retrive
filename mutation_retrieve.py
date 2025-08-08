from bs4 import BeautifulSoup
import requests
from openpyxl import load_workbook

# Open the existing Excel file
name = 'rs58.xlsx'
with open(name, 'rb') as tf:
    workbook = load_workbook(tf)
worksheet = workbook['Sheet']

# Iterate over column A in the 'Predict SNP' sheet
count = 0
for idx, cell in enumerate(worksheet['A'], start=1):
    
    val = str(cell.value)
    if val != 'None':
        count += 1
        print(count)
        url = f"https://www.ncbi.nlm.nih.gov/snp/{val}#variant_details"
        rr = requests.get(url)
        soup = BeautifulSoup(rr.text, 'html.parser')

        # Find all rows in the HTML that are in the table
        table_rows = soup.find_all("tr")
        first_missense_change = None

        # Search for the first Missense Variant
        for row in table_rows:
            cells = row.find_all("td")
            
            if cells:
                # Check if the last cell contains "Missense Variant"
                if "Missense Variant" in cells[-1].get_text(strip=True):
                    # Extract the amino acid change from the third column (2nd <td>)
                    m_type = cells[0].get_text(strip=True)
                    first_missense_change_location = cells[1].get_text(strip=True)
                    first_missense_change = cells[2].get_text(strip=True)
                    break  # Stop after finding the first Missense Variant

        # If we found a Missense Variant, write it into column B of the same row
        if first_missense_change:
            worksheet[f'B{idx}'] = m_type
            worksheet[f'C{idx}'] = first_missense_change_location
            worksheet[f'D{idx}'] = first_missense_change
        else:
            worksheet[f'B{idx}'] = 'No Missense Variant Found'

# Save the changes to the same Excel file
    workbook.save(name)
