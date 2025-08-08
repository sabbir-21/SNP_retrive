from bs4 import BeautifulSoup
import requests
from openpyxl import load_workbook
import re

# Open the existing Excel file
filename = "tm6sf2" #@param {type:"string"}
name = f'{filename}_missense.xlsx'
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

        # If we found a Missense Variant, write it into column B-D and process E-G
        if first_missense_change:
            worksheet[f'B{idx}'] = m_type
            worksheet[f'C{idx}'] = first_missense_change_location
            worksheet[f'D{idx}'] = first_missense_change

            # Column E: extract last part after ':' (e.g., Met1Arg)
            match = re.search(r'p\.([A-Za-z]+\d+[A-Za-z]+)', first_missense_change_location)
            if match:
                aa_change = match.group(1)
                worksheet[f'E{idx}'] = aa_change

                # Column F: extract number from Met1Arg
                number_match = re.search(r'\d+', aa_change)
                if number_match:
                    worksheet[f'F{idx}'] = number_match.group()

                # Column G: build M1R from D and number
                d_col_value = worksheet[f'D{idx}'].value  # e.g., "M (Met) > R (Arg)"
                compact_match = re.match(r'([A-Z]) \(.*\) > ([A-Z]) \(.*\)', d_col_value)
                if compact_match and number_match:
                    from_aa = compact_match.group(1)
                    to_aa = compact_match.group(2)
                    num = number_match.group()
                    worksheet[f'G{idx}'] = f"{from_aa}{num}{to_aa}"
        else:
            worksheet[f'B{idx}'] = 'No Missense Variant Found'

# Save the changes to the same Excel file
    workbook.save(name)
