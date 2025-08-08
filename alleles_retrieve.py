from bs4 import BeautifulSoup
import requests
from openpyxl import load_workbook
import re, time

# Open the existing Excel file
filename = "tm6sf2"  #@param {type:"string"}
name = f"{filename}_missense.xlsx"

with open(name, 'rb') as tf:
    workbook = load_workbook(tf)
worksheet = workbook['Sheet']

# Store data to be added with their corresponding row index
new_data = []

count = 0
for idx, cell in enumerate(worksheet['A'], start=1):
    rsid = str(cell.value).strip()
    if rsid and rsid.lower() != 'none':
        count += 1
        print(f"{count}. {rsid}")
        url = f"https://www.ncbi.nlm.nih.gov/snp/{rsid}#variant_details"
        time.sleep(0.7)
        rr = requests.get(url)
        soup = BeautifulSoup(rr.text, 'html.parser')

        # Get all <dl> blocks
        dl_tags = soup.find_all("dl")

        # Initialize
        position = None
        alleles = None

        for dl in dl_tags:
            dts = dl.find_all("dt")
            dds = dl.find_all("dd")

            for dt, dd in zip(dts, dds):
                key = dt.text.strip()
                val = dd.text.strip()

                if key == "Position":
                    position = val.replace('\n', '').strip()
                elif key == "Alleles":
                    alleles = val.replace('\n', '').strip()

        if position and alleles:
            chr_match = re.search(r'chr(\d+|X|Y|MT):(\d+)', position)
            if chr_match:
                chrom = chr_match.group(1)
                pos = chr_match.group(2)

                # Split multiple alleles: A>C or C>A / C>T
                parts = re.split(r'\s*/\s*', alleles)
                formatted_alleles = []
                for allele in parts:
                    allele = allele.replace(">", ">").strip()
                    if ">" in allele:
                        ref, alt = allele.split(">")
                        formatted = f"{chrom},{pos},{ref},{alt}"
                        formatted_alleles.append(formatted)
                
                # Store row index, position, and up to two alleles
                new_data.append((idx, position, formatted_alleles[:2]))

# Write data to columns B, C, and D
for idx, position, formatted_alleles in new_data:
    # Write position to column B
    worksheet[f'B{idx}'] = position
    
    # Write first allele to column C
    if len(formatted_alleles) > 0:
        worksheet[f'C{idx}'] = formatted_alleles[0]
    
    # Write second allele to column D, if it exists
    if len(formatted_alleles) > 1:
        worksheet[f'D{idx}'] = formatted_alleles[1]

# Save
workbook.save(name)
print(f"\n✅ Done! {len(new_data)} rows updated in columns B, C, and D.")