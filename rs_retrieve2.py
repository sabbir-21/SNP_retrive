import requests
from openpyxl import Workbook

filename = "tm6sf2" #@param {type:"string"} # output Excel filename without extension

# API URL and parameters
url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
params = {
    "db": "snp",
    "term": f"{filename}[All Fields] AND missense variant[Function_Class]",
    "retmax": 1000,
    "retmode": "json"
}

# Send request
response = requests.get(url, params=params)
response.raise_for_status()
data = response.json()

# Extract rs IDs from idlist, prepend "rs"
id_list = data.get("esearchresult", {}).get("idlist", [])
rs_values = ["rs" + rs_id for rs_id in id_list]

# Write to Excel
workbook = Workbook()
sheet = workbook.active

for rs_id in rs_values:
    sheet.append([rs_id])

workbook.save(f"{filename}_missense.xlsx")
print(f"Done. Saved {len(rs_values)} rs IDs to {filename}.xlsx")
