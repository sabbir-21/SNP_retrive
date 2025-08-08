import requests
import pandas as pd

# Base API endpoint
filename = "tm6sf2" #@param {type:"string"}
excel_path = f"{filename}_snp_informations.xlsx" 
base_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"

# Query filters and their human-readable labels
filters = {
    "All": f"{filename}[All Fields]",
    "3 prime UTR variant": f'{filename}[gene] AND "3 prime UTR variant"[Function Class]',
    "5 prime UTR variant": f'{filename}[gene] AND "5 prime UTR variant"[Function Class]',
    "missense variant": f'{filename}[All Fields] AND missense variant[Function_Class]',
    "synonymous": f'{filename}[All Fields] AND synonymous variant[Function_Class]',
    "intron": f'{filename}[All Fields] AND intron variant[Function_Class]',
    "snp_somatic": f'{filename}[All Fields] AND snp_snp_somatic[sb]',
    "Somatic + Missense": f'{filename}[All Fields] AND (missense variant[Function_Class] AND snp_snp_somatic[sb])'
}

results = []

# Query the API for each filter
for label, term in filters.items():
    response = requests.get(base_url, params={
        "db": "snp",
        "term": term,
        "retmode": "json"
    })
    data = response.json()
    count = int(data['esearchresult']['count'])
    results.append((label, count))

# Calculate percentage with respect to "All"
total_all = dict(results)["All"]
df = pd.DataFrame(results, columns=["Category", "Count"])
df["Percentage"] = df["Count"] / total_all * 100
df["Percentage"] = df["Percentage"].map(lambda x: f"{x:.3f}%")

# Save to Excel
df.to_excel(excel_path, index=False)