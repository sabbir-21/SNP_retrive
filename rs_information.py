#@title **Retrieve SNP Information from NCBI**
#@markdown #<font color='orange'>**1**
import requests
import pandas as pd
import time

# ====== CONFIGURATION ======
filename = "tm6sf2"  #@param {type:"string"}
excel_path = f"{filename}_snp_informations.xlsx"
base_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"

# ====== FILTER QUERIES ======
# Using valid NCBI SNP field names
filters = {
    "All": f"{filename}[All Fields]",
    "3 prime UTR variant": f'{filename}[Gene] AND "3 prime UTR variant"[Function_Class]',
    "5 prime UTR variant": f'{filename}[Gene] AND "5 prime UTR variant"[Function_Class]',
    "missense variant": f'{filename}[All Fields] AND missense variant[Function_Class]',
    "synonymous": f'{filename}[All Fields] AND synonymous variant[Function_Class]',
    "intron": f'{filename}[All Fields] AND intron variant[Function_Class]',
    "snp_somatic": f'{filename}[All Fields] AND snp_snp_somatic[sb]',
    "Somatic + Missense": f'{filename}[All Fields] AND (missense variant[Function_Class] AND snp_snp_somatic[sb])'
}

results = []

# ====== QUERY NCBI ======
for label, term in filters.items():
    time.sleep(5)
    response = requests.get(base_url, params={
        "db": "snp",
        "term": term,
        "retmode": "json"
    })

    try:
        data = response.json()
    except Exception:
        print(f"[{label}] Invalid JSON response:")
        print(response.text)
        results.append((label, 0))
        continue

    count = int(data.get("esearchresult", {}).get("count", 0))

    # Debug info if API did not return expected structure
    if "esearchresult" not in data:
        print(f"[{label}] Missing 'esearchresult' in response:")
        print(response.text)

    results.append((label, count))

# ====== PROCESS RESULTS ======
total_all = dict(results).get("All", 0)
df = pd.DataFrame(results, columns=["Category", "Count"])

if total_all > 0:
    df["Percentage"] = df["Count"] / total_all * 100
    df["Percentage"] = df["Percentage"].map(lambda x: f"{x:.3f}%")
else:
    df["Percentage"] = "0.000%"

# ====== SAVE TO EXCEL ======
df.to_excel(excel_path, index=False)

print(f"✅ Data saved to {excel_path}")
print(df)