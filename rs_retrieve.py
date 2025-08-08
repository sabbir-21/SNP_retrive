from bs4 import BeautifulSoup
from openpyxl import Workbook

filename = "rs58"  # Base name for output Excel file

# Step 1: Parse the input HTML file
with open("58.html", "r", encoding="utf-8") as file:
    content = file.read()
soup = BeautifulSoup(content, 'html.parser')

# Step 2: Find all relevant <a> tags inside <div class="rprt">
rprt_divs = soup.find_all("div", class_="rprt")

# Step 3: Extract rsIDs and store them
rs_values = []
for div in rprt_divs:
    a_tag = div.find("a", href=True)
    if a_tag and "/snp/" in a_tag['href']:
        numeric_value = a_tag['href'].split("/snp/")[1]
        rs_values.append('rs' + numeric_value)

# Step 4: Save the rsIDs to Excel
workbook = Workbook()
sheet = workbook.active
for rs_id in rs_values:
    sheet.append([rs_id])

workbook.save(f"{filename}.xlsx")
print("Done")