import pdfplumber
import pandas as pd
import os
import re
import glob

pdf_folder = "."
output_excel = "output2.xlsx"

all_rows = []
times = []

for pdf_file in glob.glob(os.path.join(pdf_folder, "*.pdf")):
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            times += tables[2][1]

all_rows = []

for entry in times:
    if not entry:
        continue
    # replace newlines with spaces
    entry = entry.replace('\n', ' ')
    # split by whitespace
    parts = entry.split()
    # every 3 parts is one row: start, end, time
    for i in range(0, len(parts), 3):
        start = parts[i]
        end = parts[i+1]
        time = parts[i+2]
        all_rows.append({
            "Start": start,
            "End": end,
            "Time": time
        })

df = pd.DataFrame(all_rows)
df.to_excel(output_excel, index=False)
print(f"Saved {len(df)} entries to {output_excel}")