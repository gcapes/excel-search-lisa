import pandas as pd

ref = pd.read_excel('./Copy of 3207 Affected markets for a CMC Variation.xlsx', sheet_name="Variation 1",
                   skiprows=14, header=0)
search = pd.read_excel('./Register CMC-05-Affected documents 02.07.2021.xlsx', sheet_name="All docs",
                       header=0)
results = pd.DataFrame()
for id in ref['Registration ID']:
    matches = search.loc[(search['Registration ID'] == id) & (search["Title"].str.contains("P7-00D"))]
    if not matches.empty:
        results = results.append(matches)

results.to_excel("matches.xlsx")