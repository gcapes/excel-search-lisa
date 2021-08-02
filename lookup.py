import pandas as pd

ref = pd.read_excel('./Copy of 3207 Affected markets for a CMC Variation.xlsx', sheet_name="Variation 1",
                   skiprows=14, header=0)
search = pd.read_excel('./Register CMC-05-Affected documents 02.07.2021.xlsx', sheet_name="All docs",
                       header=0)
doc_type = "P3-03"
matches = search.loc[(search['Registration ID'].isin(ref['Registration ID']))
                     & (search["Title"].str.contains(doc_type))]

matches.to_excel(doc_type + "_matches.xlsx")