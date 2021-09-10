import pandas as pd

ref = pd.read_excel('./Copy of 3207 Affected markets for a CMC Variation.xlsx', sheet_name="Variation 1",
                   skiprows=14, header=0)
search = pd.read_excel('./Register CMC-05-Affected documents 02.07.2021.xlsx', sheet_name="All docs",
                       header=0)
doc_type = "P3-05"

match_col = pd.Series(index=ref.index, dtype=str)
for index, ref_id in enumerate(ref['Registration ID']):
    matches = search.loc[(search['Registration ID'] == ref_id) & (search["Title"].str.contains(doc_type))]
    if not matches.empty:
        ref_nums = matches['Reference Number'].to_string(index=False).replace("\n", ", ")
    else:
        ref_nums = ""
    match_col[index] = ref_nums
ref[doc_type] = match_col



ref.to_excel("matches.xlsx")