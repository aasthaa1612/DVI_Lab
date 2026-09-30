import pandas as pd
import json
import numpy as np
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

class NpEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (np.integer,)):
            return int(obj)
        if isinstance(obj, (np.floating,)):
            return float(obj)
        return super().default(obj)

def jdumps(obj):
    return json.dumps(obj, cls=NpEncoder)

df = pd.read_excel('ecourts_dataset.xlsx', sheet_name='Case Data')

col = 'Case Type/Case Number/Case Year'
df['Case Code']   = df[col].str.extract(r'^([A-Z0-9]+)')
df['Case Year']   = df[col].str.extract(r'/(\d{4})$').astype(float)
df['Case Number'] = df[col].str.extract(r'/(\d+)/\d{4}$')
df['Search Label'] = df['_search'].astype(str).str.strip("'\" ")

party_col = 'Petitioner Name Versus Respondent Name'
df['Petitioner'] = df[party_col].str.split(r'\\nVersus\\n|Versus').str[0].str.strip()
df['Respondent'] = df[party_col].str.split(r'\\nVersus\\n|Versus').str[-1].str.strip()

by_type = df['Search Label'].value_counts()
print('BY_TYPE:' + jdumps({str(k):int(v) for k,v in by_type.items()}))

year_counts = df['Case Year'].dropna().astype(int).value_counts().sort_index()
print('BY_YEAR:' + jdumps({str(k):int(v) for k,v in year_counts.items()}))

pet = (df['Petitioner'].dropna().str.upper().str.strip()
       .replace('', np.nan).dropna().value_counts().head(15))
print('TOP_PET:' + jdumps({str(k):int(v) for k,v in pet.items()}))

pivot = df.dropna(subset=['Case Year']).copy()
pivot['Case Year'] = pivot['Case Year'].astype(int)
pivot_table = pivot.groupby(['Case Year','Search Label']).size().unstack(fill_value=0)
print('STACKED:' + jdumps({str(yr): {str(c): int(val) for c, val in row.items()} for yr, row in pivot_table.iterrows()}))

total = len(df)
bail_cases = int(len(df[df['Search Label'].str.contains('Bail', case=False)]))
order_cases = int(len(df[df['Search Label'].str.contains('09/2026', case=False)]))
party_cases = int(len(df[df['Search Label'].str.contains('Uttrakhand', case=False)]))
print('STATS:' + jdumps({'total': total, 'bail': bail_cases, 'order': order_cases, 'party': party_cases}))

codes = df['Case Code'].value_counts().head(10)
print('CODES:' + jdumps({str(k):int(v) for k,v in codes.items()}))

has_order = int(df['Order Date'].notna().sum())
no_order = int(df['Order Date'].isna().sum())
print('ORDER_STATUS:' + jdumps({'Has Order': has_order, 'Pending/No Order': no_order}))

sorted_years = sorted(year_counts.index.tolist())
year_data = {str(int(y)): int(year_counts[y]) for y in sorted_years}
print('YEAR_SORTED:' + jdumps(year_data))

resp = (df['Respondent'].dropna().str.upper().str.strip()
        .replace('', np.nan).dropna().value_counts().head(10))
print('TOP_RESP:' + jdumps({str(k):int(v) for k,v in resp.items()}))

print('DONE')
