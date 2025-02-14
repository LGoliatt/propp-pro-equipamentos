import pandas as pd
import requests
from io import BytesIO

# Replace with your Google Sheet ID
sheet_id = '1D-BwWxNHHnIvcFT85EkLUhLT52fUH8KnDtiQiGTD4HA'
# Google Sheets export URL for Excel format
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
# Send request to download the file
response = requests.get(url)
response.raise_for_status()  # Raise an error if request fails

# Read the Excel file into Pandas
df = pd.read_excel(BytesIO(response.content), sheet_name=None)  # Read first sheet
# Display sheet names
print("Sheets found:", df.keys())

df_all = pd.DataFrame()
# Access individual sheets
for sheet_name, df1 in df.items():
    if 'Consolidado' not in sheet_name:
        print(f"Sheet: {sheet_name}")
        df2 = pd.read_excel(BytesIO(response.content), sheet_name=sheet_name, header=1)  # Read first sheet
        df2['PPG']=sheet_name
        df_all = pd.concat([df_all, df2])


cols=list(df_all.columns[-1:])+list(df_all.columns[:-1])

df_all=df_all[cols]
df_all = df_all[df_all['Descrição do Equipamento'].notna()]
df_all = df_all[df_all['Preço unitário'].notna()]

df_all.to_excel('propp-pro-equipamentos-todos.xlsx', index=False)
# filtros
df_all = df_all[df_all["Preço unitário"] >= 10000]
df_all.to_excel('propp-pro-equipamentos-filtrado.xlsx', index=False)
