import pandas as pd

# Load Excel file
df = pd.read_excel("incorporation_jan24_dec24.xlsx")

# Convert to datetime (day first format)
df['Incorporation Date'] = pd.to_datetime(
    df['Incorporation Date'],
    dayfirst=True,
    errors='coerce'
)

# Filter only 2024 data
df_2024 = df[df['Incorporation Date'].dt.year == 2024]

# Save result
df_2024.to_excel("only_2024_data.xlsx", index=False)

print(df_2024)
