import pandas as pd

# Load Excel file
df = pd.read_excel("incorporation_jan24_dec24.xlsx")

# Find incorporation date column (supports "IncorporationDate" and variants)
normalized_map = {str(col).strip().lower().replace(" ", ""): col for col in df.columns}
date_col = normalized_map.get("incorporationdate")

if date_col is None:
    raise KeyError(
        "Incorporation date column not found. "
        f"Available columns: {list(df.columns)}"
    )

df[date_col] = pd.to_datetime(
    df[date_col],
    dayfirst=True,
    errors='coerce'
)

df_2024 = df[df[date_col].dt.year == 2024]

df_2024.to_excel("only_2024_data.xlsx", index=False)

print(df_2024)
