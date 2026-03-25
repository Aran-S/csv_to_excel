import pandas as pd

# Load Excel file
df = pd.read_excel("incorporation_jan24_dec24.xlsx")

# Normalize column names to find incorporation date column
normalized_map = {
    str(col).strip().lower().replace(" ", ""): col 
    for col in df.columns
}

date_col = normalized_map.get("incorporationdate")

if date_col is None:
    raise KeyError(
        "Incorporation date column not found. "
        f"Available columns: {list(df.columns)}"
    )

# Convert to datetime
df[date_col] = pd.to_datetime(
    df[date_col],
    dayfirst=True,
    errors='coerce'
)

# Remove invalid dates (optional but recommended)
df = df.dropna(subset=[date_col])

# Filter only 2024 data
df_2024 = df[df[date_col].dt.year == 2024].copy()

# Format date to avoid #### in Excel
df_2024[date_col] = df_2024[date_col].dt.strftime('%d/%m/%Y')

# Save to Excel with proper column width
with pd.ExcelWriter("only_2024_data.xlsx", engine='xlsxwriter') as writer:
    df_2024.to_excel(writer, index=False, sheet_name='Sheet1')
    worksheet = writer.sheets['Sheet1']
    
    # Auto-adjust column width
    for i, col in enumerate(df_2024.columns):
        max_len = max(
            df_2024[col].astype(str).map(len).max(),
            len(col)
        ) + 2
        worksheet.set_column(i, i, max_len)

# Print output
print(df_2024)
