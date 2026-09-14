import pandas as pd
import re

# Input and output files
input_file = r"C:\Users\imran\Desktop\amazon-scrapper\spoon\spoon.xlsx"
output_file = r"C:\Users\imran\Desktop\amazon-scrapper\spoon\Amazon_spoon_Cleaned.xlsx"

# Read dataset
df = pd.read_excel(input_file)

# Clean text columns
for col in df.select_dtypes(include="object").columns:
    df[col] = (
        df[col].astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )

# Convert Price to USD
def convert_to_usd(value):
    if pd.isna(value):
        return ""

    value = str(value).strip().replace(",", "")

    match = re.search(r"(\d+(?:\.\d+)?)", value)
    if not match:
        return ""

    amount = float(match.group(1))

    # PKR to USD: 1 USD = 280 PKR
    if re.search(r"PKR|Rs\.?|Rupees?", value, re.IGNORECASE):
        return round(amount / 280, 2)

    # Already USD
    return round(amount, 2)


if "Price" in df.columns:
    df["Price"] = df["Price"].apply(convert_to_usd)

# Clean Rating
if "Rating" in df.columns:
    df["Rating"] = (
        df["Rating"].astype(str)
        .str.extract(r"(\d+(?:\.\d+)?)")[0]
    )
    df["Rating"] = pd.to_numeric(df["Rating"], errors="coerce")

# Clean Review Count
if "Review Count" in df.columns:
    df["Review Count"] = (
        df["Review Count"].astype(str)
        .str.replace(",", "", regex=False)
        .str.extract(r"(\d+)")[0]
    )
    df["Review Count"] = pd.to_numeric(df["Review Count"], errors="coerce")

# Clean Bought in Past Month
if "Bought in past month" in df.columns:
    df["Bought in past month"] = (
        df["Bought in past month"].astype(str)
        .str.replace(",", "", regex=False)
        .str.extract(r"(\d+)")[0]
    )
    df["Bought in past month"] = pd.to_numeric(
        df["Bought in past month"], errors="coerce"
    )

# Clean ASIN
if "ASIN" in df.columns:
    df["ASIN"] = df["ASIN"].astype(str).str.strip().str.upper()

# Remove duplicates
if "ASIN" in df.columns:
    df = df.drop_duplicates(subset="ASIN")
else:
    df = df.drop_duplicates()

# Replace missing values with blank
df = df.fillna("")

# Save cleaned + USD converted dataset
df.to_excel(output_file, index=False)

print("Dataset cleaned and prices converted to USD successfully!")
print(output_file)