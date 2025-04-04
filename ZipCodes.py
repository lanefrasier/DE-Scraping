import pandas as pd

# Define file paths and sheet names
input_file = "InputData.xlsx"
zip_file = "C:/Users/LFrasier/NRG Energy, Inc/Digital Operations-NRG365 - General/Process Documents/USN State Rate Board Documents/USN SS Logins.xlsx"
sheet_input = "USN_Pricing_Master"
sheet_zip = "ZipCodes by State"

# Read the InputData file into a DataFrame
df_input = pd.read_excel(input_file, sheet_name=sheet_input, dtype={"LDC": str})

# Read the USN SS Logins file into a DataFrame, ensuring the ZipCode 1 column is read as a string
df_zip = pd.read_excel(zip_file, sheet_name=sheet_zip, dtype={"LDC": str, "ZipCode 1": str})

# Merge df_input with df_zip based on the 'LDC' column to bring in the zip codes
df_merged = df_input.merge(df_zip[["LDC", "ZipCode 1"]], on="LDC", how="left")

# Rename the merged zip code column to 'Zip Code'
df_merged.rename(columns={"ZipCode 1": "Zip Code"}, inplace=True)

# Convert the "Zip Code" column to string, replace "nan" with "00000", then pad with zeros to ensure 5 digits
df_merged["Zip Code"] = df_merged["Zip Code"].astype(str)
df_merged["Zip Code"] = df_merged["Zip Code"].replace("nan", "00000")
df_merged["Zip Code"] = df_merged["Zip Code"].str.zfill(5)

# Mismatches of Name vs Popup
df_merged = df_merged.map(lambda x: x.replace("&", "and") if isinstance(x, str) else x)
df_merged = df_merged.map(lambda x: x.replace("PECO", "Peco Electricity") if isinstance(x, str) else x)
df_merged = df_merged.map(lambda x: x.replace("Columbia Gas of PA", "Columbia Gas of Pennsylvania") if isinstance(x, str) else x)
df_merged = df_merged.map(lambda x: x.replace("Nicor", "Nicor Gas") if isinstance(x, str) else x)
df_merged = df_merged.map(lambda x: x.replace("Northshore", "North Shore") if isinstance(x, str) else x)
df_merged = df_merged.map(lambda x: x.replace("Michigan Consumers", "Consumers Energy") if isinstance(x, str) else x)
df_merged = df_merged.map(lambda x: x.replace("Michigan Consolidated", "DTE Gas") if isinstance(x, str) else x)

# Save the merged DataFrame as a new file "InputData_ZipCodes.xlsx" with the same sheet name as the original
output_file = "InputData_ZipCodes.xlsx"
df_merged.to_excel(output_file, sheet_name=sheet_input, index=False)

# Build a dictionary with unique entries; keys are (ZipCode, LDC, Commodity) tuples.
ZipsDict = {}

# Iterate through each row in the merged DataFrame
for index, row in df_merged.iterrows():
    commodity = str(row["Commodity"]).lower()
    zip_code = row["Zip Code"]
    utility = row["Name"]
    
    # Skip rows where zip code is missing (represented by "00000")
    if zip_code == "00000":
        print(f"Skipping row {index}: Missing Zip Code for {row['LDC']}")
        continue

    key = (zip_code, utility, commodity)
    if key not in ZipsDict:
        ZipsDict[key] = {
            "ZipCode": zip_code,
            "Utility": utility,
            "Commodity": commodity
        }
    
# Resulting dictionary
print("Unique Zipcodes:")
for key, entry in ZipsDict.items():
    print(entry)