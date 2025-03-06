import pandas as pd

# Define file paths and sheet names
input_file = "InputData.xlsx"
zip_file = "USN SS Logins.xlsx"
sheet_input = "COM_QC"
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

# Save the merged DataFrame as a new file "InputData_ZipCodes.xlsx" with the same sheet name as the original
output_file = "InputData_ZipCodes.xlsx"
df_merged.to_excel(output_file, sheet_name=sheet_input, index=False)

# Initialize empty lists for Gas and Electricity zip codes
list_ZipGas = []
list_ZipElectricity = []

# Iterate through each row in the merged DataFrame
for index, row in df_merged.iterrows():
    commodity = str(row["Commodity"])
    zip_code = row["Zip Code"]
    
    # Skip rows where zip code is missing (represented by "00000")
    if zip_code == "00000":
        print(f"Skipping row {index}: Missing Zip Code")
        continue
    
    # Check if Commodity contains "Electric"
    if "Electric" in commodity:
        if zip_code in list_ZipElectricity:
            print(f"Zip Code {zip_code} is already in Electric List")
        else:
            list_ZipElectricity.append(zip_code)
    # Check if Commodity contains "Gas"
    elif "Gas" in commodity:
        if zip_code in list_ZipGas:
            print(f"Zip Code {zip_code} is already in Gas List")
        else:
            list_ZipGas.append(zip_code)
    else:
        raise Exception(f"New commodity detected: {commodity}")

# Print the resulting lists
print("Electricity Zip Codes:", list_ZipElectricity)
print("Gas Zip Codes:", list_ZipGas)