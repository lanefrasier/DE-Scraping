import pandas as pd

# Define file path and sheet name (adjust as needed)
input_file = "InputData.xlsx"
sheet_name = "COM_QC"

# Read the Excel file into a DataFrame
dt_Input = pd.read_excel(input_file, sheet_name=sheet_name, dtype={"Zip Code": str})

# Initialize empty lists for Gas and Electricity zip codes
list_ZipGas = []
list_ZipElectricity = []

# Iterate through each row in the DataFrame
for index, row in dt_Input.iterrows():
    # Convert Commodity to string and ensure Zip Code is a 5-character string
    commodity = str(row["Commodity"])
    zip_code = str(row["Zip Code"]).zfill(5)  # Ensures leading zeros are preserved
    
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
    
    # If commodity is neither Electric nor Gas, raise an exception
    else:
        raise Exception(f"New commodity detected: {commodity}")

# Optional: print the resulting lists
print("Electricity Zip Codes:", list_ZipElectricity)
print("Gas Zip Codes:", list_ZipGas)
