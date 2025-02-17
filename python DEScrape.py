import pandas as pd

def get_unique_zipcodes(file_path: str):
    # Read the Excel file
    df = pd.read_excel("InputData.xlsx")
    
    # Ensure the necessary columns exist
    if "Commodity" not in df.columns or "Zip Code" not in df.columns:
        raise ValueError("Missing required columns in the Excel file.")
    
    # Filter for Gas and Electricity commodities and extract unique zip codes
    gas_zip_codes = list(df[df["Commodity"].str.lower() == "gas"]["Zip Code"].dropna().astype(str).unique())
    electric_zip_codes = list(df[df["Commodity"].str.lower() == "electricity"]["Zip Code"].dropna().astype(str).unique())
    
    return gas_zip_codes, electric_zip_codes

# Example usage
if __name__ == "__main__":
    file_path = "InputData.xlsx"  # Update this path to your actual file
    unique_gas_zip_codes, unique_electric_zip_codes = get_unique_zipcodes("InputData.xlsx")
    print("Unique Gas Zip Codes:", unique_gas_zip_codes)
    print("Unique Electric Zip Codes:", unique_electric_zip_codes)
