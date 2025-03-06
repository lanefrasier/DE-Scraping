import pandas as pd

# Load the Excel files
df_input = pd.read_excel("InputData_ZipCodes.xlsx")
df_output = pd.read_excel("OutputData.xlsx")

# Convert relevant columns to strings
df_input["Zip Code"] = df_input["Zip Code"].astype(str)
df_output["ZipCode"] = df_output["ZipCode"].astype(str)

df_input["Bundle Id"] = df_input["Bundle Id"].astype(str)
df_output["Bundle"] = df_output["Bundle"].astype(str)

# Sort both DataFrames 
df_input_sorted = df_input.sort_values(by=["Zip Code", "Bundle Id"])
df_output_sorted = df_output.sort_values(by=["ZipCode", "Bundle"])

# Merge on Zip Code and Bundle Id.
df_merged = df_input_sorted.merge(
    df_output_sorted, 
    left_on=["Zip Code", "Bundle Id"], 
    right_on=["ZipCode", "Bundle"], 
    how="outer", 
    suffixes=("_input", "_output"),
    indicator=True
)

# Convert numeric fields to float for proper comparison
df_merged["rateamt"] = pd.to_numeric(df_merged["rateamt"], errors="coerce")
df_merged["Price"] = pd.to_numeric(df_merged["Price"], errors="coerce")

# Initialize Mismatch column with empty values
df_merged["Mismatch"] = ""

# Identify mismatches and append the reason
if "Term/End Date" in df_merged.columns and "Plan Term" in df_merged.columns:
    df_merged.loc[df_merged["Term/End Date"] != df_merged["Plan Term"], "Mismatch"] += "Plan Term, "

df_merged.loc[df_merged["Price"] != df_merged["rateamt"], "Mismatch"] += "Price, "
df_merged.loc[df_merged["UOM_input"] != df_merged["UOM_output"], "Mismatch"] += "UOM, "
df_merged.loc[df_merged["ZipCode"] == "00000", "Mismatch"] += "Zipcode, "

# Identify missing rows
df_merged.loc[df_merged["_merge"] == "left_only", "Mismatch"] += "Missing in output, "
df_merged.loc[df_merged["_merge"] == "right_only", "Mismatch"] += "Missing in input, "

# Remove trailing commas and spaces
df_merged["Mismatch"] = df_merged["Mismatch"].str.rstrip(", ")

# Filter only rows with mismatches
df_qc = df_merged[df_merged["Mismatch"] != ""]

df_qc = df_qc.drop(columns=["_merge"], errors="ignore")
columns_order = ["Mismatch"] + [col for col in df_qc.columns if col != "Mismatch"]
df_qc = df_qc[columns_order]

# Save the mismatched rows to a new Excel file "QC.xlsx"
df_qc.to_excel("QC.xlsx", index=False)
print("QC file generated successfully!")
