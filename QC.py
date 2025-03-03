import pandas as pd

# Load the Excel files
df_input = pd.read_excel("InputData.xlsx")
df_output = pd.read_excel("OutputData.xlsx")

# Convert relevant columns to strings
df_input["Zip Code"] = df_input["Zip Code"].astype(str)
df_output["ZipCode"] = df_output["ZipCode"].astype(str)

df_input["Bundle Id"] = df_input["Bundle Id"].astype(str)
df_output["Bundle"] = df_output["Bundle"].astype(str)

# Sort both DataFrames 
df_input_sorted = df_input.sort_values(by=["Zip Code", "Bundle Id"])
df_output_sorted = df_output.sort_values(by=["ZipCode", "Bundle"])

# Merge on Zip Code and Commodity.
df_merged = df_input_sorted.merge(
    df_output_sorted, 
    left_on=["Zip Code", "Bundle Id"], 
    right_on=["ZipCode", "Bundle"], 
    how="left", 
    suffixes=("_input", "_output")
)

# Convert numeric fields to float for proper comparison
df_merged["rateamt"] = pd.to_numeric(df_merged["rateamt"], errors="coerce")
df_merged["Price"] = pd.to_numeric(df_merged["Price"], errors="coerce")

# Create a column that checks if PlanName (from OutputData) is contained in Description (from InputData)
df_merged["PlanNameInDescription"] = df_merged.apply(
    lambda row: isinstance(row["PlanName"], str) and isinstance(row["Description"], str) 
                and (row["PlanName"].lower() in row["Description"].lower()),
    axis=1
)

# Define mismatch conditions:
# 1. PlanName not found in Description.
# 2. Price mismatch between rateamt (input) and Price (output).
# 3. UOM mismatch between UOM_input and UOM_output.
# 4. Terms and Conditions mismatch between TandC (input) and TermsAndConditions (output).
# 5. Missing output row (e.g. ZipCode is NaN).
mismatch_conditions = (
    (~df_merged["PlanNameInDescription"]) |
    (df_merged["Price"] != df_merged["rateamt"]) |
    (df_merged["UOM_input"] != df_merged["UOM_output"]) |
    (df_merged["TandC"] != df_merged["TermsAndConditions"]) |
    (df_merged["ZipCode"].isna())
)

df_qc = df_merged[mismatch_conditions]

# Save the mismatched rows to a new Excel file "QC.xlsx"
df_qc.to_excel("QC.xlsx", index=False)
print("QC file generated successfully!")
