import pandas as pd
import glob
import os
from datetime import datetime

def run_qc():
    # Load the Excel files
    df_input = pd.read_excel("InputData_ZipCodes.xlsx")

    # Find all files matching the pattern "Output_*.xlsx"
    files = glob.glob("Output_*.xlsx")

    # If there are any matching files, find the most recent one
    if files:
        latest_file = max(files, key=os.path.getctime)  # Get the most recently created file
        print(f"Reading the latest file: {latest_file}")

        # Load the most recent file into a DataFrame
        df_output = pd.read_excel(latest_file)
    else:
        print("No matching files found.")
        df_output = pd.DataFrame() 

    # Convert relevant columns to strings
    df_input["Zip Code"] = df_input["Zip Code"].astype(str)
    df_input["Bundle Id"] = df_input["Bundle Id"].astype(str)
    df_input["Name"] = df_input["Name"].astype(str)

    df_output["ZipCode"] = df_output["ZipCode"].astype(str)
    df_output["Bundle"] = df_output["Bundle"].astype(str)
    df_output["Utility"] = df_output["Utility"].astype(str)

    # Sort both DataFrames 
    df_input_sorted = df_input.sort_values(by=["Zip Code", "Bundle Id", "Name"])
    df_output_sorted = df_output.sort_values(by=["ZipCode", "Bundle", "Utility"])

    # Merge on Zip Code and Bundle Id.
    df_merged = df_input_sorted.merge(
        df_output_sorted, 
        left_on=["Zip Code", "Bundle Id", "Name"], 
        right_on=["ZipCode", "Bundle", "Utility"], 
        how="outer", 
        suffixes=("_input", "_output"),
        indicator=True
    )

    # Convert numeric fields to float for proper comparison
    df_merged["rateamt"] = pd.to_numeric(df_merged["rateamt"], errors="coerce")
    df_merged["Price"] = pd.to_numeric(df_merged["Price"], errors="coerce")

    # Initialize Match column with empty values
    df_merged["Match"] = ""

    # Identify mismatches and append the reason
    if "Term/End Date" in df_merged.columns and "Plan Term" in df_merged.columns:
        df_merged.loc[df_merged["Term/End Date"] != df_merged["Plan Term"], "Match"] += "Plan Term, "
    df_merged.loc[df_merged["Price"] != df_merged["rateamt"], "Match"] += "Price, "
    df_merged.loc[df_merged["UOM_input"] != df_merged["UOM_output"], "Match"] += "UOM, "
    df_merged.loc[df_merged["ZipCode"] == "00000", "Match"] += "Zipcode, "
    df_merged.loc[df_merged["_merge"] == "left_only", "Match"] += "Missing in output, "
    df_merged.loc[df_merged["_merge"] == "right_only", "Match"] += "Missing in input, "

    # Remove trailing commas and spaces
    df_merged["Match"] = df_merged["Match"].str.rstrip(", ")

    # Mark successful matches
    df_merged.loc[df_merged["Match"] == "", "Match"] = "Successful"

    # Remove _merge column
    df_merged = df_merged.drop(columns=["_merge"], errors="ignore")

    # Move Match column to the first position
    columns_order = ["Match"] + [col for col in df_merged.columns if col != "Match"]
    df_merged = df_merged[columns_order]

    # Generate filename with today's date
    current_date = datetime.now().strftime("%m%d%Y")  # Format: MMDDYYYY
    qc_filename = f"QC_{current_date}.xlsx"

    # Save the entire dataset to the new file
    df_merged.to_excel(qc_filename, index=False)
    print(f"QC file generated successfully: {qc_filename}")
