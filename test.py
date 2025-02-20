        """ try:
            df_existing = pd.read_excel("OutputData.xlsx", sheet_name="Data")
        except Exception:
            df_existing = pd.DataFrame()
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
 """
""" 
        with pd.ExcelWriter("OutputData.xlsx", mode="w", engine="openpyxl") as writer:
            df_combined.to_excel(writer, index=False, sheet_name="Data")
        print(f"Data successfully written for Zip Code: {in_ZipCode}")
    else: