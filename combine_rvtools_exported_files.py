import pandas as pd
import os

print("Current Working Directory:", os.getcwd())

input_dir = "./exports"

# Create the directory if it doesn't exist
if not os.path.exists(input_dir):
    os.makedirs(input_dir)
    print(f"Created folder: {input_dir}")

# List all .xlsx files in the current directory
files = [f for f in os.listdir('.') if f.endswith('.xlsx')]

all_data = []

# Define the columns we want to extract
wanted_columns = {
    "vInfo": [
        "vInfoVMName",     
        "vInfoPowerstate",           
        "vInfoGuestHostName",            
        "vInfoPrimaryIPAddress",
        "vInfoOSTools",
        "vInfoNetwork1",   
        "vInfoCPUs",
        "vInfoMemory",
        "vInfoTotalDiskCapacityMiB",
        "vInfo_tags_Team",
        "vInfo_tags_Department",
        "vInfo_tags_Prod-Dev-Test",
        "vInfo_tags_BackupFrequency",
        "vInfoHost",
        "vInfoDataCenter",             
        "vInfoCluster",
        "vInfoVISDKServer",        
        "vInfoVISDKServerType"                      
    ]
}

# Process each file
for file in files:
    print(f"Processing {file}")
    xl = pd.ExcelFile(file)
    
    for sheet, columns in wanted_columns.items():
        if sheet in xl.sheet_names:
            df = xl.parse(sheet)
            available_columns = [col for col in columns if col in df.columns]
            
            missing_columns = set(columns) - set(available_columns)
            if missing_columns:
                print(f"Alert: {file} is missing columns: {missing_columns}")
            
            if available_columns:
                df_filtered = df[available_columns].copy()
                df_filtered["source_file"] = file
                all_data.append(df_filtered)
            else:
                print(f"{file} has no matching columns in sheet '{sheet}'.")

# If data is available, create the output files
if all_data:
    final_df = pd.concat(all_data, ignore_index=True)
    
    # Save as Excel
    output_xlsx = os.path.join(input_dir, "rvtools_export.xlsx")
    final_df.to_excel(output_xlsx, index=False)

    # Save as CSV
    output_csv = os.path.join(input_dir, "rvtools_export.csv")
    final_df.to_csv(output_csv, index=False)

    print(f"Exported files:\n- {output_xlsx}\n- {output_csv}")
else:
    print("None of the files contained the desired columns.")
