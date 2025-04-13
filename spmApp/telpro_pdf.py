# telpro_pdf.py
import pdfplumber
import pandas as pd
import warnings
import sys
import re
import numpy as np
import os
import pandas as pd

def process_pdf(input_path, output_path):
    warnings.filterwarnings("ignore")
    all_data = []

    with pdfplumber.open(input_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                all_data.extend(table)
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    all_data.append(line.split())

    if not all_data:
        print("No data extracted from PDF.")
        return

    max_cols = max(len(row) for row in all_data)
    df = pd.DataFrame(all_data, columns=[f"Column_{i+1}" for i in range(max_cols)])

    for col in df.columns:
        if (df[col] == "RUN").sum() >= 10:
            df.rename(columns={col: "Status"}, inplace=True)

    if "Column_4" in df.columns:
        df.insert(df.columns.get_loc("Column_4") + 1, "data", "")

    if "Status" in df.columns:
        column_order = [
            "Column_1", "Column_2", "Column_3", "Column_4", "data",
            "Column_6", "Column_7", "Column_8", "Column_9",
            "Column_10", "Column_11", "Status"
        ]
        df = df[[col for col in column_order if col in df.columns]]
        print("Status & Data Column Created and Reordered")

    # -------------------------------------------------------------------------------------------------------
    # Identify columns to clear (all except the required ones)
    columns_to_clear = [col for col in df.columns if col not in ["Column_1", "Column_2", "Column_3", "Column_4", "Status"]]
    # Clear data from unwanted columns by setting values to empty strings or NaN
    df[columns_to_clear] = ""
# -------------------------------------------------------------------------------------------------------
    # Copy Column_1 data to Column_5 for those rows who has Concatenated Value with Driver ID:-
    if not df.empty and "Column_1" in df and "data" in df:  # Identify rows where Column_1 contains "Driver"
        driver_rows = df["Column_1"].astype(str).str.contains("Driver ID:-", na=False, regex=True)
        df.loc[driver_rows, "data"] = df.loc[driver_rows, "Column_1"]
    if "data" not in df.columns:  # Check if "Column_1" exists before filtering
        print("Error: data not found in extracted data.")
        return
# -------------------------------------------------------------------------------------------------------    
    def extract_values(row):
        match = re.search(
            r"Driver ID:-\s*([\S]+)\s*Train No:-\s*([\S ]+?)\s*Tr Load\(Ton\):-\s*([\S]+)\s*Loco No:-\s*([\S]+)\s*Wheel Dia\(mm\):-\s*([\S]+)\s*SpeedLimit\(kmph\):-\s*([\S]+)",
            str(row),
            re.DOTALL  # Enables multi-line matching
        )
        if match:
            return list(match.groups())  # Convert tuple to list
        return [None] * 6  # Ensure every row has exactly 6 values

    # Apply function to DataFrame
    df[["Column_6", "Column_7", "Column_8", "Column_9", "Column_10", "Column_11"]] = df["data"].apply(lambda x: pd.Series(extract_values(x)))
    print("Splittings Done")
# -------------------------------------------------------------------------------------------------------
    def extract_values(row):
        pattern = r"""
            Driver\s+ID:-\s*(.+?)\s+
            Train\s+No:-\s*(.+?)\s+
            Tr\s+Load\(Ton\):-\s*(.+?)\s+
            Loco\s+No:-\s*(.+?)\s+
            Wheel\s+Dia\(mm\):-\s*(.+?)\s+
            SpeedLimit\(kmph\):-\s*(.+?)
        """
        match = re.search(pattern, str(row), re.VERBOSE | re.DOTALL)
        if match:
            return list(match.groups())
        return [None] * 6
    df[["Column_6", "Column_7", "Column_8", "Column_9", "Column_10", "Column_11"]] = df["data"].apply(lambda x: pd.Series(extract_values(x)))
    print("Splittings Done")
# -------------------------------------------------------------------------------------------------------
     # Reaaarnge Columns
    desired_order = [
        "Column_1", "Column_2", "Column_3", "Column_4", "Status", "data", 
        "Column_6", "Column_7", "Column_8", "Column_9", "Column_10", "Column_11"
    ]
    df = df[desired_order]
# ----------------------------------------------------------------------------------------------
    # Columns to apply the fill operation
    columns_to_fill = ["Column_6", "Column_7", "Column_8", "Column_9", "Column_10", "Column_11"]
    changed_idx = df.index[df["Status"] == "CHANGED"].tolist()
    start = 0
    for i in range(len(changed_idx)):
        # Fill down from the last valid data point until "CHANGED"
        df.loc[start:changed_idx[i]-1, columns_to_fill] = df.loc[start:changed_idx[i]-1, columns_to_fill].ffill()
        next_valid = df[columns_to_fill].iloc[changed_idx[i]+1:].first_valid_index()
        if next_valid is not None:
            df.loc[changed_idx[i]+1:next_valid, columns_to_fill] = df.loc[changed_idx[i]+1:next_valid, columns_to_fill].bfill()
        start = changed_idx[i] + 1
    df.loc[start:, columns_to_fill] = df.loc[start:, columns_to_fill].ffill()
# -------------------------------------------------------------------------------------------------------
    # Filter: Keep rows where Column_1 contains "Driver ID:-" or "/2"
    df = df[df["Status"].astype(str).str.contains("RUN|STOP", na=False, regex=True)]
    desired_order = [
        "Column_1", "Column_2", "Column_3", "Column_4", 
        "Column_6", "Column_7", "Column_8", "Column_9", "Column_10", "Column_11", "Status", "data"
    ]
    df = df[desired_order]

    # Rename first four columns
    df.columns = ["Date", "Time", "Distance", "Speed", "CMS_ID", "Train_No","Column_8","Loco_No"] + list(df.columns[8:])
    # Rearrange the Columns
    df = df[["Date", "Time", "Speed", "Distance", "CMS_ID", "Train_No", "Loco_No", "Column_8","Status", "data"]]

    # Remove Gaps in CMS ID
    df['CMS_ID'] = df['CMS_ID'].str.replace(' ', '', regex=False)


    print("Analyzing PDF File, Please Wait.........")
    # message = f" Analysing PDF File, Please Wait........."
 # # Columns to be deleted
    columns_to_delete = ["Column_8","data", "Status"]
     # Ensure the columns exist before dropping
    df = df.drop(columns=[col for col in columns_to_delete if col in df.columns], errors="ignore")
   
    print("Basic Columns Done")
    # Trim white spaces from all string columns
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
     # Convert certain columns to numeric format .................................................................................................
    numeric_columns = ['Speed', 'Distance', 'Loco_No']
    df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')
    # Drop rows where "Speed" or "Distance" is NaN (non-numeric values)
    df = df.dropna(subset=["Speed", "Distance"])

# Function to clean and convert the text to time format ..........................................................................................
    def clean_and_convert_time(time_str):
        # Remove leading and trailing white spaces
        time_str = time_str.strip()
        # Use regular expression to find and extract time components
        match = re.match(r'^(\d+):(\d+):(\d+)$', time_str)
        if match:
            # If the regex pattern matches, extract components and format the time string
            hour, minute, second = match.groups()
            return f"{hour.zfill(2)}:{minute.zfill(2)}:{second.zfill(2)}"
        else:
            # If the regex pattern doesn't match, return None (or handle as desired)
            return None
    df['Time'] = df['Time'].apply(clean_and_convert_time)

# Cumulate the value of 'Distance' based on the Start & Stop .................................................................................
    df['Distance'] = pd.to_numeric(df['Distance'], errors='coerce')
    # Drop rows where 'Distance' is NaN
    df = df.dropna(subset=['Distance'])
    df['Cum_Dist_Run'] = df.groupby(df['Speed'].eq(0).cumsum())['Distance'].cumsum()

# Create a new column 'Run_No' based on the conditions ...................................................................................
    # Initialize 'Run_No' as 1
    df['Run_No'] = 1  
    cms_change = df['CMS_ID'] != df['CMS_ID'].shift()
    distance_reset = df['Cum_Dist_Run'] < df['Cum_Dist_Run'].shift()
    group = cms_change.cumsum()
    df['Run_No'] = distance_reset.groupby(group).cumsum() + 1

# Cumulate the value of 'Distance' based on the Start & Stop .................................................................................
    df['Cum_Dist_Run'] = df.groupby(df['Speed'].eq(0).cumsum())['Distance'].cumsum()

# Cumulate the value of 'Distance' based on the LP .................................................................................
    if {'Speed', 'CMS_ID', 'Distance'}.issubset(df.columns):
        df['Cum_Dist_LP'] = df.groupby('CMS_ID')['Distance'].cumsum()

# Insert a new column 'Run_Sum' with the sum of 'Distance' per 'Run_No' ....................................................................
    df['Run_Sum'] = df.groupby(['Run_No','CMS_ID'])['Distance'].transform('sum')

# Deduct values from 'Run_Sum' to 'Distance' column ........................................................................................
    df['Rev_Dist'] = df['Run_Sum'] - df['Cum_Dist_Run']


#  Delete Short Run ........................................................................................
#     df = df[df['Run_No'] >= 10].reset_index(drop=True)

# # Create a new column 'Pin_Point' with value "10 Meters" for the rows closest to 10 in each 'Run_No' group .........................................................
    df['Pin_Point'] = ''
    for dist, label in [(10, '10 Meters'), (250, '250 Meters'), (500, '500 Meters'), (1000, '1000 Meters')]:
        idxs = (
            df.groupby(['Run_No', 'CMS_ID'])['Rev_Dist']
            .apply(lambda x: (x - dist).abs().idxmin())
        )
        df.loc[idxs.values, 'Pin_Point'] = label


# Add BFT Column ......................................................................................................................
    df['Speed_shift'] = df['Speed'].shift(-1)
    unique_cms_ids = set()
    def add_bft(row):
        if row['Cum_Dist_LP'] < 10000 and 12 <= row['Speed'] <= 18 and row['Speed'] > row['Speed_shift']:
            if row['CMS_ID'] not in unique_cms_ids:
                unique_cms_ids.add(row['CMS_ID'])
                return 'BFT'
        return ''
    df['BFT'] = df.apply(add_bft, axis=1)

# Add BPT Column...............................................................................................................................
    unique_cms_ids = set()
    def add_bpt(row):
        if row['Cum_Dist_LP'] < 10000 and 40 <= row['Speed'] <= 90 and row['Speed'] > row['Speed_shift']:
            if row['CMS_ID'] not in unique_cms_ids:
                unique_cms_ids.add(row['CMS_ID'])
                return 'BPT'
        return ''
    df['BPT'] = df.apply(add_bpt, axis=1)

    # Reset shift column for comparison
    df['Speed_shift'] = df['Speed'].shift(1)

# BFT_END .............................................................................................................................
    def get_bft_end(df):
        df['BFT_END'] = ''
        for cms_id, group in df.groupby('CMS_ID'):
            group = group.reset_index()
            bft_idx = group[group['BFT'] == 'BFT'].index
            if not bft_idx.empty:
                bft_idx = bft_idx[0]
                for i in range(bft_idx + 1, len(group)):
                    if group.loc[i, 'Speed'] > group.loc[i, 'Speed_shift'] and 0 <= group.loc[i, 'Speed'] <= 9 and group.loc[i, 'Cum_Dist_LP'] < 10000:
                        end_idx = i - 1
                        if end_idx >= 0:
                            df.loc[group.loc[end_idx, 'index'], 'BFT_END'] = 'BFT_END'
                        break
        return df
# BPT_END ..............................................................................................................................
    def get_bpt_end(df):
        df['BPT_END'] = ''
        for cms_id, group in df.groupby('CMS_ID'):
            group = group.reset_index()
            bpt_idx = group[group['BPT'] == 'BPT'].index
            if not bpt_idx.empty:
                bpt_idx = bpt_idx[0]
                for i in range(bpt_idx + 1, len(group)):
                    if group.loc[i, 'Speed'] > group.loc[i, 'Speed_shift'] and 0 <= group.loc[i, 'Speed'] <= 45 and group.loc[i, 'Cum_Dist_LP'] < 10000:
                        end_idx = i - 1
                        if end_idx >= 0:
                            df.loc[group.loc[end_idx, 'index'], 'BPT_END'] = 'BPT_END'
                        break
        return df
    
    # Apply end marking functions
    df = get_bft_end(df)
    df = get_bpt_end(df)

    # Drop Speed_shift
    df.drop(columns=['Speed_shift'], inplace=True)

# Adding a new column 'BFT_BPT' ............................................................................................................................
    df['BFT_BPT'] = df.apply(lambda row: 
    (str(row['BFT']) if pd.notna(row['BFT']) and row['BFT'] != '' else '') +
    (' ' + str(row['BPT']) if pd.notna(row['BPT']) and row['BPT'] != '' else '') +
    (' ' + str(row['BFT_END']) if pd.notna(row['BFT_END']) and row['BFT_END'] != '' else '') +
    (' ' + str(row['BPT_END']) if pd.notna(row['BPT_END']) and row['BPT_END'] != '' else ''),
    axis=1
    ).str.strip()

# Adding a new column 'Crew Name' , 'CLI Name', 'Desig'............................................................................................................................
    try:
        # Construct path relative to this script
        base_dir = os.path.dirname(os.path.abspath(__file__))
        cms_file_path = os.path.join(base_dir, "CMS_Data.xlsx")
        
        cms_df = pd.read_excel(cms_file_path)

        # Normalize CMS_IDs to avoid mismatch due to leading/trailing spaces
        df['CMS_ID'] = df['CMS_ID'].astype(str).str.strip()
        cms_df[cms_df.columns[0]] = cms_df[cms_df.columns[0]].astype(str).str.strip()

        # Map values safely
        df['Crew_Name'] = df['CMS_ID'].map(cms_df.set_index(cms_df.columns[0])[cms_df.columns[1]])
        df['Desig'] = df['CMS_ID'].map(cms_df.set_index(cms_df.columns[0])[cms_df.columns[2]])
        df['Nom_CLI'] = df['CMS_ID'].map(cms_df.set_index(cms_df.columns[0])[cms_df.columns[4]])
         
        # Rearrange the Columns
        df = df[["Date", "Time", "Speed", "Distance", "CMS_ID", "Train_No", "Loco_No", "Crew_Name", "Desig","Nom_CLI","BFT_BPT","Cum_Dist_Run","Cum_Dist_LP","Run_No","Run_Sum","Rev_Dist","Pin_Point","BFT","BFT_END","BPT","BPT_END"]]

    except FileNotFoundError:
        print("CMS_Data.xlsx not found in script directory.")
        df['Crew_Name'] = None
        df['Nom_CLI'] = None
        df['Desig'] = None




    df.to_excel(output_path, index=False)
    print(f" File saved to {output_path}")
if __name__ == "__main__":
    import traceback

    if len(sys.argv) != 3:
        print("Usage: python telpro_pdf.py <input_pdf_path> <output_excel_path>")
    else:
        input_pdf = sys.argv[1]
        output_excel = sys.argv[2]
        try:
            process_pdf(input_pdf, output_excel)
        except Exception as e:
            print("Error while processing PDF:")
            traceback.print_exc()

