import os
import pandas as pd
import numpy as np
import sys
import re

# Get the base directory of the Django project
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Define the media folder path
MEDIA_FOLDER = os.path.join(BASE_DIR, "media")

# Ensure media folder exists
if not os.path.exists(MEDIA_FOLDER):
    os.makedirs(MEDIA_FOLDER)
    print(f"✅ Created media folder at: {MEDIA_FOLDER}")

# Processed file path
PROCESSED_FILE_PATH = os.path.join(MEDIA_FOLDER, "processed_medha.xlsx")

def process_medha(file_path):
    try:
        print(f"✅ Processing file: {file_path}")

        # # Read input file
        # df = pd.read_csv(file_path, delimiter="\t", encoding="utf-8", on_bad_lines="skip")
        encodings_to_try = ["utf-8", "latin-1", "utf-16", "windows-1252"]
        for encoding in encodings_to_try:
            try:
                df = pd.read_csv(file_path, delimiter="\t", encoding=encoding, on_bad_lines="skip")
                break  # Stop trying once successful
            except UnicodeDecodeError:
                continue  # Try next encoding if Unicode error occurs

        # ✅ Data Processing
        df.columns = ["New Header 1"]
        split_values = df["New Header 1"].str.split("|", expand=True)
        df["Date"] = split_values[0]
        df["Time"] = split_values[1]
        df["Speed"] = split_values[2]
        df["Distance"] = split_values[3]
        df['Column 5'] = ''
        df['Column 6'] = ''
        df['Column 7'] = ''
        df['Column 8'] = ''
        df['Column 9'] = ''
        df['Column 10'] = ''
        df['Column 11'] = ''

        # Delete unnecessary columns
        columns_to_delete = ['Column 5', 'Column 6', 'Column 7', 'Column 8', 'Column 9', 'Column 10', 'Column 11']
        df.drop(columns=columns_to_delete, inplace=True)
        # df = df.drop(columns=["New Header 1"])

         # Cut and paste Column 1 cell value to Column 5 if it contains "Driver"
        df.loc[df['Date'].str.contains('Driver', na=False), 'Column 5'] = df['Date']

        # Copy down cell value of Column 5 if Column 4 data is not null
        df['Column 5'] = df['Column 5'].ffill()

        # Delete the Entire row with garbage value (Condition applied "Where Col 3 is null")
        df = df.dropna(subset=['Speed'])
        df.reset_index(drop=True, inplace=True)

        # Delete Column 1 which is useless now
        df = df.drop("New Header 1", axis=1)
        # Split values in Column 5 using ":"
        split_values = df['Column 5'].str.split(":", expand=True)
        df['Column 6'] = split_values[0]
        df['Column 7'] = split_values[1]
        df['Column 8'] = split_values[2]
        df['Column 9'] = split_values[3]

        # Delete extra text from each row in Column 7, 8, 9 and Remove White Space Using ".str.strip()"
        df['CMS_ID'] = df['Column 7'].str.replace('Train No', '').str.strip()
        df['Train_No'] = df['Column 8'].str.replace('Locono', '').str.strip()
        df['Loco_No'] = df['Column 9'].str.replace('Spd Limit', '').str.strip()

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

        # Apply the function to the 'Time' column and create a new column 'Formatted_Time' ..................................................................
        df['Time'] = df['Time'].apply(clean_and_convert_time)
        
        # Delete unnecessary columns
        columns_to_delete = ['Column 5', 'Column 6', 'Column 7', 'Column 8', 'Column 9']
        df.drop(columns=columns_to_delete, inplace=True) 

        # Convert certain columns to numeric format .................................................................................................
        numeric_columns = ['Speed', 'Distance','Loco_No']
        df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')
        df = df.dropna(subset=["Speed", "Distance"]) 

            # Remove Gaps in CMS ID
        df['CMS_ID'] = df['CMS_ID'].str.replace(' ', '', regex=False)  

        # Cumulate the value of 'Distance' based on the Start & Stop .................................................................................
        # Convert 'Distance' column to numeric, coercing errors to NaN
        df['Distance'] = pd.to_numeric(df['Distance'], errors='coerce')

        # Drop rows where 'Distance' is NaN
        df = df.dropna(subset=['Distance'])
        df['Cum_Dist_Run'] = df.groupby(df['Speed'].eq(0).cumsum())['Distance'].cumsum()

        # Create a new column 'Run_No' based on the conditions ...................................................................................
        # Initialize 'Run_No' as 1
        df['Run_No'] = 1
        # Identify when 'CMS_ID' changes
        cms_change = df['CMS_ID'] != df['CMS_ID'].shift()
        # Identify when 'Cum_Dist_Run' decreases
        distance_reset = df['Cum_Dist_Run'] < df['Cum_Dist_Run'].shift()
        # Increment 'Run_No' where 'Cum_Dist_Run' resets but within the same CMS_ID
        df['Run_No'] = cms_change.cumsum() + distance_reset.groupby(cms_change.cumsum()).cumsum()
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
        df = df[df['Run_No'] >= 10].reset_index(drop=True)

    # Create a new column 'Pin_Point' with value "10 Meters" for the rows closest to 10 in each 'Run_No' group .........................................................
        df['Pin_Point'] = np.where(df.groupby(['Run_No','CMS_ID'])['Rev_Dist'].transform(lambda x: abs(x - 10).idxmin()) == df.index, '10 Meters', '')
        # Update 'Pin_Point' column with other values
        df['Pin_Point'] = np.where(df.groupby(['Run_No','CMS_ID'])['Rev_Dist'].transform(lambda x: abs(x - 250).idxmin()) == df.index, '250 Meters', df['Pin_Point'])
        df['Pin_Point'] = np.where(df.groupby(['Run_No','CMS_ID'])['Rev_Dist'].transform(lambda x: abs(x - 500).idxmin()) == df.index, '500 Meters', df['Pin_Point'])
        df['Pin_Point'] = np.where(df.groupby(['Run_No','CMS_ID'])['Rev_Dist'].transform(lambda x: abs(x - 1000).idxmin()) == df.index, '1000 Meters', df['Pin_Point'])

# Add BFT Column ......................................................................................................................
        df['Speed_shift'] = df['Speed'].shift(-1)
        unique_cms_ids = set()
        def add_bft(row):
            if row['Cum_Dist_LP'] < 10000 and 15 <= row['Speed'] <= 18 and row['Speed'] > row['Speed_shift']:
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
        
    # Rearrange the Columns
        df = df[["Date", "Time", "Speed", "Distance", "CMS_ID", "Train_No", "Loco_No", "Cum_Dist_Run","Cum_Dist_LP","Run_No","Run_Sum","Rev_Dist","Pin_Point","BFT","BPT","BFT_BPT"]]


    # Adding a new column 'Crew Name' , 'CLI Name', 'Desig'............................................................................................................................
        try:
            cms_file_path = 'CMS_Data.xlsx'
            cms_df = pd.read_excel(cms_file_path)

            # Your existing mapping logic using the CMS data
            df['Crew_Name'] = df['CMS_ID'].map(cms_df.set_index(cms_df.columns[0])[cms_df.columns[1]])
            df['Nom_CLI'] = df['CMS_ID'].map(cms_df.set_index(cms_df.columns[0])[cms_df.columns[4]])
            df['Desig'] = df['CMS_ID'].map(cms_df.set_index(cms_df.columns[0])[cms_df.columns[2]])

        except FileNotFoundError:
            # File not found handling: Create columns with null values
            df['Crew_Name'] = None
            df['Nom_CLI'] = None
            df['Desig'] = None


# ✅ Save processed file in media folder ==============================================================================
        df.to_excel(PROCESSED_FILE_PATH, index=False)

        print(f"✅ Processed file saved at: {PROCESSED_FILE_PATH}")

    except Exception as e:
        print(f"❌ Error processing file: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("❌ Error: Missing arguments! Usage: python medha.py <input_file>")
        sys.exit(1)

    file_path = sys.argv[1]

    if not os.path.exists(file_path):
        print(f"❌ Error: Input file '{file_path}' not found!")
        sys.exit(1)

    process_medha(file_path)

    # ✅ Confirm processed file exists
    if os.path.exists(PROCESSED_FILE_PATH):
        print(f"✅ Final confirmation: {PROCESSED_FILE_PATH} exists.")
    else:
        print("❌ ERROR: Processed file was not saved!")
