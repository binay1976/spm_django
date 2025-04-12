import os
import sys
import pandas as pd
import re
import datetime

# Get the base directory of the Django project
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Define the media folder path
MEDIA_FOLDER = os.path.join(BASE_DIR, "media")

# Ensure media folder exists
if not os.path.exists(MEDIA_FOLDER):
    os.makedirs(MEDIA_FOLDER)
    print(f"✅ Created media folder at: {MEDIA_FOLDER}")

def process_laxvan(file_path, output_path, cms_id, train_no, loco_no):
    messages = []
    try:
        print(f"✅ Processing file: {file_path}")
        print(f"🧑 CMS_ID: {cms_id}, 🚂 Train No: {train_no}, 🔧 Loco No: {loco_no}")
        messages.append(f"✅ File processed for Loco No- {loco_no} & CMS ID {cms_id}")
                

        encodings_to_try = ["utf-8", "latin-1", "utf-16", "windows-1252"]
        content = None
        for encoding in encodings_to_try:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.readlines()
                break
            except UnicodeDecodeError:
                continue

        if content is None:
            raise ValueError("Could not decode file with available encodings.")

        # Parse each non-empty line into columns by whitespace or tab
        data = [line.strip().split() for line in content if line.strip()]

        # Make sure there are enough columns
        if not data or len(data) < 1 or len(data[0]) < 2:
            raise ValueError("Invalid or empty data.")

        # Optional: manually define column names if the file lacks headers
        num_cols = max(len(row) for row in data)
        col_names = [f"Col_{i+1}" for i in range(num_cols)]

        df = pd.DataFrame(data, columns=col_names)
        messages.append(f"✅ Please wait..... {len(df)} rows processing.")

        # Add extra columns
        df["CMS_ID"] = cms_id
        df["Train_No"] = train_no
        df["Loco_No"] = loco_no

        print(df.head())

        # Add computed Distance in meters
        df['Col_3'] = pd.to_numeric(df['Col_3'], errors='coerce')
        df['Distance'] = df['Col_3'].diff() * 1000
        df.loc[df.index[0], 'Distance'] = 0
        df['Distance'] = df['Distance'].round(2)

        # Step 3: Drop 'Km' and '---' columns if they exist
        columns_to_drop = ['Col_3', 'Col_5', 'Col_6', 'Col_7', 'Col_8', 'Col_9', 'Col_10', 'Col_11','Col_12', 'Col_13', 'Col_14', 'Col_15', 'Col_16']
        df.drop(columns=[col for col in columns_to_drop if col in df.columns], inplace=True)

        df.rename(columns={'Col_1': 'Date','Col_2': 'Time','Col_4': 'Speed'}, inplace=True)
        # Rearrange columns as "Date, Time, Speed, Distance, Train_No"............
        df = df[["Date", "Time", "Speed", "Distance", "CMS_ID", "Train_No", "Loco_No"]]

        # Delete rows where 'Date' column does not contain '/'
        if "Date" in df.columns:
            original_count = len(df)
            df = df[df["Date"].astype(str).str.contains("/", na=False)]
            removed_count = original_count - len(df)
            print(f"🧹 Removed {removed_count} rows without '/' in 'Date' column.")
        else:
            print("⚠️ 'Date' column not found, skipping filtering step.")

    # .....Basic Columns Done..............................................................................................................................
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

    # Convert certain columns to numeric format .................................................................................................
        numeric_columns = ['Speed', 'Distance', 'Loco_No']
        df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')
        # Drop rows where "Speed" or "Distance" is NaN (non-numeric values)
        df = df.dropna(subset=["Speed", "Distance"])

        try:
            # Convert to datetime and reformat to DD/MM/YYYY
            df['Date'] = pd.to_date(df['Date'], format='%d/%m/%y', errors='coerce')
            
        except Exception as e:
            print(f"⚠️ Failed to convert date format: {e}")

    # Function to clean and convert the text to time format ..........................................................................................
        def clean_and_convert_time(time_str):
            time_str = time_str.strip()
            match = re.match(r'^(\d+):(\d+):(\d+)$', time_str)
            if match:
                hour, minute, second = match.groups()
                return f"{hour.zfill(2)}:{minute.zfill(2)}:{second.zfill(2)}"
            else:
                return None
        df['Time'] = df['Time'].apply(clean_and_convert_time)

    # Cumulate the value of 'Distance' based on the Start & Stop .................................................................................
        # Convert 'Distance' column to numeric, coercing errors to NaN
        df['Distance'] = pd.to_numeric(df['Distance'], errors='coerce')
        df = df.dropna(subset=['Distance'])
        df['Cum_Dist_Run'] = df.groupby(df['Speed'].eq(0).cumsum())['Distance'].cumsum()
        messages.append(f"✅ Almost Done..... {len(df)} rows processing.")

    # Create a new column 'Run_No' based on the conditions ...................................................................................
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
#       df = df[df['Run_No'] >= 10].reset_index(drop=True)

# Create a new column 'Pin_Point' with value "10 Meters" for the rows closest to 10 in each 'Run_No' group .........................................................
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
        except FileNotFoundError:
            print("❌ CMS_Data.xlsx not found in script directory.")
            messages.append(f"✅ No Data found for CMS ID -{cms_id} Setting Blank name.")
            df['Crew_Name'] = None
            df['Nom_CLI'] = None
            df['Desig'] = None
            
            # Rearrange the Columns
            df = df[["Date", "Time", "Speed", "Distance", "CMS_ID", "Train_No", "Loco_No", "Crew_Name", "Desig","Nom_CLI","BPT_BFT","Cum_Dist_Run","Cum_Dist_LP","Run_No","Run_Sum","Rev_Dist","Pin_Point","BFT","BFT_END","BPT","BPT_END"]]








# Save to Excel
        df.to_excel(output_path, index=False)
        print(f"✅ Processed file saved at: {output_path}")

    except Exception as e:
        error_message = f"❌ Error during processing: {str(e)}"
        print(error_message)
        messages.append(error_message)

if __name__ == "__main__":
    if len(sys.argv) < 6:
        print("❌ Error: Missing arguments!\nUsage: python laxvan.py <input_file> <output_file> <cms_id> <train_no> <loco_no>")
        sys.exit(1)

    file_path = sys.argv[1]
    output_path = sys.argv[2]
    cms_id = sys.argv[3]
    train_no = sys.argv[4]
    loco_no = sys.argv[5]

    if not os.path.exists(file_path):
        print(f"❌ Error: Input file '{file_path}' not found!")
        sys.exit(1)

    process_laxvan(file_path, output_path, cms_id, train_no, loco_no)

    if os.path.exists(output_path):
        print(f"✅ Final confirmation: {output_path} exists.")
    else:
        print("❌ ERROR: Processed file was not saved!")
