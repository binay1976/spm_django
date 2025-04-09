import os
import pandas as pd
import sys
import codecs
import matplotlib.pyplot as plt
from fpdf import FPDF
import traceback
from datetime import datetime
import plotly.express as px
import plotly.io as pio

# Ensure UTF-8 encoding (Fix for Windows terminal output)
sys.stdout = codecs.getwriter("utf-8")(sys.stdout.buffer, errors="ignore")

# Define Paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MEDIA_FOLDER = os.path.join(BASE_DIR, "media")
os.makedirs(MEDIA_FOLDER, exist_ok=True)
PROCESSED_PDF_PATH = os.path.join(MEDIA_FOLDER, "quick_report.pdf")

# Define the path to the static folder
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # Move up one level to `spm_live`
STATIC_DIR = os.path.join(BASE_DIR, "spmApp", "static")  # Adjust according to your Django static folder location
logo_path = os.path.join(STATIC_DIR, "Logo.png")  # Full path to Logo.png

# Add this function for logging
def log_error(error_msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] ERROR: {error_msg}")

def Quick_Report(file_path, selected_cms_id):
    """
    Processes the uploaded Excel file and filters data based on selected CMS_ID.
    """
    try:
        print(f"Starting Quick_Report processing...")
        print(f"Input file path: {file_path}")
        print(f"Selected CMS_ID: {selected_cms_id}")

        # Check file path
        if not os.path.exists(file_path):
            log_error(f"Input file not found at: {file_path}")
            return

        # Check if file is readable
        try:
            with open(file_path, 'rb') as f:
                f.read(1)
            print("✅ Input file is readable")
        except Exception as e:
            log_error(f"Cannot read input file: {str(e)}")
            return

        # Read Excel file
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            print(f"✅ Successfully read Excel file with {len(df)} rows")
        except Exception as e:
            log_error(f"Error reading Excel file: {str(e)}")
            return

        # Check columns
        print(f"Available columns: {df.columns.tolist()}")
        required_columns = [
            "CMS_ID",
            "Train_No",
            "Loco_No",
            "Desig",
            "Crew_Name",
            "Nom_CLI",
            "Distance",
            "Speed",
            "Run_No",
            "Time",
            "Date"
        ]
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            log_error(f"Missing required columns: {missing_columns}")
            return

        # Process CMS_ID
        if not selected_cms_id:
            log_error("No CMS_ID provided")
            return

        # Filter data
        df = df[df["CMS_ID"].astype(str) == selected_cms_id.strip()]
        if df.empty:
            log_error(f"No data found for CMS_ID: {selected_cms_id}")
            return

        print(f"✅ Found {len(df)} records for CMS_ID: {selected_cms_id}")
        
        # Continue with processing
        process_and_save(selected_cms_id, df)

    except Exception as e:
        log_error(f"Unexpected error: {str(e)}")
        traceback.print_exc()
        raise

def process_and_save(cms_id, df):
    """Processes filtered data and generates a PDF report."""
    try:
        print("\n=== Starting process_and_save ===")
        print(f"Processing data for CMS_ID: {cms_id}")
        print(f"DataFrame shape: {df.shape}")
        
        # Debug: Print sample of date and time columns
        print("\nSample of input data:")
        print(df[['Date', 'Time']].head())
        print("\nColumn dtypes:")
        print(df.dtypes)

        # Extract basic details
        try:
            train_no = df["Train_No"].iloc[0] if not df.empty else "N/A"
            loco_no = df["Loco_No"].iloc[0] if not df.empty else "N/A"
            designation = df["Desig"].iloc[0] if not df.empty else "N/A"
            pilot_name = df["Crew_Name"].iloc[0] if not df.empty else "N/A"
            nominated_cli = df["Nom_CLI"].iloc[0] if not df.empty else "N/A"
            total_km = round(df["Distance"].sum()/1000, 3) if not df.empty else 0
            top_speed = df["Speed"].max() if not df.empty else 0
            total_halt = df["Run_No"].max() if not df.empty else 0
        except Exception as e:
            print(f"Error extracting basic details: {str(e)}")
            raise

        # Date and Time Processing
        try:
            # Convert Date strings to datetime objects
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            print("\nDate conversion successful")
            print(f"Sample converted dates: {df['Date'].head()}")

            # Convert Time strings to time objects
            if isinstance(df['Time'].iloc[0], str):
                df['Time'] = pd.to_datetime(df['Time'], format='%H:%M:%S', errors='coerce').dt.time
            print("\nTime conversion successful")
            print(f"Sample times: {df['Time'].head()}")

            # Combine Date and Time
            df['DateTime'] = df.apply(
                lambda row: pd.Timestamp.combine(row['Date'].date(), row['Time']) 
                if pd.notnull(row['Date']) and row['Time'] is not None 
                else pd.NaT, 
                axis=1
            )

            # Format datetime range
            if not df['DateTime'].isna().all():
                min_datetime = df['DateTime'].min().strftime('%d-%m-%Y %H:%M:%S')
                max_datetime = df['DateTime'].max().strftime('%d-%m-%Y %H:%M:%S')
                date_range = f"{min_datetime} to {max_datetime}"
            else:
                min_datetime = "No valid date"
                max_datetime = "No valid date"
                date_range = "No valid date range available"

            print(f"\nProcessed date range: {date_range}")

        except Exception as e:
            print(f"\n❌ Error in date processing: {str(e)}")
            print("Detailed error information:")
            import traceback
            traceback.print_exc()
            min_datetime = "Date processing error"
            max_datetime = "Date processing error"
            date_range = "Date processing error"

        # Time calculations
        try:
            df["TimeDiff"] = df["DateTime"].diff().dt.total_seconds().fillna(0)
            running_time_seconds = df.loc[df["Speed"] > 0, "TimeDiff"].sum()
            halt_time_seconds = df.loc[df["Speed"] == 0, "TimeDiff"].sum()
            
            running_time_str = str(pd.Timedelta(seconds=running_time_seconds)).split()[-1]
            halt_time_str = str(pd.Timedelta(seconds=halt_time_seconds)).split()[-1]
            
            avg_speed = round((total_km / (running_time_seconds + halt_time_seconds)) * 3600, 2) if running_time_seconds > 0 else 0
        except Exception as e:
            print(f"Error in time calculations: {str(e)}")
            running_time_str = "N/A"
            halt_time_str = "N/A"
            avg_speed = 0

        # Calculate WS to WS Duration
        try:
            if not df['DateTime'].isna().all():
                start_time = df['DateTime'].min()
                end_time = df['DateTime'].max()
                duration = end_time - start_time
                
                # Convert duration to hours and minutes
                total_hours = duration.total_seconds() / 3600
                hours = int(total_hours)
                minutes = int((total_hours - hours) * 60)
                
                ws_duration = f"{hours:02d}:{minutes:02d} Hrs"
                print(f"WS to WS Duration: {ws_duration}")
            else:
                ws_duration = "Duration not available"
        except Exception as e:
            print(f"Error calculating WS duration: {str(e)}")
            ws_duration = "Error calculating duration"

        # Generate PDF with all analyses
        save_to_pdf(
            cms_id=cms_id,
            train_no=train_no,
            loco_no=loco_no,
            total_km=total_km,
            top_speed=top_speed,
            total_duration=date_range,
            ws_duration=ws_duration,
            designation=designation,
            pilot_name=pilot_name,
            total_halt=total_halt,
            nominated_cli=nominated_cli,
            min_datetime=min_datetime,
            max_datetime=max_datetime,
            running_time_str=running_time_str,
            halt_time_str=halt_time_str,
            avg_speed=avg_speed,
            df=df  # Pass the DataFrame
        )

    except Exception as e:
        print(f"❌ Error in process_and_save: {str(e)}")
        raise

def save_to_pdf(cms_id, train_no, loco_no, total_km, top_speed, total_duration,
                ws_duration, designation, pilot_name, total_halt, nominated_cli,
                min_datetime, max_datetime, running_time_str, halt_time_str, avg_speed, df):
    """Generates and saves the PDF report."""
    try:
        print("\n=== Starting save_to_pdf ===")
        
        # Check if logo exists
        if not os.path.exists(logo_path):
            print(f"❌ Warning: Logo file not found at {logo_path}")
        
        # Use consistent file name
        pdf_file_path = os.path.join(MEDIA_FOLDER, "processed_quick_report.pdf")
        print(f"PDF will be saved to: {pdf_file_path}")

        class CustomPDF(FPDF):
            def footer(self):
                self.set_y(-15)
                self.set_font("Arial", size=6)
                self.cell(0, 10, "Western Railway, Mumbai Division @BDTS1022", align='R')

        # Initialize PDF
        pdf = CustomPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        # Add content with error handling
        try:
            # Add logo if it exists
            if os.path.exists(logo_path):
                pdf.image(logo_path, x=10, y=10, w=25, h=30)
            
            # Add header
            pdf.set_font("Arial", "B", 16)
            pdf.set_text_color(176, 224, 230)
            pdf.set_xy(40, 15)
            pdf.cell(0, 10, "Loco Pilot Driving Technique Analysis", ln=True, align="C")
            pdf.set_font("Arial", 13)
            pdf.set_text_color(0,0,0)
            pdf.set_xy(40, 25)
            pdf.cell(0, 10, "Western Railway", ln=True, align="C")
            pdf.line(60, 35, 200, 35)
            
            # Format the date range string
            if min_datetime == max_datetime == "Data not available":
                date_range = "Data not available"
            elif min_datetime == max_datetime == "Error processing date":
                date_range = "Error processing date information"
            else:
                date_range = f"{min_datetime} to {max_datetime}"
            
            # Add content
            pdf.set_font("Arial", size=10)
            pdf.cell(200, 10, f"Report For Crew CMS ID: {cms_id}", ln=True, align="C")
            pdf.ln(10)
            
            # Add all the details with better error handling
            details = [
                ("Prepared By", "............................................"),
                ("Record Period", total_duration if total_duration != "Date processing error" 
                 else "Date information not available"),
                ("WS to WS Duration", ws_duration),
                ("Loco Pilot Name", pilot_name),
                ("Designation", designation),
                ("CMS_ID", cms_id),
                ("Nominated CLI", nominated_cli),
                ("Loco Number", loco_no),
                ("Train Number", train_no),
                ("Start Date & Time", min_datetime),
                ("Finished Date & Time", max_datetime),
                ("Total Distance", f"{total_km} KM"),
                ("Total Running Time", f"{running_time_str} Hrs."),
                ("Total Halt Time", f"{halt_time_str} Hrs."),
                ("Top Speed", f"{top_speed} Kmph"),
                ("Average Speed", f"{avg_speed} Kmph"),
                ("Total Halt", f"{total_halt} Times")
            ]
            
            for label, value in details:
                pdf.cell(200, 10, f"{label}: {value}", ln=True, align="L")
            
            print("✅ Successfully added all content to PDF")
            
        except Exception as e:
            print(f"❌ Error adding content to PDF: {str(e)}")
            raise

        # ====== BPT/BFT Table & Data ======
        # pdf.add_page()
        if "BPT" in df.columns:
            bpt_filtered = df[df["BPT"] == "BPT"]
        else:
            bpt_filtered = pd.DataFrame()

        pdf.set_font("Arial", style="B", size=14)
        pdf.cell(200, 10, f"Brake Feel & Brake Power Test Conducted", ln=True, align="L")
        
        # Data for First Row - BFT Done
        bft_filtered = df[df["BFT"] == "BFT"]
        bft_time = bft_filtered["Time"].iloc[0] if not bft_filtered.empty else "N/A"
        bft_dist = bft_filtered["Cum_Dist_LP"].iloc[0] if not bft_filtered.empty else "N/A"
        bft_speed = bft_filtered["Speed"].iloc[0] if not bft_filtered.empty else "N/A"
        
        # Data for BPT
        bpt_time = bpt_filtered["Time"].iloc[0] if not bpt_filtered.empty else "N/A"
        bpt_dist = bpt_filtered["Cum_Dist_LP"].iloc[0] if not bpt_filtered.empty else "N/A"
        bpt_speed = bpt_filtered["Speed"].iloc[0] if not bpt_filtered.empty else "N/A"

        # Create Table Header
        pdf.set_font("Arial", style="B", size=12)
        pdf.cell(40, 10, "Test Done", border=1, align="C")
        pdf.cell(40, 10, "Time", border=1, align="C")
        pdf.cell(40, 10, "Distance (In Mtr)", border=1, align="C")
        pdf.cell(40, 10, "At Speed", border=1, align="C")
        pdf.ln()

        # Table content
        pdf.set_font("Arial", size=12)
        # First Row - BFT Done
        pdf.cell(40, 10, "BFT Done", border=1, align="C")
        pdf.cell(40, 10, str(bft_time), border=1, align="C")
        pdf.cell(40, 10, str(bft_dist), border=1, align="C")
        pdf.cell(40, 10, str(bft_speed), border=1, align="C")
        pdf.ln()
        # Second Row - BPT Done
        pdf.cell(40, 10, "BPT Done", border=1, align="C")
        pdf.cell(40, 10, str(bpt_time), border=1, align="C")
        pdf.cell(40, 10, str(bpt_dist), border=1, align="C")
        pdf.cell(40, 10, str(bpt_speed), border=1, align="C")
        pdf.ln()

        # ===== Speed Slab Table & Data =====
        pdf.add_page()
        pdf.set_font("Arial", style="B", size=16)
        pdf.cell(200, 10, f"SPM Report For Loco_No {loco_no} & CMS_ID {cms_id}", ln=True, align="C")
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, f"Speed Slab Wise Distance and Time", ln=True, align="L")

        # Process datetime
        if not pd.api.types.is_datetime64_any_dtype(df["Time"]):
            df["Time"] = pd.to_datetime(df["Time"], format="%H:%M:%S", errors="coerce")
        df["DateTime"] = pd.to_datetime(df["Date"].astype(str) + " " + df["Time"].astype(str), errors="coerce")
        df = df.dropna(subset=["DateTime"])
        df = df.sort_values(by="DateTime")
        df["TimeDiff"] = df["DateTime"].diff().fillna(pd.Timedelta(0))

        # Define speed slabs
        speed_slabs = [
            (1, 10), (10, 20), (20, 30), (30, 40), (40, 50),
            (50, 60), (60, 70), (70, 80), (80, 90), (90, 100),
            (100, 110), (110, 120), (120, 130), (130, 150)
        ]

        # Table Headers
        pdf.set_font("Arial", style="B", size=12)
        pdf.cell(60, 10, "Speed Slab (km/h)", border=1, align="C")
        pdf.cell(60, 10, "Total Distance (Meter)", border=1, align="C")
        pdf.cell(60, 10, "Total Time (HH:MM:SS)", border=1, align="C")
        pdf.ln()

        # Populate Table Data
        pdf.set_font("Arial", size=12)
        for start, end in speed_slabs:
            subset_df = df[(df['Speed'] >= start) & (df['Speed'] < end)]
            total_distance = subset_df["Distance"].sum() / 1000 if not subset_df.empty else 0
            total_time = subset_df["TimeDiff"].sum() if not subset_df.empty else pd.Timedelta(0)
            pdf.cell(60, 10, f"{start}-{end}", border=1, align="C")
            pdf.cell(60, 10, f"{total_distance:.3f} km", border=1, align="C")
            pdf.cell(60, 10, strfdelta(total_time), border=1, align="C")
            pdf.ln()

        # ===== Bar Graph For Run Numbers =====
        pdf.add_page(orientation='L')
        max_speed = df.groupby(['Run_No', 'CMS_ID'])['Speed'].max().reset_index()
        max_speed = max_speed[max_speed['CMS_ID'] == cms_id]
        sum_max_speed = max_speed.groupby('Run_No')['Speed'].sum().reset_index()
        
        # Create bar chart using plotly
        fig = px.bar(sum_max_speed, x="Run_No", y="Speed", text="Speed")
        
        # Set colors
        colors = ['#05B7B7'] * len(sum_max_speed)
        max_sum_max_speed_index = sum_max_speed['Speed'].idxmax()
        if max_sum_max_speed_index >= 0:
            colors[max_sum_max_speed_index] = '#854c03'
        
        fig.update_traces(marker_color=colors)
        fig.update_layout(
            title=f"Max. Speed of Each Halt (CMS_ID: {cms_id})",
            xaxis_title="Halt Count",
            yaxis_title="Speed"
        )
        fig.update_xaxes(type='category')
        
        # Save chart
        chart_path = os.path.join(MEDIA_FOLDER, "max_speed_chart.png")
        pio.write_image(fig, chart_path, format="png", width=1200, height=600, scale=2)
        
        # Add to PDF
        pdf.set_font("Arial", style="B", size=14)
        pdf.cell(280, 10, f"Top Speed Between Each Halt (CMS_ID: {cms_id} & Loco_No {loco_no})", ln=True, align='C')
        pdf.image(chart_path, x=10, y=40, w=285)

        # Save the complete PDF
        pdf.output(pdf_file_path)
        print("✅ PDF generated successfully with all pages")

        # Clean up temporary files
        if os.path.exists(chart_path):
            os.remove(chart_path)

    except Exception as e:
        print(f"❌ Error in save_to_pdf: {str(e)}")
        raise

def strfdelta(timedelta):
    """Helper function to format timedelta"""
    if pd.isna(timedelta):
        return " "
    total_seconds = timedelta.total_seconds()
    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    seconds = int(total_seconds % 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"

if __name__ == "__main__":
    try:
        if len(sys.argv) != 3:
            print(f"Error: Expected 2 arguments, got {len(sys.argv) - 1}")
            print("Usage: python quick_report.py <file_path> <cms_id>")
            sys.exit(1)
            
        file_path = sys.argv[1]
        selected_cms_id = sys.argv[2]
        print(f"Starting script with file_path: {file_path}, cms_id: {selected_cms_id}")
        Quick_Report(file_path, selected_cms_id)
        print("Script completed successfully")
    except Exception as e:
        print(f"Fatal error: {str(e)}")
        sys.exit(1)

