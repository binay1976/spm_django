from django.http import HttpResponse, FileResponse
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
from django.conf import settings
import pandas as pd
import os
import time
import subprocess
from django.shortcuts import render, redirect
from django.contrib import messages  # For showing success messages
from django.http import JsonResponse
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
from django.conf import settings
from django.views.decorators.csrf import csrf_exempt
import sys


def index(request):
    return render(request, 'home/index.html')
def home(request):
    return render(request, "home/index.html")

# ====Delete Temp files============================================================================================
def clean_temp_files(request):
    """Deletes all temporary files in the MEDIA folder that start with 'temp'."""
    media_path = settings.MEDIA_ROOT
    deleted_files = []

    if os.path.exists(media_path):
        for file_name in os.listdir(media_path):
            if file_name.startswith("temp"):  # Check if file starts with "temp"
                file_path = os.path.join(media_path, file_name)
                try:
                    os.remove(file_path)  # Delete the file
                    deleted_files.append(file_name)
                except Exception as e:
                    messages.error(request, f"Error deleting {file_name}: {e}")

    if deleted_files:
        messages.success(request, f"Deleted {len(deleted_files)} temp files successfully! ✅")
    else:
        messages.info(request, "No temp files found. 😊")

    return redirect("/")  # Redirect back to the homepage


# =====Medha Upload ============================================================================================
def upload_medha(request):
    print("🚀 Views.Upload_Medha running")
    message = None
    processed_file_name = "processed_medha.xlsx"
    processed_file_path = os.path.join(settings.MEDIA_ROOT, processed_file_name)

    if request.method == "POST" and request.FILES.get("file"):
        uploaded_file = request.FILES["file"]

        # Allow only .txt files
        if not uploaded_file.name.endswith(".txt"):
            message = "Error: Only .txt files are allowed!"
        else:
            # Save uploaded file temporarily
            temp_file_path = default_storage.save("temp_medha.txt", ContentFile(uploaded_file.read()))
            full_temp_path = os.path.join(settings.MEDIA_ROOT, temp_file_path)

            print(f"✅ Uploaded file saved at: {full_temp_path}")
            print(f"✅ Processed file should be saved at: {processed_file_path}")

            # Run `telpro_pdf.py` script
            try:
                script_path = os.path.join(settings.BASE_DIR, "spmApp", "medha.py")
                subprocess.run([sys.executable, script_path, full_temp_path, processed_file_path], check=True)

                # Check if file exists after processing
                if os.path.exists(processed_file_path):
                    print(f"✅ Processed file FOUND at: {processed_file_path}")

                    # Serve file as an HTTP response for direct download
                    response = FileResponse(open(processed_file_path, "rb"), as_attachment=True)
                    response["Content-Disposition"] = f'attachment; filename="{processed_file_name}"'
                    message = "🚀 Great: File Processed Successfully!"
                    return response
                else:
                    print("❌ ERROR: Processed file is missing after execution!")
                    message = "Error: Processed file not found!"

            except subprocess.CalledProcessError as e:
                message = f"Error processing file: {e}"

    return render(request, "home/upload_medha.html", {"message": message})


# Telpro_PDf ============================================================================================
import os
import subprocess
from django.conf import settings
from django.shortcuts import render
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
from django.http import FileResponse
import sys

def upload_telpro_pdf(request):
    print("🚀 Views.upload_telpro running")
    message = None
    processed_file_name = "processed_telpro.xlsx"
    processed_file_path = os.path.join(settings.MEDIA_ROOT, processed_file_name)

    if request.method == "POST" and request.FILES.get("file"):
        uploaded_file = request.FILES["file"]

        # Allow only PDF files
        if not uploaded_file.name.endswith(".pdf"):
            message = "❌ Error: Only .pdf files are allowed!"
        else:
            # Save uploaded PDF temporarily
            temp_file_path = default_storage.save("temp_telpro.pdf", ContentFile(uploaded_file.read()))
            full_temp_path = os.path.join(settings.MEDIA_ROOT, temp_file_path)

            print(f"✅ Uploaded PDF saved at: {full_temp_path}")
            print(f"✅ Processed Excel will be saved at: {processed_file_path}")

            try:
                # Call telpro_pdf.py using subprocess and pass input/output paths
                script_path = os.path.join(settings.BASE_DIR, "spmApp", "telpro_pdf.py")
                subprocess.run([sys.executable, script_path, full_temp_path, processed_file_path], check=True)


                # Serve the Excel file if created
                if os.path.exists(processed_file_path):
                    print("✅ Processed Excel FOUND!")
                    response = FileResponse(open(processed_file_path, "rb"), as_attachment=True)
                    response["Content-Disposition"] = f'attachment; filename="{processed_file_name}"'
                    return response
                else:
                    print("❌ ERROR: Processed Excel file not found!")
                    message = "❌ Error: Excel file missing after processing!"

            except subprocess.CalledProcessError as e:
                print(f"❌ Subprocess error: {e}")
                message = f"❌ Error during processing: {e}"

    return render(request, "home/upload_telpro_pdf.html", {"message": message})


# ======Laxvan Text File============================================================================================
def upload_laxvan(request):
    print("🚀 Views.Upload_laxvan running")
    message = None
    processed_file_name = "processed_laxvan.xlsx"
    processed_file_path = os.path.join(settings.MEDIA_ROOT, processed_file_name)

    if request.method == "POST" and request.FILES.get("file"):
        uploaded_file = request.FILES["file"]

        # Allow only .txt files
        if not uploaded_file.name.endswith(".xls"):
            message = "Error: Only Excel files are allowed!"
        else:
            # Save uploaded file temporarily
            temp_file_path = default_storage.save("temp_laxvan.txt", ContentFile(uploaded_file.read()))
            full_temp_path = os.path.join(settings.MEDIA_ROOT, temp_file_path)

            print(f"✅ Uploaded file saved at: {full_temp_path}")
            print(f"✅ Processed file should be saved at: {processed_file_path}")

            # Run `telpro_pdf.py` script
            try:
                script_path = os.path.join(settings.BASE_DIR, "spmApp", "laxvan.py")
                subprocess.run([sys.executable, script_path, full_temp_path, processed_file_path], check=True)

                # Check if file exists after processing
                if os.path.exists(processed_file_path):
                    print(f"✅ Processed file FOUND at: {processed_file_path}")

                    # Serve file as an HTTP response for direct download
                    response = FileResponse(open(processed_file_path, "rb"), as_attachment=True)
                    response["Content-Disposition"] = f'attachment; filename="{processed_file_name}"'
                    message = "🚀 Great: File Processed Successfully!"
                    return response
                else:
                    print("❌ ERROR: Processed file is missing after execution!")
                    message = "Error: Processed file not found!"

            except subprocess.CalledProcessError as e:
                message = f"Error processing file: {e}"

    return render(request, "home/upload_laxvan.html", {"message": message})

# ===========Quick Report =============================================================================
def upload_quick_report(request):
    message = None
    processed_file_name = "processed_quick_report.pdf"
    processed_file_path = os.path.join(settings.MEDIA_ROOT, processed_file_name)
    selected_cms_id = request.POST.get("cms_id", "").strip()

    if request.method == "POST" and request.FILES.get("file"):
        uploaded_file = request.FILES["file"]
        full_temp_path = None

        if not uploaded_file.name.endswith(".xlsx"):
            message = "❌ Error: Only Excel files are allowed!"
            return render(request, "home/upload_quick_report.html", {"message": message})

        try:
            # Save uploaded file temporarily
            temp_file_path = default_storage.save(
                "temp_quick_report.xlsx", ContentFile(uploaded_file.read())
            )
            full_temp_path = os.path.join(settings.MEDIA_ROOT, temp_file_path)

            print(f"✅ Uploaded file saved at: {full_temp_path}")
            print(f"📌 Selected CMS_ID: {selected_cms_id}")

            # Run quick_report.py with the uploaded file and selected CMS_ID
            script_path = os.path.join(settings.BASE_DIR, "spmApp", "quick_report.py")
            
            
            # Debug prints
            print(f"Script path: {script_path}")
            print(f"Command: python {script_path} {full_temp_path} {selected_cms_id}")

            # Run the script and capture output
            process = subprocess.run(
                [sys.executable, script_path, full_temp_path, selected_cms_id],
                capture_output=True,
                text=True,
                check=False
            )

            # Print complete debug information
            print("\n=== Script Execution Details ===")
            print(f"Return code: {process.returncode}")
            print("\nStandard Output:")
            print(process.stdout)
            print("\nStandard Error:")
            print(process.stderr)
            print("==============================\n")

            if process.returncode != 0:
                error_msg = process.stderr if process.stderr else process.stdout
                raise Exception(f"Script execution failed with output:\n{error_msg}")

            # Verify the PDF file
            if os.path.exists(processed_file_path):
                file_size = os.path.getsize(processed_file_path)
                print(f"✅ PDF found at: {processed_file_path}")
                print(f"✅ File size: {file_size} bytes")

                if file_size == 0:
                    raise Exception("Generated PDF file is empty")

                # Verify PDF header
                with open(processed_file_path, 'rb') as f:
                    header = f.read(4)
                    if header != b'%PDF':
                        raise Exception("Generated file is not a valid PDF")

                # Serve the file
                with open(processed_file_path, 'rb') as pdf_file:
                    response = HttpResponse(pdf_file.read(), content_type='application/pdf')
                    response['Content-Disposition'] = f'attachment; filename="{processed_file_name}"'
                    response['Content-Length'] = file_size
                    return response
            else:
                raise FileNotFoundError(f"PDF file not found at {processed_file_path}")

        except Exception as e:
            print(f"❌ Detailed Error: {str(e)}")
            message = f"Error processing file: {str(e)}"
            # Log the full traceback
            import traceback
            print("Full traceback:")
            traceback.print_exc()

        finally:
            # Clean up temporary files
            try:
                if full_temp_path and os.path.exists(full_temp_path):
                    os.remove(full_temp_path)
                    print(f"✅ Cleaned up temporary file: {full_temp_path}")
            except Exception as e:
                print(f"❌ Error cleaning up temporary file: {str(e)}")

    return render(request, "home/upload_quick_report.html", {"message": message})



@csrf_exempt  # Allow AJAX requests without CSRF issues
def extract_cms_ids(request):
    if request.method == "POST" and request.FILES.get("file"):
        uploaded_file = request.FILES["file"]

        # Ensure only .xlsx files are processed
        if not uploaded_file.name.endswith(".xlsx"):
            return JsonResponse({"error": "Only .xlsx files are allowed!"}, status=400)

        # Save the uploaded file temporarily
        temp_file_path = default_storage.save("quick_report.pdf", ContentFile(uploaded_file.read()))
        full_temp_path = os.path.join(settings.MEDIA_ROOT, temp_file_path)

        try:
            # Read the Excel file
            df = pd.read_excel(full_temp_path)

            # Extract unique CMS_ID values
            if "CMS_ID" in df.columns:
                cms_ids = df["CMS_ID"].dropna().astype(str).unique().tolist()
            else:
                cms_ids = []

        except Exception as e:
            return JsonResponse({"error": f"Error processing file: {e}"}, status=500)

        finally:
            # Delete temp file after processing
            os.remove(full_temp_path)

        return JsonResponse({"cms_ids": cms_ids})

    return JsonResponse({"error": "Invalid request"}, status=400)