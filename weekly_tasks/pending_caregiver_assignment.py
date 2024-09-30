# weekly_tasks/pending_admission_caregiver.py

import win32com.client as win32  # type: ignore
import os
from datetime import datetime
import sys

def find_specific_file(base_path, filename, required_subpath, max_depth=5):
    """
    Search for a specific file within directories that contain a required subpath.

    Args:
        base_path (str): The root directory to start searching from.
        filename (str): The exact name of the file to search for.
        required_subpath (str): A substring that must be present in the file's directory path.
        max_depth (int): Maximum depth to search within subdirectories.

    Returns:
        str or None: The full path to the found file, or None if not found.
    """
    print(f"Searching for {filename} within directories containing '{required_subpath}' in {base_path}...")

    def scan_directory(path, current_depth):
        # Stop recursion if the current depth exceeds max_depth
        if current_depth > max_depth:
            return None

        try:
            with os.scandir(path) as it:
                for entry in it:
                    if entry.is_file() and entry.name.lower() == filename.lower():
                        # Check if the required_subpath is in the file's directory path
                        if required_subpath.lower() in os.path.dirname(entry.path).lower():
                            print(f"Found {filename} at: {entry.path}")
                            return entry.path
                        else:
                            print(f"Found {filename} at {entry.path}, but it does not contain '{required_subpath}' in its path. Skipping.")
                    elif entry.is_dir():
                        found_file = scan_directory(entry.path, current_depth + 1)
                        if found_file:
                            return found_file
        except PermissionError:
            # Skip directories that are not accessible
            print(f"Permission denied: {path}")
            return None

    return scan_directory(base_path, 0)

def get_signature_by_path(signature_path):
    """
    Retrieves the specified Outlook signature as HTML from the provided path.

    Args:
        signature_path (str): The full path to the signature file.

    Returns:
        str: The specified Outlook signature in HTML format, or empty string if not found.
    """
    try:
        if os.path.exists(signature_path):
            with open(signature_path, 'r', encoding='utf-8') as f:
                signature = f.read()
            return signature
        else:
            print(f"Signature file not found at {signature_path}")
    except Exception as e:
        print(f"Failed to get Outlook signature: {e}")
    return ""

def extract_pending_admissions():
    """
    Extracts a list of patients with pending admissions from the Excel workbook.

    Returns:
        list: A list of patient names with pending admissions.
    """
    # Capture the current user's username
    try:
        username = os.getlogin()
    except Exception as e:
        print(f"Failed to get the current username: {e}")
        return []

    print(f"Current username: {username}")

    # Construct the base path using the username
    base_path = f"C:\\Users\\{username}\\OneDrive - Ability Home Health, LLC\\"
    print(f"Base path: {base_path}")

    # Check if base path exists
    if not os.path.exists(base_path):
        print(f"Base path {base_path} does not exist.")
        return []

    # Define the filename to search for
    filename = "Absolute Patient Records.xlsm"
    required_subpath = "Absolute Operation"

    # Find the file in the specified subdirectory
    # The expected path: C:\Users\{username}\OneDrive - Ability Home Health, LLC\{unknown}\Absolute Operation\Absolute Patient Records.xlsm
    file_path = find_specific_file(base_path, filename, required_subpath)
    if file_path is None:
        print(f"File {filename} not found within directories containing '{required_subpath}' starting from {base_path}.")
        return []

    password = "abs$1018$B"

    try:
        # Initialize Excel application object using DispatchEx
        excel = win32.DispatchEx("Excel.Application")  # Changed from Dispatch to DispatchEx
        excel.DisplayAlerts = False
        excel.Visible = False
        print("Excel application object created successfully.")

        # Open the workbook in read-only mode with password
        # Pass parameters positionally up to Password
        # Excel.Workbooks.Open(Filename, UpdateLinks, ReadOnly, Format, Password, ...)
        wb = excel.Workbooks.Open(file_path, False, True, None, password)  # Changed to positional parameters
        print(f"Workbook {filename} opened successfully.")

        # Access the 'Patient Information' sheet
        ws = wb.Sheets("Patient Information")
        print(f"Accessed 'Patient Information' sheet in {filename}.")

        # Identify the required columns by header names in the first row
        headers = {}
        last_col = ws.UsedRange.Columns.Count
        for col in range(1, last_col + 1):
            header = ws.Cells(1, col).Value
            if header:
                headers[header.strip().lower()] = col

        # Updated required headers to include "Care Giver"
        required_headers = [
            "admission date (mm/dd/yyyy)", 
            "name (last name, first name)", 
            "discharge date",
            "care giver"  # Added "Care Giver" to the required headers
        ]

        for req_header in required_headers:
            if req_header not in headers:
                print(f"Required column '{req_header}' not found in the sheet.")
                wb.Close(SaveChanges=False)
                excel.Quit()
                del excel
                return []

        admission_date_col = headers["admission date (mm/dd/yyyy)"]
        name_col = headers["name (last name, first name)"]
        discharge_date_col = headers["discharge date"]
        care_giver_col = headers["care giver"]  # Added extraction of "Care Giver" column

        print(f"Admission Date column: {admission_date_col}")
        print(f"Name column: {name_col}")
        print(f"Discharge Date column: {discharge_date_col}")
        print(f"Care Giver column: {care_giver_col}")  # Print "Care Giver" column index

        # Determine the last row with data in the Admission Date column
        last_row = ws.Cells(ws.Rows.Count, admission_date_col).End(-4162).Row  # -4162 corresponds to xlUp

        print(f"Last row with data: {last_row}")

        pending_admissions = []

        for row in range(2, last_row + 1):
            # Fetch values from the relevant columns
            admission_date = ws.Cells(row, admission_date_col).Value
            discharge_date = ws.Cells(row, discharge_date_col).Value
            patient_name = ws.Cells(row, name_col).Value
            care_giver = ws.Cells(row, care_giver_col).Value  # Fetch "Care Giver" value

            if patient_name:
                # Updated condition: Check if both Discharge Date and Care Giver are blank
                discharge_blank = discharge_date in (None, "")
                care_giver_blank = care_giver in (None, "")
                if discharge_blank and care_giver_blank:
                    pending_admissions.append(patient_name.strip())

        # Close the workbook without saving
        wb.Close(SaveChanges=False)
        print(f"Workbook {filename} closed successfully.")

        # Quit Excel application
        excel.Quit()
        print("Excel application closed successfully.")

        # Release COM objects
        del excel

    except Exception as e:
        print(f"An error occurred while processing {file_path}: {e}")
        return []

    print(f"Number of patients with pending admissions: {len(pending_admissions)}")
    return pending_admissions

def send_email(pending_admissions):
    """
    Compose and send an email via Outlook with the list of patients pending admission.

    Args:
        pending_admissions (list): The list of patient names with pending admissions.
    """
    try:
        outlookApp = win32.DispatchEx('Outlook.Application')  # Changed from Dispatch to DispatchEx
        outlookMail = outlookApp.CreateItem(0)
        outlookMail.To = "kaitlyn.moss@absolutecaregivers.com; raegan.lopez@absolutecaregivers.com; ulyana.stokolosa@absolutecaregivers.com; Liliia.Reshetnyk@absolutecaregivers.com"
        outlookMail.CC = "alexander.nazarov@absolutecaregivers.com; luke.kitchel@absolutecaregivers.com"
        outlookMail.Subject = "Pending Caregiver Assignment Reminder"

        # Construct the signature path dynamically
        username = os.getlogin()
        signature_filename = "Absolute Signature (seth.riley@absolutecaregivers.com).htm"
        sig_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Signatures', signature_filename)

        # Get the specified signature
        signature = get_signature_by_path(sig_path)

        # Compose the email body in HTML format
        email_body = (
            "<div style='font-family: Calibri, sans-serif; font-size: 11pt;'><p>Dear Team,</p>"
            "<p>I hope this message finds you well.</p>"
            "<p>This is an automated reminder regarding pending caregiver assignments. The following patients are labeled as active and do not have a caregiver assigned to them yet:</p>"
            "<ul>"
        )

        # Add each patient name as a list item
        for patient in pending_admissions:
            email_body += f"<div style='font-family: Calibri, sans-serif; font-size: 11pt;'><li>{patient}</li>"
        email_body += "</ul>"

        email_body += (
            "<div style='font-family: Calibri, sans-serif; font-size: 11pt;'><p>Please review and update the patient records file accordingly. If you notice any discrepancies or have already addressed these admissions, kindly update the records to maintain accuracy.</p>"
            "<p>Thank you for your prompt attention to this matter.</p>"
            "<p>Best regards,</p></div>"
        )

        # Append the signature if available
        if signature:
            email_body += signature
        else:
            email_body += "<p>{username}<br>Absolute Caregivers</p>"

        # Set the email body and format
        outlookMail.HTMLBody = email_body

        outlookMail.Display()  # Change to .Send() to send automatically without displaying
        print("Email composed successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")

def main():
    """
    Main function to extract pending admissions and send an email report.
    """
    pending_admissions = extract_pending_admissions()
    if pending_admissions:
        send_email(pending_admissions)
    else:
        print("No patients with pending admissions found.")

def run_task():
    """
    Wrapper function to execute the main function.
    Returns the result string or raises an exception.
    """
    try:
        result = main()
        return result
    except Exception as e:
        raise e

if __name__ == "__main__":
    main()
