# weekly_tasks/pending_IHCC_admission.py

import win32com.client as win32 # type: ignore
import os
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
    print(f"Searching for '{filename}' within directories containing '{required_subpath}' in '{base_path}'...")

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
                            print(f"Found '{filename}' at: {entry.path}")
                            return entry.path
                        else:
                            print(f"Found '{filename}' at '{entry.path}', but it does not contain '{required_subpath}' in its path. Skipping.")
                    elif entry.is_dir():
                        found_file = scan_directory(entry.path, current_depth + 1)
                        if found_file:
                            return found_file
        except PermissionError:
            # Skip directories that are not accessible
            print(f"Permission denied: '{path}'")
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
            print(f"Signature file not found at '{signature_path}'")
    except Exception as e:
        print(f"Failed to get Outlook signature: {e}")
    return ""

def extract_pending_admissions():
    """
    Extracts a list of IHCC patients with pending admissions from the Excel workbook.

    Returns:
        list: A list of patient names with pending admissions.
    """
    # Capture the current user's username
    try:
        username = os.getlogin()
    except Exception as e:
        print(f"Failed to get the current username: {e}")
        return []

    print(f"Current username: '{username}'")

    # Construct the base path using the username
    base_path = f"C:\\Users\\{username}\\OneDrive - Ability Home Health, LLC\\"
    print(f"Base path: '{base_path}'")

    # Check if base path exists
    if not os.path.exists(base_path):
        print(f"Base path '{base_path}' does not exist.")
        return []

    # Define the filename to search for
    filename = "Absolute Patient Records IHCC.xlsm"
    required_subpath = "IHCC"

    # Find the file in the specified subdirectory
    # The expected path: C:\Users\{username}\OneDrive - Ability Home Health, LLC\{unknown}\IHCC\Absolute Patient Records IHCC.xlsm
    file_path = find_specific_file(base_path, filename, required_subpath)
    if file_path is None:
        print(f"File '{filename}' not found within directories containing '{required_subpath}' starting from '{base_path}'.")
        return []

    # It's recommended to use environment variables or secure storage for passwords
    # For demonstration purposes, it's hardcoded here
    password = "abs$1018$B"

    excel = None
    wb = None

    try:
        # Initialize Excel application object
        excel = win32.Dispatch("Excel.Application")
        excel.DisplayAlerts = False  # Disable Excel alerts
        excel.Visible = False        # Make Excel invisible
        excel.ScreenUpdating = False # Prevent screen updates
        print("Excel application object created successfully.")

        # Open the workbook in read-only mode with password
        wb = excel.Workbooks.Open(file_path, Password=password, ReadOnly=True)
        print(f"Workbook '{filename}' opened successfully.")

        # Access the 'Patient Information' sheet
        ws = wb.Sheets("Patient Information")
        print(f"Accessed 'Patient Information' sheet in '{filename}'.")

        # Identify the required columns by header names in the first row
        headers = {}
        last_col = ws.UsedRange.Columns.Count
        for col in range(1, last_col + 1):
            header = ws.Cells(1, col).Value
            if header:
                headers[header.strip().lower()] = col

        # Define required headers in lowercase to match the keys in 'headers'
        required_headers = ["ihcc admission date", "name", "discharge date"]

        missing_headers = [req for req in required_headers if req.lower() not in headers]
        if missing_headers:
            print(f"Required columns {missing_headers} not found in the sheet.")
            return []

        admission_date_col = headers["ihcc admission date"]
        name_col = headers["name"]
        discharge_date_col = headers["discharge date"]

        print(f"Admission Date column index: {admission_date_col}")
        print(f"Name column index: {name_col}")
        print(f"Discharge Date column index: {discharge_date_col}")

        # Determine the last row with data in the Admission Date column
        last_row = ws.Cells(ws.Rows.Count, admission_date_col).End(-4162).Row  # -4162 corresponds to xlUp
        print(f"Last row with data: {last_row}")

        pending_admissions = []

        for row in range(2, last_row + 1):
            admission_date = ws.Cells(row, admission_date_col).Value
            discharge_date = ws.Cells(row, discharge_date_col).Value
            patient_name = ws.Cells(row, name_col).Value

            if patient_name:
                admission_blank = admission_date in (None, "")
                discharge_blank = discharge_date in (None, "")
                if admission_blank and discharge_blank:
                    pending_admissions.append(patient_name.strip())

        print(f"Number of IHCC patients with pending admissions: {len(pending_admissions)}")
        return pending_admissions

    except Exception as e:
        print(f"An error occurred while processing '{file_path}': {e}")
        return []

    finally:
        # Ensure that the workbook and Excel application are properly closed
        if wb:
            wb.Close(SaveChanges=False)
            print(f"Workbook '{filename}' closed successfully.")
        if excel:
            excel.Quit()
            print("Excel application closed successfully.")
        # Release COM objects
        if excel:
            del excel

def send_email(pending_admissions):
    """
    Compose and send an email via Outlook with the list of IHCC patients pending admission.

    Args:
        pending_admissions (list): The list of patient names with pending admissions.
    """
    outlook = None
    mail = None

    try:
        # Initialize Outlook application object
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)  # 0: olMailItem

        mail.To = "lyudmila.slepaya@absolutecaregivers.com; ulyana.stokolosa@absolutecaregivers.com; victoria.shmoel@absolutecaregivers.com "
        mail.CC = "alexander.nazarov@absolutecaregivers.com; luke.kitchel@absolutecaregivers.com"
        mail.Subject = "Pending IHCC Admissions"

        # Construct the signature path dynamically
        signature_filename = "Absolute Signature (seth.riley@absolutecaregivers.com).htm"
        sig_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Signatures', signature_filename)

        # Get the specified signature
        signature = get_signature_by_path(sig_path)

        # Compose the email body in HTML format
        email_body = (
            "<p>Dear Team,</p>"
            "<p>I hope this message finds you well.</p>"
            "<p>This is an automated reminder regarding pending IHCC admissions. The following patients currently do not have an admission date for their IHCC services, and require your immediate attention:</p>"
            "<ul>"
        )

        # Add each patient name as a list item
        for patient in pending_admissions:
            email_body += f"<li>{patient}</li>"
        email_body += "</ul>"

        email_body += (
            "<p>Please review and update the patient records file accordingly. If you notice any discrepancies or have already addressed these admissions, kindly update the records to maintain accuracy.</p>"
            "<p>Thank you for your prompt attention to this matter.</p>"
            "<p>Best regards,</p>"
        )

        # Append the signature if available
        if signature:
            email_body += signature
        else:
            email_body += "<p>Your Name<br>Your Title<br>Absolute Caregivers</p>"

        # Set the email body and format
        mail.HTMLBody = email_body

        # Display the email for manual curation before sending
        mail.Display(False)  # False to open the email without modal dialog
        print("Email displayed successfully.")

    except Exception as e:
        print(f"Failed to send email: {e}")

    finally:
        # Release COM objects
        if mail:
            del mail
        if outlook:
            del outlook

def main():
    """
    Main function to extract pending IHCC admissions and send an email report.
    """
    pending_admissions = extract_pending_admissions()
    if pending_admissions:
        send_email(pending_admissions)
    else:
        print("No IHCC patients with pending admissions found.")

if __name__ == "__main__":
    main()
