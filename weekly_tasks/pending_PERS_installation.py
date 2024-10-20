# weekly_tasks/pending_PERS_Installation.py

import win32com.client as win32  # type: ignore
import os
import sys
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(
    filename='pending_PERS_Installation.log',
    level=logging.INFO,
    format='%(asctime)s:%(levelname)s:%(message)s'
)

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
    logging.info(f"Searching for '{filename}' within directories containing '{required_subpath}' in '{base_path}'...")
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
                            logging.info(f"Found '{filename}' at: {entry.path}")
                            print(f"Found '{filename}' at: {entry.path}")
                            return entry.path
                        else:
                            logging.info(f"Found '{filename}' at '{entry.path}', but it does not contain '{required_subpath}' in its path. Skipping.")
                            print(f"Found '{filename}' at '{entry.path}', but it does not contain '{required_subpath}' in its path. Skipping.")
                    elif entry.is_dir():
                        found_file = scan_directory(entry.path, current_depth + 1)
                        if found_file:
                            return found_file
        except PermissionError:
            # Skip directories that are not accessible
            logging.warning(f"Permission denied: '{path}'")
            print(f"Permission denied: '{path}'")
            return None
        except Exception as e:
            logging.error(f"Error accessing '{path}': {e}")
            print(f"Error accessing '{path}': {e}")
            return None

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
            logging.info(f"Signature found at '{signature_path}'")
            return signature
        else:
            logging.warning(f"Signature file not found at '{signature_path}'")
            print(f"Signature file not found at '{signature_path}'")
    except Exception as e:
        logging.error(f"Failed to get Outlook signature: {e}")
        print(f"Failed to get Outlook signature: {e}")
    return ""


def extract_pending_admissions():
    """
    Extracts a list of patients with pending PERS installations from the Excel workbook.

    Returns:
        list: A list of patient names with pending PERS installations.
    """
    # Capture the current user's username
    try:
        username = os.getlogin()
    except Exception as e:
        logging.error(f"Failed to get the current username: {e}")
        print(f"Failed to get the current username: {e}")
        return []

    logging.info(f"Current username: '{username}'")
    print(f"Current username: '{username}'")

    # Construct the base path using the username
    base_path = f"C:\\Users\\{username}\\OneDrive - Ability Home Health, LLC\\"
    logging.info(f"Base path: '{base_path}'")
    print(f"Base path: '{base_path}'")

    # Check if base path exists
    if not os.path.exists(base_path):
        logging.warning(f"Base path '{base_path}' does not exist.")
        print(f"Base path '{base_path}' does not exist.")
        return []

    # Define the filename to search for
    filename = "Absolute Patient Records PERS.xlsm"
    required_subpath = "IHCC"

    # Find the file in the specified subdirectory
    # The expected path: C:\Users\{username}\OneDrive - Ability Home Health, LLC\{unknown}\IHCC\Absolute Patient Records PERS.xlsm
    file_path = find_specific_file(base_path, filename, required_subpath)
    if file_path is None:
        logging.warning(f"File '{filename}' not found within directories containing '{required_subpath}' starting from '{base_path}'.")
        print(f"File '{filename}' not found within directories containing '{required_subpath}' starting from '{base_path}'.")
        return []

    # It's recommended to use environment variables or secure storage for passwords
    # For demonstration purposes, it's hardcoded here
    password = os.getenv('EXCEL_PASSWORD', 'abs$1018$B')  # Replace with environment variable if possible

    excel = None
    wb = None

    try:
        # Initialize Excel application object using DispatchEx
        excel = win32.DispatchEx("Excel.Application")  # Changed from Dispatch to DispatchEx
        excel.DisplayAlerts = False  # Disable Excel alerts
        excel.Visible = False        # Make Excel invisible
        excel.ScreenUpdating = False # Prevent screen updates
        excel.EnableEvents = False    # Disable events
        logging.info("Excel application object created successfully.")
        print("Excel application object created successfully.")

        # Open the workbook in read-only mode with password
        # Pass parameters positionally up to Password
        # Excel.Workbooks.Open(Filename, UpdateLinks, ReadOnly, Format, Password, ...)
        wb = excel.Workbooks.Open(file_path, False, True, None, password)  # Changed to positional parameters
        logging.info(f"Workbook '{filename}' opened successfully.")
        print(f"Workbook '{filename}' opened successfully.")

        # Access the 'Patient Information' sheet
        ws = wb.Sheets("Patient Information")
        logging.info(f"Accessed 'Patient Information' sheet in '{filename}'.")
        print(f"Accessed 'Patient Information' sheet in '{filename}'.")

        # Identify the required columns by header names in the first row
        headers = {}
        last_col = ws.UsedRange.Columns.Count
        for col in range(1, last_col + 1):
            header = ws.Cells(1, col).Value
            if header:
                headers[header.strip().lower()] = col

        # Define required headers in lowercase to match the keys in 'headers'
        required_headers = ["installation date", "name", "discharge date"]

        # Check for missing headers
        missing_headers = [req for req in required_headers if req.lower() not in headers]
        if missing_headers:
            logging.warning(f"Required columns {missing_headers} not found in the sheet.")
            print(f"Required columns {missing_headers} not found in the sheet.")
            wb.Close(SaveChanges=False)
            return []

        installation_date_col = headers["installation date"]
        name_col = headers["name"]
        discharge_date_col = headers["discharge date"]

        logging.info(f"Installation Date column index: {installation_date_col}")
        logging.info(f"Name column index: {name_col}")
        logging.info(f"Discharge Date column index: {discharge_date_col}")

        print(f"Installation Date column index: {installation_date_col}")
        print(f"Name column index: {name_col}")
        print(f"Discharge Date column index: {discharge_date_col}")

        # Determine the last row with data in the Installation Date column
        last_row = ws.Cells(ws.Rows.Count, installation_date_col).End(-4162).Row  # -4162 corresponds to xlUp
        logging.info(f"Last row with data: {last_row}")
        print(f"Last row with data: {last_row}")

        pending_admissions = []

        for row in range(2, last_row + 1):
            installation_date = ws.Cells(row, installation_date_col).Value
            discharge_date = ws.Cells(row, discharge_date_col).Value
            patient_name = ws.Cells(row, name_col).Value

            if patient_name:
                installation_blank = installation_date in (None, "")
                discharge_blank = discharge_date in (None, "")
                if installation_blank and discharge_blank:
                    pending_admissions.append(patient_name.strip())

        logging.info(f"Number of patients with pending PERS installations: {len(pending_admissions)}")
        print(f"Number of patients with pending PERS installations: {len(pending_admissions)}")
        return pending_admissions

    except Exception as e:
        logging.error(f"An error occurred while processing '{file_path}': {e}")
        print(f"An error occurred while processing '{file_path}': {e}")
        return []

    finally:
        # Ensure that the workbook and Excel application are properly closed
        if wb:
            wb.Close(SaveChanges=False)
            logging.info(f"Workbook '{filename}' closed successfully.")
            print(f"Workbook '{filename}' closed successfully.")
        if excel:
            excel.Quit()
            logging.info("Excel application closed successfully.")
            print("Excel application closed successfully.")
        # Release COM objects
        if excel:
            del excel


def send_email(pending_admissions):
    """
    Compose and send an email via Outlook with the list of patients pending PERS installations.

    Args:
        pending_admissions (list): The list of patient names with pending PERS installations.
    """
    if not pending_admissions:
        print("No PERS patients with pending installations found. Email will not be sent.")
        return

    outlook = None
    mail = None

    try:
        # Initialize Outlook application object using DispatchEx
        outlook = win32.DispatchEx('Outlook.Application')  # Changed from Dispatch to DispatchEx
        mail = outlook.CreateItem(0)  # 0: olMailItem

        # Corrected email recipient formatting by removing trailing space
        mail.To = "ilya.karpov@absolutecaregivers.com;raegan.lopez@absolutecaregivers.com;ulyana.stokolosa@absolutecaregivers.com;victoria.shmoel@absolutecaregivers.com"
        mail.CC = "alexander.nazarov@absolutecaregivers.com;luke.kitchel@absolutecaregivers.com"
        mail.Subject = "Pending PERS Installations"

        # Construct the signature path dynamically
        signature_filename = "Absolute Signature (seth.riley@absolutecaregivers.com).htm"
        sig_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Signatures', signature_filename)

        # Get the specified signature
        signature = get_signature_by_path(sig_path)

        # Compose the email body in HTML format
        email_body = (
            "<div style='font-family: Calibri, sans-serif; font-size: 11pt;'><p>Dear Team,</p>"
            "<p>I hope this message finds you well.</p>"
            "<p>This is an automated reminder regarding pending PERS Installations. The following patients currently do not have an installation date and discharge date, and require your immediate attention:</p>"
            "<ul>"
        )

        # Add each patient name as a list item
        for patient in pending_admissions:
            email_body += f"<li style='font-family: Calibri, sans-serif; font-size: 11pt;'>{patient}</li>"
        email_body += "</ul></div>"

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
        logging.info("Email displayed successfully.")
        print("Email displayed successfully.")
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        print(f"Failed to send email: {e}")
    finally:
        # Release COM objects
        if mail:
            del mail
        if outlook:
            del outlook


def main():
    """
    Main function to extract pending PERS installations and send an email report.
    """
    pending_admissions = extract_pending_admissions()
    if pending_admissions:
        send_email(pending_admissions)
    else:
        logging.info("No PERS patients with pending installations found.")
        print("No PERS patients with pending installations found.")


def run_task():
    """
    Wrapper function to execute the main function.
    Returns the result string or raises an exception.
    """
    try:
        result = main()
        return result
    except Exception as e:
        logging.error(f"An error occurred in run_task: {e}")
        raise e


if __name__ == "__main__":
    main()
