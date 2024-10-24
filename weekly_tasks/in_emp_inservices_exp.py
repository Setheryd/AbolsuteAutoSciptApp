# weekly_tasks/in_emp_inservices_exp.py

import win32com.client as win32  # type: ignore
import os
import urllib.parse
import webbrowser
import traceback
from datetime import datetime, timedelta
from multiprocessing import Process
import time


print("Starting script execution.")

def find_file_in_documents_audit_files(base_path, filename):
    """
    Search for a file in any 'Documents Audit Files' directory under base_path.

    Args:
        base_path (str): The root directory to start searching from.
        filename (str): The exact name of the file to search for.

    Returns:
        str or None: The full path to the found file, or None if not found.
    """
    print(f"Searching for '{filename}' in 'Documents Audit Files' directories under '{base_path}'...")
    for root, dirs, files in os.walk(base_path):
        if root.lower().endswith('documents audit files'):
            print(f"Checking directory: {root}")
            if filename.lower() in (f.lower() for f in files):
                file_path = os.path.join(root, filename)
                print(f"Found file: {file_path}")
                return file_path
    print(f"File '{filename}' not found in any 'Documents Audit Files' directory under '{base_path}'.")
    return None

def get_column_index(ws, header_name):
    """
    Find the column index for a given header name in row 2.

    Args:
        ws: The worksheet object.
        header_name (str): The header name to search for.

    Returns:
        int or None: The column index (1-based) if found, else None.
    """
    print(f"Getting column index for header '{header_name}'...")
    last_col = ws.UsedRange.Columns.Count
    print(f"Last column in used range: {last_col}")
    for col in range(1, last_col + 1):
        cell_value = ws.Cells(2, col).Value
        if cell_value and cell_value.strip().lower() == header_name.strip().lower():
            print(f"Found header '{header_name}' at column {col}")
            return col
    print(f"Header '{header_name}' not found.")
    return None

def process_evaluation_expirations(ws_active):
    """
    Process the 'Active' sheet to find employees who require evaluations (non-blank 'Eval Required' cells).

    Args:
        ws_active: The worksheet object for the 'Active' sheet.

    Returns:
        str: A formatted string listing employee names.
    """
    print("Processing evaluation expirations...")
    employees = []

    # Get the column indices
    name_header = "Last Name, First Name"
    eval_required_header = "In-Services Required"

    name_col = get_column_index(ws_active, name_header)
    eval_col = get_column_index(ws_active, eval_required_header)

    if not name_col or not eval_col:
        print("Required headers not found in 'Active' sheet.")
        return ""

    # Find the last row in the sheet based on the 'Eval Required' column
    last_row_eval = ws_active.Cells(ws_active.Rows.Count, eval_col).End(-4162).Row  # -4162 corresponds to xlUp
    print(f"Last row in 'Active' sheet based on '{eval_required_header}': {last_row_eval}")

    # Loop through each row starting from row 3
    total_employees = 0
    for i in range(3, last_row_eval + 1):
        emp_name = ws_active.Cells(i, name_col).Value
        eval_value = ws_active.Cells(i, eval_col).Value

        if emp_name and eval_value not in (None, "", "-"):
            # Include employee if 'Eval Required' is not blank
            employees.append(emp_name.strip())
            total_employees += 1

    # Remove duplicates and sort the list
    employees = sorted(set(employees))
    print(f"Total employees requiring evaluations: {len(employees)}")
    print("Employee list prepared.")
    employees_str = '\n'.join(employees)
    return employees_str

def get_signature_by_path(signature_path):
    """
    Retrieves the specified Outlook signature as HTML from the provided path.

    Args:
        signature_path (str): The full path to the signature file.

    Returns:
        str: The specified Outlook signature in HTML format, or empty string if not found.
    """
    print(f"Retrieving signature from path: {signature_path}")
    try:
        if os.path.exists(signature_path):
            with open(signature_path, 'r', encoding='utf-8') as f:
                signature = f.read()
                print("Signature retrieved successfully.")
            return signature
        else:
            print(f"Signature file not found at {signature_path}")
    except Exception as e:
        print(f"Failed to get Outlook signature: {e}")
    return ""

def compose_email_classic(employees_str):
    """
    Composes an email via COM automation for classic Outlook.
    """
    try:
        outlookApp = win32.Dispatch('Outlook.Application')
        version = outlookApp.Version
        print(f"Outlook version detected: {version}")
        is_classic = True  # Assume classic Outlook if COM dispatch works
    except Exception as com_exception:
        print(f"COM automation failed: {com_exception}")
        traceback.print_exc()
        outlookApp = None
        is_classic = False

    if outlookApp and is_classic:
        # Use COM automation to compose the email for classic Outlook
        outlookMail = outlookApp.CreateItem(0)
        outlookMail.To = "kaitlyn.moss@absolutecaregivers.com; raegan.lopez@absolutecaregivers.com; ulyana.stokolosa@absolutecaregivers.com"
        outlookMail.CC = "alexander.nazarov@absolutecaregivers.com; luke.kitchel@absolutecaregivers.com"
        outlookMail.Subject = "In-Services Employee In-Services Expiration Reminder"
        print("Email recipients and subject set.")

        # Construct the signature path dynamically
        username = os.getlogin()
        print(f"Current username: {username}")
        signature_filename = "Absolute Signature (seth.riley@absolutecaregivers.com).htm"
        sig_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Signatures', signature_filename)
        print(f"Signature file path: {sig_path}")

        # Get the specified signature
        signature = get_signature_by_path(sig_path)

        # Compose the email body in HTML format
        print("Composing email body...")
        email_body = (
            "<div style='font-family: Calibri, sans-serif; font-size: 11pt;'><p>Dear Team,</p>"
            "<p>I hope this email finds you well. This is an automated reminder regarding the Indianapolis Employee Audit file.</p>"
            "<p>The following Indianapolis employees require In-Services as indicated in the audit file. "
            "Please follow up with them and update the Indy Employee Audit file accordingly. Thank you for your hard work!</p></div>"
            "<ul>"
        )

        # Add each employee name as a list item
        for emp_name in employees_str.split('\n'):
            email_body += f"<li style='font-family: Calibri, sans-serif; font-size: 11pt;'>{emp_name}</li>"
        email_body += "</ul>"

        email_body += "<p>Best regards,</p>"

        # Append the signature if available
        if signature:
            email_body += signature
            print("Signature appended to email body.")
        else:
            email_body += f"<p>{username}<br>Absolute Caregivers</p>"
            print("No signature found; using fallback signature.")

        # Set the email body and format
        outlookMail.HTMLBody = email_body
        print("Email body set.")

        outlookMail.Display()  # Display the email instead of sending
        print("Email composed successfully and displayed for review.")
    else:
        # Raise an exception to indicate failure
        raise Exception("Failed to compose email via COM automation.")

def send_email(employees_str):
    """
    Compose and display an email via Outlook with the list of employees.
    This function attempts to support both the classic and new versions of Outlook.
    If composing via COM automation takes longer than 5 seconds or fails,
    it falls back to using 'mailto'.

    Args:
        employees_str (str): The formatted string of employees.
    """
    print("Preparing to send email...")

    # Run compose_email_classic in a separate process
    try:
        process = Process(target=compose_email_classic, args=(employees_str,))
        process.start()
        process.join(timeout=5)  # Wait up to 5 seconds

        if process.is_alive():
            print("Composing email via COM automation took too long, terminating process.")
            process.terminate()
            process.join()
            raise Exception("Timeout composing email via COM automation.")
        else:
            print("Email composed via COM automation successfully.")
            return  # Exit the function, since email is composed
    except Exception as e:
        print(f"Exception during composing email via COM automation: {e}")
        # Proceed to fallback method

    # Fallback method for new Outlook or if COM automation fails
    print("Using fallback method to compose email.")

    # Prepare email components
    to_addresses = "kaitlyn.moss@absolutecaregivers.com; raegan.lopez@absolutecaregivers.com; ulyana.stokolosa@absolutecaregivers.com"
    cc_addresses = "alexander.nazarov@absolutecaregivers.com; luke.kitchel@absolutecaregivers.com"
    subject = "In-Services Employee In-Services Expiration Reminder"

    # Build the body text
    body = (
        "Dear Team,\n\n"
        "I hope this email finds you well. This is an automated reminder regarding the Indianapolis Employee Audit file.\n\n"
        "The following Indianapolis employees require In-Services as indicated in the audit file. "
        "Please follow up with them and update the Indy Employee Audit file accordingly. Thank you for your hard work!\n\n"
    )
    body += employees_str
    body += "\n\nBest regards,\n"
    body += f"{os.getlogin()}\nAbsolute Caregivers"

    # Create the mailto link
    mailto_link = f"mailto:{urllib.parse.quote(to_addresses)}"
    mailto_link += f"?cc={urllib.parse.quote(cc_addresses)}"
    mailto_link += f"&subject={urllib.parse.quote(subject)}"
    mailto_link += f"&body={urllib.parse.quote(body)}"

    print(f"Mailto link: {mailto_link}")

    # Open the mailto link
    webbrowser.open(mailto_link)
    print("Email composed using 'mailto' and opened in default email client.")


def extract_evaluation_expirations():
    """
    Main function to extract employees requiring evaluations and send an email report.
    """
    print("Starting extraction of evaluation expirations...")
    try:
        username = os.getlogin()
        print(f"Current username: {username}")
    except Exception as e:
        print(f"Failed to get the current username: {e}")
        return

    base_path = f"C:\\Users\\{username}\\OneDrive - Ability Home Health, LLC\\"
    print(f"Base path set to: {base_path}")

    # Define the exact filename to search for
    audit_filename = "Employee Audit Checklist.xlsm"
    print(f"Looking for file: {audit_filename}")

    # Find the file in any 'Documents Audit Files' directory
    audit_file = find_file_in_documents_audit_files(base_path, audit_filename)

    if not audit_file:
        print("Required file not found in specified directories.")
        return

    print(f"Found {audit_filename} at: {audit_file}")

    try:
        print("Initializing Excel application...")
        excel = win32.DispatchEx("Excel.Application")  # Use DispatchEx
        excel.DisplayAlerts = False
        excel.Visible = False
        print("Excel application initialized.")

        # Open the Audit Workbook
        print(f"Opening workbook: {audit_file}")
        wb_audit = excel.Workbooks.Open(audit_file, False, True, None, "abs$1004$N")
        ws_active = wb_audit.Sheets("Active")
        print("Employee Audit Checklist workbook opened successfully.")
    except Exception as e:
        print(f"Failed to open {audit_filename}: {e}")
        return

    # Process the active sheet to find employees requiring evaluations
    employees_str = process_evaluation_expirations(ws_active)

    # Close the workbook without saving
    try:
        print("Closing workbook...")
        wb_audit.Close(SaveChanges=False)
        excel.Quit()
        del excel
        print("Workbook closed successfully.")
    except Exception as e:
        print(f"Failed to close workbook: {e}")

    # Send email if there are employees requiring evaluations
    if employees_str.strip():
        print("Employees requiring evaluations found. Sending email...")
        send_email(employees_str)
    else:
        print("No employees requiring evaluations found.")

def run_task():
    """
    Wrapper function to execute the extract_evaluation_expirations function.
    Returns the result string or raises an exception.
    """
    print("Running task...")
    try:
        result = extract_evaluation_expirations()
        print("Task completed successfully.")
        return result
    except Exception as e:
        print(f"Error occurred during task execution: {e}")
        raise e

if __name__ == "__main__":
    print("Script execution started.")
    extract_evaluation_expirations()
    print("Script execution finished.")
