# weekly_tasks/indy_emp_eval.py

import win32com.client as win32  # type: ignore
import os
from datetime import datetime, timedelta
import sys
from multiprocessing import Process
import urllib.parse
import webbrowser
from bs4 import BeautifulSoup  # For parsing HTML signatures
import logging

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
    print(f"find_file_in_documents_audit_files called with base_path: {base_path}, filename: {filename}")
    for root, dirs, files in os.walk(base_path):
        if root.lower().endswith('documents audit files'):
            print(f"Searching in directory: {root}")
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
    print(f"get_column_index called with header_name: '{header_name}'")
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
    print("process_evaluation_expirations called.")
    employees = []

    # Get the column indices
    name_header = "Last Name, First Name"
    eval_required_header = "Eval Required"

    name_col = get_column_index(ws_active, name_header)
    eval_col = get_column_index(ws_active, eval_required_header)

    if not name_col or not eval_col:
        print("Required headers not found in 'Active' sheet.")
        return ""

    # Find the last row in the sheet based on the 'Eval Required' column
    last_row_eval = ws_active.Cells(ws_active.Rows.Count, eval_col).End(-4162).Row  # -4162 corresponds to xlUp
    print(f"Last row in 'Active' sheet based on '{eval_required_header}': {last_row_eval}")

    # Loop through each row starting from row 3
    print("Processing rows to find employees requiring evaluations...")
    for i in range(3, last_row_eval + 1):
        emp_name = ws_active.Cells(i, name_col).Value
        eval_value = ws_active.Cells(i, eval_col).Value

        if emp_name and eval_value not in (None, "", "-"):
            # Include employee if 'Eval Required' is not blank
            employees.append(emp_name.strip())

    # Remove duplicates and sort the list
    employees = sorted(set(employees))
    total_employees = len(employees)
    print(f"Total employees requiring evaluations: {total_employees}")

    # Concatenate employee names into a single string
    employees_str = '\n'.join(employees)
    print("Employee list prepared.")
    return employees_str

def get_signature_by_path(signature_path):
    """
    Retrieves the specified Outlook signature as HTML from the provided path.

    Args:
        signature_path (str): The full path to the signature file.

    Returns:
        str: The specified Outlook signature in HTML format, or empty string if not found.
    """
    print(f"get_signature_by_path called with signature_path: {signature_path}")
    try:
        if os.path.exists(signature_path):
            print(f"Signature file found at {signature_path}")
            with open(signature_path, 'r', encoding='utf-8') as f:
                signature = f.read()
            print("Signature retrieved successfully.")
            return signature
        else:
            print(f"Signature file not found at {signature_path}")
    except Exception as e:
        print(f"Failed to get Outlook signature: {e}")
    return ""

def get_default_outlook_email():
    """
    Retrieves the default Outlook email address of the current user.

    Returns:
        str: The default email address if available, otherwise None.
    """
    print("get_default_outlook_email called.")
    try:
        outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        accounts = namespace.Accounts
        if accounts.Count > 0:
            # Outlook accounts are 1-indexed
            default_account = accounts.Item(1)
            email = default_account.SmtpAddress
            print(f"Default Outlook email: {email}")
            return email
        else:
            logging.error("No Outlook accounts found.")
            return None
    except Exception as e:
        logging.error(f"Unable to retrieve default Outlook email: {e}")
        return None

def get_default_signature():
    """
    Retrieves the user's default email signature based on their default Outlook account.

    Returns:
        str: The signature HTML content if available, otherwise None.
    """
    print("get_default_signature called.")
    email = get_default_outlook_email()
    if not email:
        logging.error("Default Outlook email not found.")
        return None

    # Define the signature directory
    appdata = os.environ.get('APPDATA')
    if not appdata:
        logging.error("APPDATA environment variable not found.")
        return None

    sig_dir = os.path.join(appdata, 'Microsoft', 'Signatures')
    if not os.path.isdir(sig_dir):
        logging.error(f"Signature directory does not exist: {sig_dir}")
        return None

    # Iterate through signature files to find a match
    for filename in os.listdir(sig_dir):
        if filename.lower().endswith(('.htm', '.html')):
            # Extract the base name without extension
            base_name = os.path.splitext(filename)[0].lower()
            if email.lower() in base_name:
                sig_path = os.path.join(sig_dir, filename)
                signature = get_signature_by_path(sig_path)
                if signature:
                    logging.info(f"Signature found: {sig_path}")
                    return signature

    logging.error(f"No signature file found containing email: {email}")
    return None

def compose_email_classic(employees_str):
    """
    Composes and displays an email via COM automation for classic Outlook.
    """
    print("compose_email_classic called.")
    try:
        print("Initializing Outlook application...")
        outlookApp = win32.Dispatch("Outlook.Application")
        outlookMail = outlookApp.CreateItem(0)
        print("Outlook application initialized.")

        # Define recipients
        outlookMail.To = "kaitlyn.moss@absolutecaregivers.com; raegan.lopez@absolutecaregivers.com; ulyana.stokolosa@absolutecaregivers.com"
        outlookMail.CC = "alexander.nazarov@absolutecaregivers.com; luke.kitchel@absolutecaregivers.com"
        outlookMail.Subject = "Indianapolis Employee Evaluation Reminder"
        print("Email recipients and subject set.")

        # Get the default signature
        signature = get_default_signature()

        # Compose the email body in HTML format
        print("Composing email body...")
        email_body = (
            "<div style='font-family: Calibri, sans-serif; font-size: 11pt;'><p>Dear Kaitlyn,</p>"
            "<p>I hope this email finds you well. This is an automated reminder regarding the Indianapolis Employee Audit file.</p>"
            "<p>The following Indianapolis employees require evaluations as indicated in the audit file. "
            "Please follow up with them and update the Indy Employee Audit file accordingly. Thank you for your hard work!</p></div>"
            "<ul>"
        )

        # Add each employee name as a list item
        for emp_name in employees_str.split('\n'):
            email_body += f"<li>{emp_name}</li>"
        email_body += "</ul>"

        email_body += "<p>Best regards,</p>"

        # Append the signature if available
        if signature:
            print("Appending signature to email body.")
            email_body += signature
        else:
            print("No signature found; using fallback signature.")
            username = os.getlogin()
            email_body += f"<p>{username}<br>Absolute Caregivers</p>"

        # Set the email body and format
        outlookMail.HTMLBody = email_body
        print("Email body set.")

        # Display the email for review
        print("Displaying email for review...")
        outlookMail.Display()  # Change to .Send() to send automatically without displaying
        print("Email composed successfully.")
    except Exception as e:
        logging.error(f"Failed to compose or display email via COM automation: {e}")
        raise

def send_email(employees_str):
    """
    Composes and sends an email with a 5-second timeout for COM automation.
    Falls back to using a mailto link if the timeout is exceeded.
    """
    print("send_email called.")
    # Try composing email via COM automation with a timeout
    try:
        process = Process(target=compose_email_classic, args=(employees_str,))
        process.start()
        process.join(timeout=5)  # Wait up to 5 seconds

        if process.is_alive():
            logging.warning("Composing email via COM automation took too long, terminating process.")
            process.terminate()
            process.join()
            raise Exception("Timeout composing email via COM automation.")
        else:
            logging.info("Email composed via COM automation successfully.")
            return  # Exit the function, as the email has been composed successfully
    except Exception as e:
        logging.error(f"Exception during composing email via COM automation: {e}")
        # Proceed to fallback method

    # Fallback method using 'mailto' link
    logging.info("Using fallback method to compose email.")
    print("Using fallback method to compose email.")

    # Prepare email components
    to_addresses = "kaitlyn.moss@absolutecaregivers.com; raegan.lopez@absolutecaregivers.com; ulyana.stokolosa@absolutecaregivers.com"
    cc_addresses = "alexander.nazarov@absolutecaregivers.com; luke.kitchel@absolutecaregivers.com"
    subject = "Indianapolis Employee Evaluation Reminder"

    # Convert HTML content to plain text
    email_body = (
        "Dear Kaitlyn,\n\n"
        "I hope this email finds you well. This is an automated reminder regarding the Indianapolis Employee Audit file.\n\n"
        "The following Indianapolis employees require evaluations as indicated in the audit file. "
        "Please follow up with them and update the Indy Employee Audit file accordingly. Thank you for your hard work!\n\n"
    )

    # Add each employee name as a list item
    for emp_name in employees_str.split('\n'):
        email_body += f"- {emp_name}\n"
    email_body += "\nBest regards,\n"

    # Append the signature if available
    signature = get_default_signature()
    if signature:
        # Remove HTML tags from signature
        soup = BeautifulSoup(signature, 'html.parser')
        signature_text = soup.get_text()
        email_body += signature_text
        print("Signature appended to email body.")
    else:
        print("No signature found; using fallback signature.")
        username = os.getlogin()
        email_body += f"{username}\nAbsolute Caregivers"

    # Prepare email addresses
    to_addresses_plain = to_addresses.replace(';', ',')
    cc_addresses_plain = cc_addresses.replace(';', ',')

    # Create the mailto link
    mailto_link = f"mailto:{urllib.parse.quote(to_addresses_plain)}"
    mailto_link += f"?cc={urllib.parse.quote(cc_addresses_plain)}"
    mailto_link += f"&subject={urllib.parse.quote(subject)}"
    mailto_link += f"&body={urllib.parse.quote(email_body)}"

    # Open the mailto link
    webbrowser.open(mailto_link)
    logging.info("Email composed using 'mailto' and opened in default email client.")
    print("Email composed using 'mailto' and opened in default email client.")


def extract_evaluation_expirations():
    """
    Main function to extract employees requiring evaluations and send an email report.
    """
    print("extract_evaluation_expirations called.")
    try:
        username = os.getlogin()
        print(f"Current username: {username}")
    except Exception as e:
        print(f"Failed to get the current username: {e}")
        return

    base_path = os.path.join("C:\\Users", username, "OneDrive - Ability Home Health, LLC")
    print(f"Base path set to: {base_path}")

    # Define the exact filename to search for
    audit_filename = "Employee Audit Checklist.xlsm"
    print(f"Looking for file '{audit_filename}' in 'Documents Audit Files' directories...")

    # Find the file in any 'Documents Audit Files' directory
    audit_file = find_file_in_documents_audit_files(base_path, audit_filename)

    if not audit_file:
        print("Required file not found in specified directories.")
        return

    print(f"Found '{audit_filename}' at: {audit_file}")

    try:
        print("Initializing Excel application...")
        excel = win32.DispatchEx("Excel.Application")  # Use DispatchEx
        excel.DisplayAlerts = False
        excel.Visible = False
        print("Excel application initialized.")

        # Open the Audit Workbook
        print(f"Opening workbook '{audit_filename}'...")
        wb_audit = excel.Workbooks.Open(audit_file, False, True, None, "abs$1004$N")
        ws_active = wb_audit.Sheets("Active")
        print("Employee Audit Checklist workbook opened successfully.")
    except Exception as e:
        print(f"Failed to open '{audit_filename}': {e}")
        if excel:
            excel.Quit()
            print("Excel application closed due to error.")
        return

    # Process the active sheet to find employees requiring evaluations
    print("Processing 'Active' sheet for evaluation expirations...")
    employees_str = process_evaluation_expirations(ws_active)

    # Close the workbook without saving
    try:
        print("Closing workbook and quitting Excel...")
        wb_audit.Close(SaveChanges=False)
        excel.Quit()
        del excel
        print("Workbook and Excel application closed successfully.")
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
    """
    print("run_task called.")
    try:
        extract_evaluation_expirations()
        print("Task completed successfully.")
    except Exception as e:
        print(f"An error occurred during task execution: {e}")
        sys.exit(1)

def main():
    print("Script execution started.")
    run_task()
    print("Script execution finished.")

if __name__ == "__main__":
    main()