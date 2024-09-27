# weekly_tasks/indy_emp_eval.py

import win32com.client as win32 # type:ignore
import os
from datetime import datetime, timedelta

def find_file_in_documents_audit_files(base_path, filename):
    """
    Search for a file in any 'Documents Audit Files' directory under base_path.

    Args:
        base_path (str): The root directory to start searching from.
        filename (str): The exact name of the file to search for.

    Returns:
        str or None: The full path to the found file, or None if not found.
    """
    for root, dirs, files in os.walk(base_path):
        if root.lower().endswith('documents audit files'):
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
    last_col = ws.UsedRange.Columns.Count
    for col in range(1, last_col + 1):
        cell_value = ws.Cells(2, col).Value
        if cell_value and cell_value.strip().lower() == header_name.strip().lower():
            return col
    return None

def process_evaluation_expirations(ws_active):
    """
    Process the 'Active' sheet to find employees who require evaluations (non-blank 'Eval Required' cells).

    Args:
        ws_active: The worksheet object for the 'Active' sheet.

    Returns:
        str: A formatted string listing employee names.
    """
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

    print(f"Last row in 'Active' sheet: {last_row_eval}")

    # Loop through each row starting from row 3
    for i in range(3, last_row_eval + 1):
        emp_name = ws_active.Cells(i, name_col).Value
        eval_value = ws_active.Cells(i, eval_col).Value

        if emp_name and eval_value not in (None, "", "-"):
            # Include employee if 'Eval Required' is not blank
            employees.append(emp_name.strip())

    # Remove duplicates and sort the list
    employees = sorted(set(employees))

    # Concatenate employee names into a single string
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

def send_email(employees_str):
    """
    Compose and send an email via Outlook with the list of employees.

    Args:
        employees_str (str): The formatted string of employees.
    """
    try:
        outlookApp = win32.DispatchEx('Outlook.Application')  # Use DispatchEx
        outlookMail = outlookApp.CreateItem(0)
        outlookMail.To = "alejandra.gamboa@absolutecaregivers.com; kaitlyn.moss@absolutecaregivers.com; raegan.lopez@absolutecaregivers.com; ulyana.stokolosa@absolutecaregivers.com"
        outlookMail.CC = "alexander.nazarov@absolutecaregivers.com; luke.kitchel@absolutecaregivers.com; thea.banks@absolutecaregivers.com"
        outlookMail.Subject = "Indianapolis Employee Evaluation Reminder"

        # Construct the signature path dynamically
        username = os.getlogin()
        signature_filename = "Absolute Signature (seth.riley@absolutecaregivers.com).htm"
        sig_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Signatures', signature_filename)

        # Get the specified signature
        signature = get_signature_by_path(sig_path)

        # Compose the email body in HTML format
        email_body = (
            "<div style='font-family: Calibri, sans-serif; font-size: 11pt;'><p>Dear Kaitlyn,</p>"
            "<p>I hope this email finds you well. This is an automated reminder regarding the Indianapolis Employee Audit file.</p>"
            "<p>The following Indianapolis employees require evaluations as indicated in the audit file. "
            "Please follow up with them and update the Indy Employee Audit file accordingly. Thank you for your hard work!</p> </div>"
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
        else:
            email_body += "<p>{username}<br>Absolute Caregivers</p>"

        # Set the email body and format
        outlookMail.HTMLBody = email_body

        outlookMail.Display()  # Change to .Send() to send automatically without displaying
        print("Email composed successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")

def extract_evaluation_expirations():
    """
    Main function to extract employees requiring evaluations and send an email report.
    """
    try:
        username = os.getlogin()
    except Exception as e:
        print(f"Failed to get the current username: {e}")
        return

    print(f"Current username: {username}")

    base_path = f"C:\\Users\\{username}\\OneDrive - Ability Home Health, LLC\\"

    # Define the exact filename to search for
    audit_filename = "Employee Audit Checklist South Bend.xlsm"

    # Find the file in any 'Documents Audit Files' directory
    audit_file = find_file_in_documents_audit_files(base_path, audit_filename)

    if not audit_file:
        print("Required file not found in specified directories.")
        return

    print(f"Found Employee Audit Checklist.xlsm at: {audit_file}")

    try:
        excel = win32.DispatchEx("Excel.Application")  # Use DispatchEx
        excel.DisplayAlerts = False
        excel.Visible = False

        # Open the Audit Workbook
        # Pass parameters positionally up to Password
        # Excel.Workbooks.Open(Filename, UpdateLinks, ReadOnly, Format, Password, ...)
        wb_audit = excel.Workbooks.Open(audit_file, False, True, None, "abs$1004$N")
        ws_active = wb_audit.Sheets("Active")
        print("Employee Audit Checklist workbook opened successfully.")
    except Exception as e:
        print(f"Failed to open Employee Audit Checklist.xlsm: {e}")
        return

    # Process the active sheet to find employees requiring evaluations
    employees_str = process_evaluation_expirations(ws_active)

    # Close the workbook without saving
    try:
        wb_audit.Close(SaveChanges=False)
        excel.Quit()
        del excel
        print("Workbook closed successfully.")
    except Exception as e:
        print(f"Failed to close workbook: {e}")

    # Send email if there are employees requiring evaluations
    if employees_str.strip():
        send_email(employees_str)
    else:
        print("No employees requiring evaluations found.")

def run_task():
    """
    Wrapper function to execute the extract_evaluation_expirations function.
    Returns the result string or raises an exception.
    """
    try:
        result = extract_evaluation_expirations()
        return result
    except Exception as e:
        raise e

if __name__ == "__main__":
    extract_evaluation_expirations()
