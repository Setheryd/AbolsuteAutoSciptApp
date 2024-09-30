# weekly_tasks/sb_pat_sup_exp.py

import win32com.client as win32  # type:ignore
import os
from datetime import datetime, timedelta
import sys

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

def process_evaluation_expirations(workbook, sheet_name="Patient Docs"):
    """
    Process the specified sheet to find employees whose supervisory visits have expired 
    (non-blank 'Sup Visit Required' cells).

    Args:
        workbook: The workbook object containing the sheets.
        sheet_name (str): The name of the sheet to process. Default is "Patient Docs".

    Returns:
        list: A sorted list of unique employee names requiring evaluations.
    """
    employees = []
    
    try:
        # Access the specified sheet
        ws = workbook.Sheets(sheet_name)
    except Exception as e:
        print(f"Sheet '{sheet_name}' not found: {e}")
        return []
    
    # Define headers
    name_header = "Name (Last , First )"
    eval_required_header = "Sup Visit Required"
    
    # Get column indices based on headers
    name_col = get_column_index(ws, name_header)
    eval_col = get_column_index(ws, eval_required_header)
    
    if not name_col or not eval_col:
        print(f"Required headers '{name_header}' or '{eval_required_header}' not found in '{sheet_name}' sheet.")
        return []
    
    # Find the last row in the sheet based on the 'Sup Visit Required' column
    # -4162 corresponds to xlUp
    last_row_eval = ws.Cells(ws.Rows.Count, eval_col).End(-4162).Row
    
    print(f"Last row in '{sheet_name}' sheet: {last_row_eval}")
    
    # Loop through each row starting from row 3
    for i in range(3, last_row_eval + 1):
        emp_name = ws.Cells(i, name_col).Value
        eval_value = ws.Cells(i, eval_col).Value
        
        if emp_name and eval_value not in (None, "", "-"):
            # Include employee if 'Sup Visit Required' is not blank
            employees.append(emp_name.strip())
    
    # Remove duplicates and sort the list
    employees = sorted(set(employees))
    
    print(f"Employees requiring evaluations: {employees}")
    
    return employees

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

def send_email(employees_list, signature, username):
    """
    Compose and send an email via Outlook with the list of employees.

    Args:
        employees_list (list): The list of employee names.
        signature (str): The HTML signature to append to the email.
        username (str): The current user's username for fallback in signature.
    """
    try:
        outlookApp = win32.DispatchEx('Outlook.Application')  # Use DispatchEx for a new instance
        outlookMail = outlookApp.CreateItem(0)
        outlookMail.To = "alejandra.gamboa@absolutecaregivers.com; kaitlyn.moss@absolutecaregivers.com; raegan.lopez@absolutecaregivers.com; ulyana.stokolosa@absolutecaregivers.com"
        outlookMail.CC = "alexander.nazarov@absolutecaregivers.com; luke.kitchel@absolutecaregivers.com; thea.banks@absolutecaregivers.com"
        outlookMail.Subject = "South Bend Patient Supervisory Visit Expiration"

        # Compose the email body in HTML format
        email_body = (
            "<div style='font-family: Calibri, sans-serif; font-size: 11pt;'>"
            "<p>Dear Team,</p>"
            "<p>I hope this email finds you well. This is an automated reminder regarding the South Bend Patient Audit Checklist file.</p>"
            "<p>The following South Bend patients require a supervisory visit as indicated by the South Bend Patient Audit Checklist file. "
            "Please follow up with them and make the necessary changes. Thank you for your hard work!</p>"
            "</div>"
            "<ul>"
        )

        # Add each employee name as a list item
        for emp_name in employees_list:
            email_body += f"<div style='font-family: Calibri, sans-serif; font-size: 11pt;'> <li>{emp_name}</li>"
        email_body += "</ul>"

        email_body += "<div style='font-family: Calibri, sans-serif; font-size: 11pt;'><p>Best regards,</p>"

        # Append the signature if available
        if signature:
            email_body += signature
        else:
            email_body += f"<p>{username}<br>Absolute Caregivers</p>"

        # Set the email body and format
        outlookMail.HTMLBody = email_body

        # Uncomment the next line to send the email automatically
        # outlookMail.Send()
        
        # For testing purposes, display the email
        outlookMail.Display()
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

    base_path = os.path.join("C:\\Users", username, "OneDrive - Ability Home Health, LLC")

    # Define the exact filename to search for
    audit_filename = "Patient Audit Checklist South Bend.xlsm"

    # Find the file in any 'Documents Audit Files' directory
    audit_file = find_file_in_documents_audit_files(base_path, audit_filename)

    if not audit_file:
        print("Required file not found in specified directories.")
        return

    print(f"Found '{audit_filename}' at: {audit_file}")

    excel = None
    wb_audit = None

    try:
        excel = win32.DispatchEx("Excel.Application")  # Use DispatchEx for a new instance
        excel.DisplayAlerts = False
        excel.Visible = False

        # Open the Audit Workbook
        # Parameters: Filename, UpdateLinks, ReadOnly, Format, Password
        wb_audit = excel.Workbooks.Open(Filename=audit_file, UpdateLinks=False, ReadOnly=True, Password="abs$1004$N")
        print("Patient Audit Checklist workbook opened successfully.")
    except Exception as e:
        print(f"Failed to open '{audit_filename}': {e}")
        if excel:
            excel.Quit()
        return

    # Process the "Patient Docs" sheet to find employees requiring evaluations
    employees_list = process_evaluation_expirations(wb_audit, sheet_name="Patient Docs")

    # Close the workbook and quit Excel
    try:
        if wb_audit:
            wb_audit.Close(SaveChanges=False)
            print("Workbook closed successfully.")
        if excel:
            excel.Quit()
            del excel
            print("Excel application closed successfully.")
    except Exception as e:
        print(f"Failed to close Excel properly: {e}")

    # Send email if there are employees requiring evaluations
    if employees_list:
        # Construct the signature path dynamically
        signature_filename = "Absolute Signature (seth.riley@absolutecaregivers.com).htm"
        sig_path = os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'Signatures', signature_filename)
        signature = get_signature_by_path(sig_path)

        send_email(employees_list, signature, username)
    else:
        print("No employees requiring evaluations found.")

def run_task():
    """
    Wrapper function to execute the extract_evaluation_expirations function.
    Returns the result string or raises an exception.
    """
    try:
        extract_evaluation_expirations()
    except Exception as e:
        print(f"An error occurred during task execution: {e}")
        sys.exit(1)

if __name__ == "__main__":
    run_task()
