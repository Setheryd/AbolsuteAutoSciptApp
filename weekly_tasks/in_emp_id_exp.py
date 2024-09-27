# weekly_tasks/in_emp_id_exp.py

import os
from datetime import datetime, timedelta
from pathlib import Path
import win32com.client as win32  # type: ignore
import logging

def find_specific_file(base_path, filename, required_subpath, max_depth=5):
    """
    Search for a specific file within directories that contain a required subpath.

    Args:
        base_path (str or Path): The root directory to start searching from.
        filename (str): The exact name of the file to search for.
        required_subpath (str): A substring that must be present in the file's directory path.
        max_depth (int): Maximum depth to search within subdirectories.

    Returns:
        str or None: The full path to the found file, or None if not found.
    """
    base_path = Path(base_path)  # Ensure base_path is a Path object
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
            print(f"Permission denied: {path}")
            return None
        except Exception as e:
            print(f"Error accessing {path}: {e}")
            return None

        return None

    return scan_directory(base_path, 0)

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

def get_phone_number(employee_name, phone_sheet):
    """
    Retrieve the phone number for the given employee from the phone sheet.

    Args:
        employee_name (str): The name of the employee.
        phone_sheet: The worksheet object containing phone numbers.

    Returns:
        str: The formatted phone number or an appropriate message if not found.
    """
    last_row = phone_sheet.Cells(phone_sheet.Rows.Count, "C").End(-4162).Row  # -4162 corresponds to xlUp
    for i in range(2, last_row + 1):
        name = phone_sheet.Cells(i, "C").Value
        if name and name.strip().lower() == employee_name.strip().lower():
            phone_number = phone_sheet.Cells(i, "E").Value
            return format_phone_number(phone_number) if phone_number else "Phone Number Not Available"
    return "Could not find Phone Number"

def format_phone_number(phone_number):
    """
    Format the phone number to (555) 555-5555.

    Args:
        phone_number (str or int or float): The phone number to format.

    Returns:
        str: The formatted phone number or the original input if formatting fails.
    """
    try:
        phone_str = ''.join(filter(str.isdigit, str(int(phone_number))))
        if len(phone_str) == 10:
            return f"({phone_str[:3]}) {phone_str[3:6]}-{phone_str[6:]}"
        else:
            return phone_number
    except:
        return phone_number

def is_invalid_date(exp_date):
    """
    Check if a date is invalid (not a date, text, or missing).

    Args:
        exp_date: The date value to check.

    Returns:
        bool: True if invalid, False otherwise.
    """
    try:
        if isinstance(exp_date, (int, float)):
            _ = datetime.fromordinal(datetime(1899, 12, 30).toordinal() + int(exp_date))
            return False
        elif isinstance(exp_date, datetime):
            return False
        else:
            return True
    except:
        return True

def process_employee_audit(ws_active, phone_sheet):
    """
    Process the 'Active' sheet to find employees with invalid or expiring data.

    Args:
        ws_active: The worksheet object for the 'Active' sheet.
        phone_sheet: The worksheet object containing phone numbers.

    Returns:
        str: A formatted string listing expiring or invalid employees.
    """
    expiring_employees = []
    today = datetime.today().date()
    two_weeks = today + timedelta(days=14)

    name_header = "Last Name, First Name"
    id_exp_header = "ID /Exp date"

    name_col = get_column_index(ws_active, name_header)
    id_exp_col = get_column_index(ws_active, id_exp_header)

    if not name_col or not id_exp_col:
        print("Required headers not found in 'Active' sheet.")
        return ""

    last_row = ws_active.Cells(ws_active.Rows.Count, name_col).End(-4162).Row  # -4162 corresponds to xlUp
    print(f"Last row in 'Active' sheet: {last_row}")

    for i in range(3, last_row + 1):
        emp_name = ws_active.Cells(i, name_col).Value
        exp_date = ws_active.Cells(i, id_exp_col).Value

        if emp_name:
            if is_invalid_date(exp_date) or not exp_date:
                # Invalid date, N/A, or blank entry
                phone_number = get_phone_number(emp_name, phone_sheet)
                expiring_employees.append({
                    "Employee Name": emp_name,
                    "Expiration Date": "N/A or Invalid (Please add a valid expiration date)",
                    "Phone Number": phone_number
                })
            else:
                # Check if date is valid and process further
                if isinstance(exp_date, (int, float)):
                    exp_date_py = datetime.fromordinal(datetime(1899, 12, 30).toordinal() + int(exp_date)).date()
                elif isinstance(exp_date, datetime):
                    exp_date_py = exp_date.date()
                else:
                    exp_date_py = None

                if exp_date_py:
                    if exp_date_py <= today:
                        # Already expired
                        phone_number = get_phone_number(emp_name, phone_sheet)
                        expiring_employees.append({
                            "Employee Name": emp_name,
                            "Expiration Date": f"Expired on {exp_date_py.strftime('%m/%d/%Y')}",
                            "Phone Number": phone_number
                        })
                    elif today <= exp_date_py <= two_weeks:
                        # Expiring soon
                        phone_number = get_phone_number(emp_name, phone_sheet)
                        expiring_employees.append({
                            "Employee Name": emp_name,
                            "Expiration Date": f"Expires on {exp_date_py.strftime('%m/%d/%Y')}",
                            "Phone Number": phone_number
                        })

    # Format the list of expiring employees into a string
    expiring_employees_str = ""
    for emp in expiring_employees:
        expiring_employees_str += (
            f"{emp['Employee Name']}\n"
            f"Expiration Date: {emp['Expiration Date']}\n"
            f"Phone Number: {emp['Phone Number']}\n\n"
        )

    return expiring_employees_str

def get_signature_by_path(sig_path):
    """
    Retrieves the email signature from the specified file path.

    Args:
        sig_path (str): The full path to the signature file.

    Returns:
        str: The signature HTML content if available, otherwise None.
    """
    try:
        with open(sig_path, 'r', encoding='utf-8') as file:
            signature = file.read()
        return signature
    except Exception as e:
        logging.error(f"Unable to retrieve signature from {sig_path}: {e}")
        return None

def get_default_outlook_email():
    """
    Retrieves the default Outlook email address of the current user.

    Returns:
        str: The default email address if available, otherwise None.
    """
    try:
        outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        accounts = namespace.Accounts
        if accounts.Count > 0:
            # Outlook accounts are 1-indexed
            default_account = accounts.Item(1)
            return default_account.SmtpAddress
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

def send_email(expiring_employees_str):
    """
    Compose and send an email via Outlook with the list of expiring employees.

    Args:
        expiring_employees_str (str): The formatted string of expiring employees.
    """
    try:
        # Initialize Outlook application object using DispatchEx for better performance
        outlookApp = win32.DispatchEx('Outlook.Application')  # Changed from Dispatch to DispatchEx
        mail = outlookApp.CreateItem(0)  # 0: olMailItem

        # Define recipients
        mail.To = (
            "alejandra.gamboa@absolutecaregivers.com; "
            "kaitlyn.moss@absolutecaregivers.com; "
            "raegan.lopez@absolutecaregivers.com; "
            "ulyana.stokolosa@absolutecaregivers.com"
        )
        mail.CC = (
            "alexander.nazarov@absolutecaregivers.com; "
            "luke.kitchel@absolutecaregivers.com; "
            "thea.banks@absolutecaregivers.com"
        )
        mail.Subject = "Weekly Update: Expired or Expiring Drivers Licenses"

        # Get the default signature
        signature = get_default_signature()

        # Compose the email body in HTML format with consistent font and size
        email_body = (
            "<div style='font-family: Calibri, sans-serif; font-size: 11pt;'>"
            "<p>Hi Kaitlyn,</p>"
            "<p>I hope this message finds you well.</p>"
            "<p>This is your weekly update with the list of employees who either have expired or are close to expiring Drivers Licenses. "
            "Please contact them. Once resolved, update the employee audit checklist with their new expirations.</p>"
            "<pre style='font-family: Calibri, sans-serif; font-size: 11pt;'>"
            f"{expiring_employees_str}"
            "</pre>"
            "<p>Best regards,</p>"
            "</div>"
        )

        # Append the signature if available
        if signature:
            email_body += signature
        else:
            # Fallback signature if the specific signature file is not found
            email_body += "<p>Your Name<br>Absolute Caregivers</p>"

        # Set the email body and format
        mail.HTMLBody = email_body

        # Display the email for manual review before sending
        mail.Display(False)  # False to open the email without a modal dialog
        logging.info("Email composed and displayed successfully.")
        print("Email composed successfully.")

    except Exception as e:
        logging.error(f"Failed to compose or display email: {e}")
        print(f"Failed to compose email: {e}")

    finally:
        # Release COM objects to free up resources
        try:
            if 'mail' in locals() and mail:
                del mail
            if 'outlookApp' in locals() and outlookApp:
                del outlookApp
        except Exception as cleanup_error:
            logging.error(f"Error during cleanup: {cleanup_error}")
            print(f"Error during cleanup: {cleanup_error}")

def extract_expiring_employees():
    """
    Main function to extract expiring employees and send an email report.
    """
    try:
        username = os.getlogin()
    except Exception as e:
        print(f"Failed to get the current username: {e}")
        return

    print(f"Current username: {username}")

    base_path = Path(f"C:/Users/{username}/OneDrive - Ability Home Health, LLC/")
    print(f"Base path: {base_path}")

    # Check if base path exists
    if not base_path.exists():
        print(f"Base path {base_path} does not exist.")
        return

    # Define the exact filenames to search for
    audit_filename = "Employee Audit Checklist.xlsm"
    demographics_filename = "Absolute Employee Demographics.xlsm"  # Ensure this matches the actual filename

    # Define the required subpaths for each file
    audit_required_subpath = "Documents Audit Files"
    demographics_required_subpath = "Employee Demographics File"

    # Find the files using find_specific_file
    audit_file = find_specific_file(base_path, audit_filename, audit_required_subpath)
    demographics_file = find_specific_file(base_path, demographics_filename, demographics_required_subpath)

    if not audit_file or not demographics_file:
        print("Required files not found in specified directories.")
        return

    print(f"Found {audit_filename} at: {audit_file}")
    print(f"Found {demographics_filename} at: {demographics_file}")

    try:
        # Replace Dispatch with DispatchEx to create a new Excel instance
        excel = win32.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False

        # Open the Audit Workbook
        # Modified to pass parameters positionally up to Password
        wb_audit = excel.Workbooks.Open(audit_file, False, True, None, "abs$1004$N")
        ws_active = wb_audit.Sheets("Active")
        print("Employee Audit Checklist workbook opened successfully.")
    except Exception as e:
        print(f"Failed to open {audit_filename}: {e}")
        return

    try:
        # Open the Demographics Workbook
        # Modified to pass parameters positionally up to Password
        wb_demographics = excel.Workbooks.Open(demographics_file, False, True, None, "abs$1004$N")
        phone_sheet = wb_demographics.Sheets("Contractor_Employee")
        print("Absolute Employee Demographics workbook opened successfully.")
    except Exception as e:
        print(f"Failed to open {demographics_filename}: {e}")
        try:
            wb_audit.Close(SaveChanges=False)
        except:
            pass
        excel.Quit()
        del excel
        return

    # Process the active sheet to find expiring employees
    expiring_employees_str = process_employee_audit(ws_active, phone_sheet)

    # Close the workbooks without saving
    try:
        wb_demographics.Close(SaveChanges=False)
        wb_audit.Close(SaveChanges=False)
        excel.Quit()
        del excel
        print("All workbooks closed successfully.")
    except Exception as e:
        print(f"Failed to close workbooks: {e}")

    # Send email if there are expiring employees
    if expiring_employees_str.strip():
        send_email(expiring_employees_str)
    else:
        print("No expiring employees found.")

def run_task():
    """
    Wrapper function to execute the extract_expiring_employees function.
    Returns the result string or raises an exception.
    """
    try:
        result = extract_expiring_employees()
        return result
    except Exception as e:
        raise e

if __name__ == "__main__":
    # Configure logging
    logging.basicConfig(
        filename='in_emp_id_exp.log',
        filemode='a',
        format='%(asctime)s - %(levelname)s - %(message)s',
        level=logging.INFO
    )
    extract_expiring_employees()
