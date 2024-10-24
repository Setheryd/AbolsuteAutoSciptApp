# weekly_tasks/in_emp_id_exp.py

import os
from datetime import datetime, timedelta
from pathlib import Path
import win32com.client as win32  # type: ignore
import logging
from multiprocessing import Process
import urllib.parse
import webbrowser
from bs4 import BeautifulSoup  # For parsing HTML signatures
print("Starting script execution.")

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
    print(f"find_specific_file called with base_path: {base_path}, filename: {filename}, required_subpath: {required_subpath}")

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
                            return entry.path
                    elif entry.is_dir():
                        found_file = scan_directory(entry.path, current_depth + 1)
                        if found_file:
                            return found_file
        except PermissionError:
            # Skip directories that are not accessible
            return None
        except Exception as e:
            return None

        return None

    result = scan_directory(base_path, 0)
    if result:
        print(f"File found: {result}")
    else:
        print(f"File '{filename}' not found within the specified parameters.")
    return result

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
    for col in range(1, last_col + 1):
        cell_value = ws.Cells(2, col).Value
        if cell_value and cell_value.strip().lower() == header_name.strip().lower():
            print(f"Header '{header_name}' found at column {col}")
            return col
    print(f"Header '{header_name}' not found.")
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
            formatted_number = format_phone_number(phone_number) if phone_number else "Phone Number Not Available"
            return formatted_number
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
            formatted = f"({phone_str[:3]}) {phone_str[3:6]}-{phone_str[6:]}"
            return formatted
        else:
            return str(phone_number)
    except Exception as e:
        return str(phone_number)

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
    except Exception as e:
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
    print("process_employee_audit called.")
    expiring_employees = []
    today = datetime.today().date()
    two_weeks = today + timedelta(days=14)
    print(f"Today's date: {today}, Two weeks from today: {two_weeks}")

    name_header = "Last Name, First Name"
    id_exp_header = "ID /Exp date"

    name_col = get_column_index(ws_active, name_header)
    id_exp_col = get_column_index(ws_active, id_exp_header)

    if not name_col or not id_exp_col:
        print("Required headers not found in 'Active' sheet.")
        return ""

    last_row = ws_active.Cells(ws_active.Rows.Count, name_col).End(-4162).Row  # -4162 corresponds to xlUp

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
                else:
                    pass  # Unable to convert date; skip processing
        else:
            pass  # No employee name found; skip processing

    # Format the list of expiring employees into a string
    expiring_employees_str = ""
    print(f"Total expiring employees found: {len(expiring_employees)}")
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
    print("get_default_outlook_email called.")
    try:
        outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        accounts = namespace.Accounts
        if accounts.Count > 0:
            # Outlook accounts are 1-indexed
            default_account = accounts.Item(1)
            email = default_account.SmtpAddress
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

def compose_email_classic(expiring_employees_str):
    """
    Composes and displays an email via COM automation for classic Outlook.
    """
    print("compose_email_classic called.")
    try:
        # Initialize Outlook application object
        print("Initializing Outlook application...")
        outlookApp = win32.Dispatch("Outlook.Application")
        mail = outlookApp.CreateItem(0)  # 0: olMailItem

        # Define recipients
        mail.To = (
            "kaitlyn.moss@absolutecaregivers.com; "
            "raegan.lopez@absolutecaregivers.com; "
            "ulyana.stokolosa@absolutecaregivers.com"
        )
        mail.CC = (
            "alexander.nazarov@absolutecaregivers.com; "
            "luke.kitchel@absolutecaregivers.com; "
        )
        mail.Subject = "Weekly Update: Expired or Expiring Drivers Licenses"
        print("Email recipients and subject set.")

        # Get the default signature
        print("Retrieving default signature...")
        signature = get_default_signature()

        # Compose the email body in HTML format with consistent font and size
        print("Composing email body...")
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
            print("Signature appended to email body.")
        else:
            # Fallback signature if the specific signature file is not found
            email_body += "<p>Your Name<br>Absolute Caregivers</p>"
            print("Default fallback signature used.")

        # Set the email body and format
        mail.HTMLBody = email_body
        print("Email body set.")

        # Display the email for manual review before sending
        print("Displaying email for review...")
        mail.Display(False)  # False to open the email without a modal dialog
        logging.info("Email composed and displayed successfully.")
        print("Email composed and displayed successfully.")

    except Exception as e:
        logging.error(f"Failed to compose or display email via COM automation: {e}")
        raise

def send_email(expiring_employees_str):
    """
    Composes and sends an email with a 5-second timeout for COM automation.
    Falls back to using a mailto link if the timeout is exceeded.
    """
    print("send_email called.")
    # Try composing email via COM automation with a timeout
    try:
        process = Process(target=compose_email_classic, args=(expiring_employees_str,))
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
    to_addresses = (
        "kaitlyn.moss@absolutecaregivers.com; "
        "raegan.lopez@absolutecaregivers.com; "
        "ulyana.stokolosa@absolutecaregivers.com"
    )
    cc_addresses = (
        "alexander.nazarov@absolutecaregivers.com; "
        "luke.kitchel@absolutecaregivers.com; "
    )
    subject = "Weekly Update: Expired or Expiring Drivers Licenses"

    # Convert HTML content to plain text
    email_body = (
        "Hi Kaitlyn,\n\n"
        "I hope this message finds you well.\n\n"
        "This is your weekly update with the list of employees who either have expired or are close to expiring Drivers Licenses. "
        "Please contact them. Once resolved, update the employee audit checklist with their new expirations.\n\n"
        f"{expiring_employees_str}\n\n"
        "Best regards,\n"
    )

    # Append the signature if available
    signature = get_default_signature()
    if signature:
        # Remove HTML tags from signature
        soup = BeautifulSoup(signature, 'html.parser')
        signature_text = soup.get_text()
        email_body += signature_text
        print("Signature appended to email body.")
    else:
        # Fallback signature
        email_body += "Your Name\nAbsolute Caregivers"
        print("Default fallback signature used.")

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
    
def extract_expiring_employees():
    """
    Main function to extract expiring employees and send an email report.
    """
    print("extract_expiring_employees called.")
    try:
        username = os.getlogin()
        print(f"Current username: {username}")
    except Exception as e:
        print(f"Failed to get the current username: {e}")
        return

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
    print("Searching for audit file...")
    audit_file = find_specific_file(base_path, audit_filename, audit_required_subpath)
    print("Searching for demographics file...")
    demographics_file = find_specific_file(base_path, demographics_filename, demographics_required_subpath)

    if not audit_file or not demographics_file:
        print("Required files not found in specified directories.")
        return

    print(f"Found {audit_filename} at: {audit_file}")
    print(f"Found {demographics_filename} at: {demographics_file}")

    try:
        # Replace Dispatch with Dispatch to create a new Excel instance
        print("Initializing Excel application...")
        excel = win32.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        print("Excel application initialized.")

        # Open the Audit Workbook
        # Modified to pass parameters positionally up to Password
        print(f"Opening workbook: {audit_file}")
        wb_audit = excel.Workbooks.Open(audit_file, False, True, None, "abs$1004$N")
        ws_active = wb_audit.Sheets("Active")
        print("Employee Audit Checklist workbook opened successfully.")
    except Exception as e:
        print(f"Failed to open {audit_filename}: {e}")
        return

    try:
        # Open the Demographics Workbook
        # Modified to pass parameters positionally up to Password
        print(f"Opening workbook: {demographics_file}")
        wb_demographics = excel.Workbooks.Open(demographics_file, False, True, None, "abs$1004$N")
        phone_sheet = wb_demographics.Sheets("Contractor_Employee")
        print("Absolute Employee Demographics workbook opened successfully.")
    except Exception as e:
        print(f"Failed to open {demographics_filename}: {e}")
        try:
            wb_audit.Close(SaveChanges=False)
            print("Audit workbook closed after failure.")
        except:
            pass
        excel.Quit()
        del excel
        print("Excel application closed after failure.")
        return

    # Process the active sheet to find expiring employees
    print("Processing employee audit...")
    expiring_employees_str = process_employee_audit(ws_active, phone_sheet)

    # Close the workbooks without saving
    try:
        print("Closing workbooks...")
        wb_demographics.Close(SaveChanges=False)
        wb_audit.Close(SaveChanges=False)
        excel.Quit()
        del excel
        print("All workbooks closed successfully.")
    except Exception as e:
        print(f"Failed to close workbooks: {e}")

    # Send email if there are expiring employees
    if expiring_employees_str.strip():
        print("Expiring employees found, sending email...")
        send_email(expiring_employees_str)
    else:
        print("No expiring employees found.")

def run_task():
    """
    Wrapper function to execute the extract_expiring_employees function.
    Returns the result string or raises an exception.
    """
    print("run_task called.")
    try:
        result = extract_expiring_employees()
        print("Task completed successfully.")
        return result
    except Exception as e:
        print(f"Error running task: {e}")
        raise e

def main():
    # Configure logging
    print("Script started.")
    logging.basicConfig(
        filename='in_emp_id_exp.log',
        filemode='a',
        format='%(asctime)s - %(levelname)s - %(message)s',
        level=logging.INFO
    )
    print("Logging configured.")
    extract_expiring_employees()
    print("Script finished.")

if __name__ == "__main__":
    main()