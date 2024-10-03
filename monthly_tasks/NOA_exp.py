import win32com.client as win32
import os
import logging
import sys

# Set up logging to both a file and the console
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# File handler for logging to a file
file_handler = logging.FileHandler("extract_eligible_patients.log")
file_handler.setLevel(logging.DEBUG)
file_formatter = logging.Formatter("%(asctime)s %(levelname)s:%(message)s")
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)

# Console handler for logging to the console
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO)  # Set to DEBUG for more detailed console output
console_formatter = logging.Formatter("%(asctime)s %(levelname)s:%(message)s")
console_handler.setFormatter(console_formatter)
logger.addHandler(console_handler)

# Define the xlUp constant directly
XL_UP = -4162


def find_file(base_path, filename, max_depth=5):
    """
    Optimized search for a file in a directory up to a specified depth using os.scandir for speed.
    """
    logging.debug(
        f"Searching for '{filename}' in '{base_path}' up to depth {max_depth}"
    )

    def scan_directory(path, current_depth):
        if current_depth > max_depth:
            logging.debug(f"Maximum search depth reached at '{path}'")
            return None
        try:
            with os.scandir(path) as it:
                for entry in it:
                    if entry.is_file() and entry.name.lower() == filename.lower():
                        logging.info(f"Found file: {entry.path}")
                        return entry.path
                    elif entry.is_dir():
                        found_file = scan_directory(entry.path, current_depth + 1)
                        if found_file:
                            return found_file
        except PermissionError as e:
            logging.warning(f"PermissionError accessing '{path}': {e}")
            return None

    return scan_directory(base_path, 0)


def extract_eligible_patients():
    try:
        username = os.getlogin()
        base_path = f"C:\\Users\\{username}\\OneDrive - Ability Home Health, LLC\\"
        logging.debug(f"Base path set to: {base_path}")

        if not os.path.exists(base_path):
            logging.error(f"Base path does not exist: {base_path}")
            return None

        files_info = {
            "MDC ATTC&HMK 2023-2024.xlsx": "Absolute Billing and Payroll",
            "Anthem ATTC&HMK 2023-2024.xlsx": "Absolute Billing and Payroll",
            "Humana ATTC&HMK 2023-2024.xlsx": "Absolute Billing and Payroll",
            "United ATTC&HMK 2023-2024.xlsx": "Absolute Billing and Payroll",
            "Units Record CHOICE Indianapolis 2023-2024.xlsx": "Absolute Billing and Payroll",
            "Units Record CHOICE South Bend 2023-2024.xlsx": "Absolute Billing and Payroll",
            "Units Record SFC 2023-2024.xlsx": "Absolute Billing and Payroll",
            "Units Record IHCC 2023-2024.xlsx": "Absolute Billing and Payroll",
            "Units Record PERS 2023-2024.xlsx": "Absolute Billing and Payroll",
            "Units Record NS 2023-2024.xlsx": "Absolute Billing and Payroll",
        }

        # It's recommended to use environment variables for sensitive information
        password = os.getenv(
            "EXCEL_PASSWORD", "abs$0321$S"
        )  # Replace with secure method in production
        collected_data = []

        try:
            # Use DispatchEx to create a new Excel instance
            excel = win32.DispatchEx("Excel.Application")
            # Set Excel application properties to ensure it's hidden
            excel.Visible = False
            excel.ScreenUpdating = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False
            excel.AlertBeforeOverwriting = False
            logging.debug("Excel application started successfully.")
        except Exception as e:
            logging.error(f"Failed to create Excel application: {e}")
            return None

        try:
            workbooks = {}
            # Open all workbooks first
            for filename, required_subdir in files_info.items():
                file_path = find_file(base_path, filename)
                if file_path is None:
                    logging.warning(f"File not found: {filename}")
                    continue

                # Verify the file is in the correct subdirectory
                normalized_path = os.path.normpath(file_path)
                if required_subdir.lower() not in normalized_path.lower():
                    logging.warning(
                        f"File '{filename}' is not in the required subdirectory '{required_subdir}'. Found path: '{file_path}'"
                    )
                    continue

                logging.info(f"Opening file: {file_path}")

                try:
                    wb = excel.Workbooks.Open(
                        file_path,  # Filename
                        False,  # UpdateLinks
                        True,  # ReadOnly
                        None,  # Format
                        password,  # Password
                        "",  # WriteResPassword
                        True,  # IgnoreReadOnlyRecommended
                        None,  # Origin
                        None,  # Delimiter
                        False,  # Editable
                        False,  # Notify
                        None,  # Converter
                        False,  # AddToMru
                        False,  # Local
                        0,  # CorruptLoad
                    )
                    workbooks[filename] = wb
                    logging.debug(f"Workbook '{filename}' opened successfully.")
                except Exception as e:
                    logging.error(f"Error opening file '{file_path}': {e}")
                    continue

            # Now process each workbook
            for filename, wb in workbooks.items():
                logging.info(f"Processing workbook: {filename}")
                try:
                    ws = wb.Sheets("Expired NOAs")
                    # Get the last used row in column C and G
                    last_row_c = ws.Cells(ws.Rows.Count, "C").End(XL_UP).Row
                    last_row_g = ws.Cells(ws.Rows.Count, "G").End(XL_UP).Row
                    last_row = max(last_row_c, last_row_g)
                    logging.debug(
                        f"Last row in column C for '{filename}': {last_row_c}"
                    )
                    logging.debug(
                        f"Last row in column G for '{filename}': {last_row_g}"
                    )
                    logging.debug(f"Overall last row for '{filename}': {last_row}")

                    # First, process all entries from column C
                    first_entry_c = True
                    for row in range(2, last_row_c + 1):
                        cell_c = ws.Cells(row, "C")
                        cell_value_c = cell_c.Value
                        if cell_value_c and isinstance(cell_value_c, str):
                            name_c = cell_value_c.strip()
                            if name_c:  # Ensure the name is not empty after stripping
                                if first_entry_c:
                                    # It's the first entry from column C, mark as category
                                    collected_data.append(f"CATEGORY: {name_c}")
                                    first_entry_c = False
                                    logging.debug(
                                        f"Marked as CATEGORY from column C, row {row}: {name_c}"
                                    )
                                else:
                                    # Subsequent entries from column C, mark as item
                                    collected_data.append(f"ITEM: {name_c}")
                                    logging.debug(
                                        f"Marked as ITEM from column C, row {row}: {name_c}"
                                    )

                    # Then, process all entries from column G
                    first_entry_g = True
                    for row in range(2, last_row_g + 1):
                        cell_g = ws.Cells(row, "G")
                        cell_value_g = cell_g.Value
                        if cell_value_g and isinstance(cell_value_g, str):
                            name_g = cell_value_g.strip()
                            if name_g:  # Ensure the name is not empty after stripping
                                if first_entry_g:
                                    # It's the first entry from column G, mark as category
                                    collected_data.append(f"CATEGORY: {name_g}")
                                    first_entry_g = False
                                    logging.debug(
                                        f"Marked as CATEGORY from column G, row {row}: {name_g}"
                                    )
                                else:
                                    # Subsequent entries from column G, mark as item
                                    collected_data.append(f"ITEM: {name_g}")
                                    logging.debug(
                                        f"Marked as ITEM from column G, row {row}: {name_g}"
                                    )

                    logging.info(f"Processed workbook '{filename}' successfully.")

                except Exception as e:
                    logging.error(f"Error processing workbook '{filename}': {e}")
                    continue

        finally:
            # Close all workbooks
            for wb in workbooks.values():
                wb.Close(False)
            excel.Quit()
            del excel
            logging.debug("Excel application closed.")

        if collected_data:
            logging.info(f"Collected {len(collected_data)} entries.")

            # Optionally, log a sample of the collected names
            sample_size = min(10, len(collected_data))
            sample_names = collected_data[:sample_size]
            logging.debug(f"Sample of collected entries: {sample_names}")

            # Convert the list of entries to a formatted string with markers
            expiring_employees_str = "\n".join(entry for entry in collected_data)
            return expiring_employees_str
        else:
            logging.info("No data collected.")
            return None
    except Exception as e:
        logging.error(f"Error in extract_eligible_patients: {e}")
        return None


def get_signature_by_path(sig_path):
    """
    Retrieves the email signature from the specified file path.

    Args:
        sig_path (str): The full path to the signature file.

    Returns:
        str: The signature HTML content if available, otherwise None.
    """
    try:
        with open(sig_path, "r", encoding="utf-8") as file:
            signature = file.read()
        return signature
    except Exception as e:
        logging.error(f"Unable to retrieve signature from '{sig_path}': {e}")
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
            email_address = default_account.SmtpAddress
            logging.debug(f"Default Outlook email address: {email_address}")
            return email_address
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
    appdata = os.environ.get("APPDATA")
    if not appdata:
        logging.error("APPDATA environment variable not found.")
        return None

    sig_dir = os.path.join(appdata, "Microsoft", "Signatures")
    if not os.path.isdir(sig_dir):
        logging.error(f"Signature directory does not exist: '{sig_dir}'")
        return None

    # Iterate through signature files to find a match
    for filename in os.listdir(sig_dir):
        if filename.lower().endswith((".htm", ".html")):
            # Extract the base name without extension
            base_name = os.path.splitext(filename)[0].lower()
            if email.lower() in base_name:
                sig_path = os.path.join(sig_dir, filename)
                signature = get_signature_by_path(sig_path)
                if signature:
                    logging.info(f"Signature found: '{sig_path}'")
                    return signature

    logging.error(f"No signature file found containing email: '{email}'")
    return None


def send_email(expiring_employees_str):
    """
    Compose and send an email via Outlook with the list of expiring employees.

    Args:
        expiring_employees_str (str): The formatted string of expiring employees with markers.
    """
    try:
        # Initialize Outlook application object using DispatchEx for better performance
        outlookApp = win32.DispatchEx(
            "Outlook.Application"
        )  # Changed from Dispatch to DispatchEx
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
        mail.Subject = "Weekly Update: Expired or Expiring Patients' Authorizations"

        # Get the default signature
        signature = get_default_signature()

        # Start composing the email body with introductory text
        email_body = (
            "<div style='font-family: Calibri, sans-serif; font-size: 11pt;'>"
            "<p>Hello Team,</p>"
            "<p>This is an automated email generated using the billing files to track and maintain records of patients who may not have authorized units. "
            "The accuracy of the names and details in this list depends on the maintenance and precision of the billing files themselves. "
            "Please review the attached information carefully and ensure that any discrepancies are addressed promptly.</p>"
            "<p>Thank you for your attention to this matter.</p>"
        )

        # Initialize a flag to track if a <ul> is open
        ul_open = False

        # Process the collected data to structure categories and items
        lines = expiring_employees_str.splitlines()

        for line in lines:
            line = line.strip()
            if line.startswith("CATEGORY: "):
                # Close any open <ul> before starting a new category
                if ul_open:
                    email_body += "</ul>"
                    ul_open = False
                # Extract the category name
                category = line.replace("CATEGORY: ", "").strip()
                email_body += f"<p><strong>{category}</strong></p>"
            elif line.startswith("ITEM: "):
                # Start a new <ul> if not already open
                if not ul_open:
                    email_body += "<ul style='margin-left: 20px;'>"
                    ul_open = True
                # Extract the item name
                item = line.replace("ITEM: ", "").strip()
                email_body += f"<li>{item}</li>"
            else:
                # Handle unexpected formats
                logging.warning(f"Unexpected line format: {line}")
                if not ul_open:
                    email_body += "<ul style='margin-left: 20px;'>"
                    ul_open = True
                email_body += f"<li>{line}</li>"

        # Close any remaining <ul>
        if ul_open:
            email_body += "</ul>"

        # Add closing remarks
        email_body += "<p>Best regards,</p>"
        email_body += "</div>"

        # Append the signature if available
        if signature:
            email_body += signature
            logging.debug("Appended the default signature to the email.")
        else:
            # Fallback signature if the specific signature file is not found
            email_body += "<p>Your Name<br>Absolute Caregivers</p>"
            logging.debug("Appended the fallback signature to the email.")

        # Set the email body and format
        mail.HTMLBody = email_body
        logging.debug("Email body composed successfully.")

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
            if "mail" in locals() and mail:
                del mail
            if "outlookApp" in locals() and outlookApp:
                del outlookApp
            logging.debug("Released COM objects for Outlook.")
        except Exception as cleanup_error:
            logging.error(f"Error during cleanup: {cleanup_error}")
            print(f"Error during cleanup: {cleanup_error}")


def main():
    logging.info("Script started.")
    expiring_employees_str = extract_eligible_patients()
    if expiring_employees_str:
        logging.info("Data collected successfully. Preparing to send email.")
        send_email(expiring_employees_str)
    else:
        logging.info("No expiring patient data to send.")
        print("No expiring patient data to send.")
    logging.info("Script finished.")


if __name__ == "__main__":
    main()
