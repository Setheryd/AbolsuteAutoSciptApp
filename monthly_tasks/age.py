import win32com.client as win32
import pandas as pd
import os
import logging
from multiprocessing import Process
import time
import urllib.parse
import webbrowser
from bs4 import BeautifulSoup

# Set up logging to a file
logging.basicConfig(
    filename='extract_eligible_patients.log',
    level=logging.DEBUG,
    format='%(asctime)s %(levelname)s:%(message)s'
)

def find_file(base_path, filename, max_depth=5):
    logging.debug(f"Searching for {filename} in {base_path} up to depth {max_depth}")
    print(f"Searching for {filename} in {base_path} up to depth {max_depth}")
    def scan_directory(path, current_depth):
        if current_depth > max_depth:
            return None
        try:
            with os.scandir(path) as it:
                for entry in it:
                    if entry.is_file() and entry.name.lower() == filename.lower():
                        logging.debug(f"Found file: {entry.path}")
                        print(f"Found file: {entry.path}")
                        return entry.path
                    elif entry.is_dir():
                        found_file = scan_directory(entry.path, current_depth + 1)
                        if found_file:
                            return found_file
        except PermissionError as e:
            logging.warning(f"PermissionError: {e}")
            print(f"PermissionError: {e}")
            return None
    return scan_directory(base_path, 0)

def extract_eligible_patients():
    try:
        username = os.getlogin()
        base_path = f"C:\\Users\\{username}\\OneDrive - Ability Home Health, LLC\\"
        logging.debug(f"Base path: {base_path}")
        print(f"Base path: {base_path}")
        if not os.path.exists(base_path):
            logging.error(f"Base path does not exist: {base_path}")
            print(f"Base path does not exist: {base_path}")
            return None

        files_info = {
            "Absolute Patient Records.xlsm": "Absolute Operation",
            "Absolute Patient Records IHCC.xlsm": "IHCC",
            "Absolute Patient Records PERS.xlsm": "IHCC"
        }

        password = "abs$1018$B"
        patients_within_3_months_of_60 = []

        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.ScreenUpdating = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False
            excel.AlertBeforeOverwriting = False
            logging.debug("Excel application started successfully.")
            print("Excel application started successfully.")
        except Exception as e:
            logging.error(f"Failed to create Excel application: {e}")
            print(f"Failed to create Excel application: {e}")
            return None

        try:
            workbooks = {}
            # Open all workbooks first
            for filename, required_subdir in files_info.items():
                file_path = find_file(base_path, filename)
                if file_path is None:
                    logging.warning(f"File not found: {filename}")
                    print(f"File not found: {filename}")
                    continue

                if required_subdir not in os.path.normpath(file_path):
                    logging.warning(f"File {filename} is not in the required subdirectory: {required_subdir}")
                    print(f"File {filename} is not in the required subdirectory: {required_subdir}")
                    continue

                logging.debug(f"Opening file: {file_path}")
                print(f"Opening file: {file_path}")

                try:
                    wb = excel.Workbooks.Open(
                        file_path,            
                        False,                
                        True,                 
                        None,                 
                        password,             
                        '',                   
                        True,                 
                        None,                 
                        None,                 
                        False,                
                        False,                
                        None,                 
                        False,                
                        False,                
                        0                     
                    )
                    workbooks[filename] = wb
                except Exception as e:
                    logging.error(f"Error opening file {file_path}: {e}")
                    print(f"Error opening file {file_path}: {e}")
                    continue

            # Now process each workbook
            for filename, wb in workbooks.items():
                logging.debug(f"Processing workbook: {filename}")
                print(f"Processing workbook: {filename}")
                try:
                    ws = wb.Sheets("Patient Information")
                    used_range = ws.UsedRange
                    data = used_range.Value

                    if not data:
                        logging.warning(f"No data found in 'Patient Information' sheet in '{filename}'")
                        print(f"No data found in 'Patient Information' sheet in '{filename}'")
                        continue

                    # Get the header row and locate the necessary columns
                    header = data[0]
                    try:
                        age_index = header.index("Age")
                        discharge_date_index = header.index("Discharge Date")  # Index for "Discharge Date" column
                    except ValueError as e:
                        logging.error(f"Required column not found in '{filename}': {e}")
                        print(f"Required column not found in '{filename}': {e}")
                        continue

                    data_rows = data[1:]  # Skip header row

                    # Extract patients within 3 months of turning 60 and who are still active (no discharge date)
                    for row in data_rows:
                        age = row[age_index] if len(row) > age_index else None
                        discharge_date = row[discharge_date_index] if len(row) > discharge_date_index else None

                        if isinstance(age, (int, float)) and 59.75 <= age < 60 and not discharge_date:
                            patient_name = row[2] if len(row) >= 3 else ''  # Assuming column C has the Patient Name
                            patients_within_3_months_of_60.append({
                                "Patient Name": patient_name,
                                "Age": age,
                                "File": filename  # Add the filename to the data
                            })
                            print(f"Found eligible patient: {patient_name}, Age: {age}, File: {filename}")

                    logging.debug(f"Finished processing workbook: {filename}")
                    print(f"Finished processing workbook: {filename}")
                except Exception as e:
                    logging.error(f"Error processing workbook {filename}: {e}")
                    print(f"Error processing workbook {filename}: {e}")
                    continue

        finally:
            # Close all workbooks
            for wb in workbooks.values():
                wb.Close(False)
            excel.Quit()
            del excel
            logging.debug("Excel application closed.")
            print("Excel application closed.")

        if patients_within_3_months_of_60:
            df = pd.DataFrame(patients_within_3_months_of_60)
            logging.info("DataFrame created successfully with patients within 3 months of turning 60.")
            print("DataFrame created successfully with patients within 3 months of turning 60.")
            return df
        else:
            logging.info("No patients within 3 months of turning 60 found.")
            print("No patients within 3 months of turning 60 found.")
            return None
    except Exception as e:
        logging.error(f"Error in extract_eligible_patients: {e}")
        print(f"Error in extract_eligible_patients: {e}")
        return None

def main():
    df = extract_eligible_patients()
    if df is not None:
        logging.info("DataFrame generated successfully.")
        print("DataFrame generated successfully.")
        print(df)
        compose_and_send_email(df)
    else:
        logging.info("No data to display.")
        print("No data to display.")

def get_signature_by_path(sig_path):
    """
    Retrieves the email signature from the specified file path.
    """
    try:
        with open(sig_path, 'r', encoding='utf-8') as file:
            signature = file.read()
        logging.debug(f"Signature retrieved from {sig_path}")
        print(f"Signature retrieved from {sig_path}")
        return signature
    except Exception as e:
        logging.error(f"Unable to retrieve signature from {sig_path}: {e}")
        print(f"Unable to retrieve signature from {sig_path}: {e}")
        return None

def get_default_outlook_email():
    """
    Retrieves the default Outlook email address of the current user.
    """
    try:
        outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        accounts = namespace.Accounts
        if accounts.Count > 0:
            default_account = accounts.Item(1)
            logging.debug(f"Default Outlook email retrieved: {default_account.SmtpAddress}")
            print(f"Default Outlook email retrieved: {default_account.SmtpAddress}")
            return default_account.SmtpAddress
        else:
            logging.error("No Outlook accounts found.")
            print("No Outlook accounts found.")
            return None
    except Exception as e:
        logging.error(f"Unable to retrieve default Outlook email: {e}")
        print(f"Unable to retrieve default Outlook email: {e}")
        return None

def get_default_signature():
    """
    Retrieves the user's default email signature based on their default Outlook account.
    """
    email = get_default_outlook_email()
    if not email:
        logging.error("Default Outlook email not found.")
        print("Default Outlook email not found.")
        return None

    # Define the signature directory
    appdata = os.environ.get('APPDATA')
    if not appdata:
        logging.error("APPDATA environment variable not found.")
        print("APPDATA environment variable not found.")
        return None

    sig_dir = os.path.join(appdata, 'Microsoft', 'Signatures')
    if not os.path.isdir(sig_dir):
        logging.error(f"Signature directory does not exist: {sig_dir}")
        print(f"Signature directory does not exist: {sig_dir}")
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
                    print(f"Signature found: {sig_path}")
                    return signature

    logging.error(f"No signature file found containing email: {email}")
    print(f"No signature file found containing email: {email}")
    return None

def compose_email_classic(email_body, to_addresses, cc_addresses, subject):
    """
    Composes and displays an email via COM automation for classic Outlook.
    """
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.BodyFormat = 2  # HTML format
        mail.HTMLBody = email_body
        mail.To = to_addresses
        mail.CC = cc_addresses
        mail.Display()  # Display email for review
        logging.info("Email created successfully via COM automation.")
        print("Email created successfully via COM automation.")
    except Exception as e:
        logging.error(f"Failed to create or send email via COM automation: {e}")
        print(f"Failed to create or send email via COM automation: {e}")
        raise

def compose_and_send_email(df):
    if df.empty:
        logging.info("No patients within 3 months of turning 60 to include in the email.")
        print("No patients within 3 months of turning 60 to include in the email.")
        return

    # Filter out rows where "Patient Name" is not a string and build the list with bullet points
    patient_list = ""
    for index, row in df.iterrows():
        patient_name = row["Patient Name"]
        if isinstance(patient_name, str):  # Exclude non-text entries
            # Add a bullet point with patient name and corresponding file
            patient_list += f"<li>{patient_name} (from {row['File']})</li>\n"

    if not patient_list:
        logging.info("No valid patient names found.")
        print("No valid patient names found.")
        return

    email_body = f"""
    <div style='font-family: Calibri, sans-serif; font-size: 11pt;'>
    Good Morning Team,<br><br>
    This is a list of Patients from all of our Patient Records Files that are about to turn 60 in the next 3 months. This email is automated, and its intended purpose is to help track MCE enrollment for the patients about to turn 60. Thank you.<br><br>
    <ul>
    {patient_list}
    </ul>
    </div>
    """

    # Get default signature
    signature = get_default_signature()
    if signature:
        email_body += f"<div>{signature}</div>"
        logging.info(f"Signature found and added to email body.")
        print("Signature found and added to email body.")
    else:
        logging.info("No signature found, proceeding without signature.")
        print("No signature found, proceeding without signature.")

    to_addresses = "ulyana.stokolosa@absolutecaregivers.com; victoria.shmoel@absolutecaregivers.com"
    cc_addresses = "alexander.nazarov@absolutecaregivers.com; luke.kitchel@absolutecaregivers.com"
    subject = "List of Patients Within 3 Months of Turning 60"

    # Run compose_email_classic in a separate process with a timeout
    try:
        process = Process(target=compose_email_classic, args=(email_body, to_addresses, cc_addresses, subject))
        process.start()
        process.join(timeout=5)  # Wait up to 5 seconds

        if process.is_alive():
            logging.warning("Composing email via COM automation took too long, terminating process.")
            print("Composing email via COM automation took too long, terminating process.")
            process.terminate()
            process.join()
            raise Exception("Timeout composing email via COM automation.")
        else:
            logging.info("Email composed via COM automation successfully.")
            print("Email composed via COM automation successfully.")
            return  # Exit the function, as the email has been composed
    except Exception as e:
        logging.error(f"Exception during composing email via COM automation: {e}")
        print(f"Exception during composing email via COM automation: {e}")
        # Proceed to fallback method

    # Fallback method using 'mailto' link
    logging.info("Using fallback method to compose email.")
    print("Using fallback method to compose email.")

    # Prepare email components
    body_text = (
        "Good Morning Team,\n\n"
        "This is a list of Patients from all of our Patient Records Files that are about to turn 60 in the next 3 months. "
        "This email is automated, and its intended purpose is to help track MCE enrollment for the patients about to turn 60. Thank you.\n\n"
    )

    # Convert HTML bullet points to plain text
    patient_list_text = ""
    for index, row in df.iterrows():
        patient_name = row["Patient Name"]
        if isinstance(patient_name, str):
            patient_list_text += f"- {patient_name} (from {row['File']})\n"

    body_text += patient_list_text

    # Add signature if available
    signature = get_default_signature()
    if signature:
        # Remove HTML tags from signature
        soup = BeautifulSoup(signature, 'html.parser')
        signature_text = soup.get_text()
        body_text += f"\n\n{signature_text}"
    else:
        body_text += "\n\nBest regards,\n[Your Name]"

    # Prepare email addresses
    to_addresses_plain = to_addresses.replace(';', ',')
    cc_addresses_plain = cc_addresses.replace(';', ',')

    # Create the mailto link
    mailto_link = f"mailto:{urllib.parse.quote(to_addresses_plain)}"
    mailto_link += f"?cc={urllib.parse.quote(cc_addresses_plain)}"
    mailto_link += f"&subject={urllib.parse.quote(subject)}"
    mailto_link += f"&body={urllib.parse.quote(body_text)}"

    # Open the mailto link
    webbrowser.open(mailto_link)
    logging.info("Email composed using 'mailto' and opened in default email client.")
    print("Email composed using 'mailto' and opened in default email client.")

if __name__ == "__main__":
    main()
