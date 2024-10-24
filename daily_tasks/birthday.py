import os
import pandas as pd
import datetime
import random
import win32com.client as win32
import tempfile
import sys
import logging
from bs4 import BeautifulSoup
from multiprocessing import Process
import time


def get_current_username():
    """
    Retrieves the current logged-in username.
    """
    try:
        return os.getlogin()
    except Exception as e:
        logging.error(f"Failed to get the current username: {e}")
        return None


def get_resource_path(relative_path):
    """Get the absolute path to the resource, works for PyInstaller executable."""
    try:
        # PyInstaller creates a temporary folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # If not running as an executable, use the current script directory
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)


def find_employee_demographics_file(base_path):
    """
    Recursively searches for 'Absolute Employee Demographics.xlsm' within
    any subdirectory under the base path that contains 'Employee Demographics File'.
    """
    target_folder = "Employee Demographics File"
    target_file = "Absolute Employee Demographics.xlsm"
    
    try:
        with os.scandir(base_path) as it:
            for entry in it:
                if entry.is_dir():
                    # Check if the current directory is the target folder
                    if entry.name == target_folder:
                        demographics_path = os.path.join(entry.path, target_file)
                        if os.path.exists(demographics_path):
                            return demographics_path
                    else:
                        # Recursively search in subdirectories
                        result = find_employee_demographics_file(entry.path)
                        if result:
                            return result
    except Exception as e:
        logging.error(f"Error while searching for the demographics file: {e}")
    
    return None

def read_password_protected_excel(excel_path, password, sheet_name):
    """
    Opens a password-protected Excel file, saves a temporary unprotected copy,
    reads it with pandas, and then deletes the temporary file.
    
    Parameters:
        excel_path (str): Path to the password-protected Excel file.
        password (str): Password for the Excel file.
        sheet_name (str): Name of the sheet to read.
        
    Returns:
        pd.DataFrame or None: DataFrame containing the sheet data, or None if failed.
    """
    excel_app = None
    wb = None
    try:
        # Initialize Excel application
        excel_app = win32.Dispatch("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

        # Open the workbook with password using positional arguments
        # Parameters based on VBA's Workbooks.Open method:
        # Filename, UpdateLinks, ReadOnly, Format, Password
        wb = excel_app.Workbooks.Open(
            excel_path,      # Filename
            False,           # UpdateLinks: 0 = don't update
            False,           # ReadOnly: False = open as read/write
            None,            # Format: Not specifying
            password         # Password
        )

        # Save it to a temporary file without password
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, f"temp_{os.path.basename(excel_path)}")
        wb.SaveAs(
            Filename=temp_file,
            FileFormat=wb.FileFormat,
            Password="",                # Remove password
            WriteResPassword="",        # Remove write-reservation password
            ReadOnlyRecommended=False,
            CreateBackup=False
        )

        wb.Close(False)
        excel_app.Quit()

        # Read the temp file with pandas, specifying header=1 for row 2
        df = pd.read_excel(temp_file, sheet_name=sheet_name, engine='openpyxl', header=1)

        # Log the columns and sample data for debugging
        logging.info(f"Temporary file created at: {temp_file}")
        logging.info(f"Columns found in the DataFrame: {df.columns.tolist()}")
        logging.info("Sample data from the DataFrame:")
        logging.info(df.head().to_string())

        # Remove the temp file
        os.remove(temp_file)

        return df

    except Exception as e:
        logging.error(f"Failed to read Excel file: {e}")
        return None

    finally:
        try:
            if wb:
                wb.Close(False)
            if excel_app:
                excel_app.Quit()
        except:
            pass

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

def embed_images_in_signature(signature_html, sig_dir):
    """
    Parses the signature HTML, embeds images as attachments, and updates the HTML to reference the images via cid.

    Args:
        signature_html (str): The HTML content of the signature.
        sig_dir (str): The directory where signature images are stored.

    Returns:
        tuple: (updated_signature_html, list_of_image_attachments)
    """
    soup = BeautifulSoup(signature_html, 'html.parser')
    images = soup.find_all('img')
    image_attachments = []
    
    for img in images:
        src = img.get('src')
        if not src:
            continue
        # Handle relative paths
        img_path = os.path.join(sig_dir, src)
        if not os.path.isfile(img_path):
            logging.warning(f"Signature image not found: {img_path}")
            continue
        # Generate a unique CID
        cid = os.path.basename(img_path).replace(' ', '_')
        img['src'] = f"cid:{cid}"
        image_attachments.append((img_path, cid))
    
    return str(soup), image_attachments

def get_relevant_birthdays(demographics_path, password, date_list):
    """
    Reads the password-protected Excel file and returns a list of employees whose birthday matches any date in date_list.
    
    Parameters:
        demographics_path (str): Path to the Excel file.
        password (str): Password for the Excel file.
        date_list (list of tuple): List of (month, day) tuples to search for birthdays.
        
    Returns:
        list of dict: List containing employee records with matching birthdays.
    """
    try:
        # Read the Excel file using the helper function
        df = read_password_protected_excel(demographics_path, password, sheet_name='Contractor_Employee')
        if df is None:
            return []
    except Exception as e:
        logging.error(f"Error during reading Excel file: {e}")
        return []
    
    # Ensure necessary columns exist
    required_columns = ['Last, First M', 'DOB (MM/DD/YYYY)', 'Termination date', 'e-mail address']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        logging.error(f"Missing required columns: {missing_columns}")
        return []
    
    # Exclude terminated employees (Termination date not blank)
    active_df = df[df['Termination date'].isna()].copy()
    
    # Convert DOB to datetime
    active_df.loc[:, 'DOB'] = pd.to_datetime(active_df['DOB (MM/DD/YYYY)'], errors='coerce')
    
    # Drop rows with invalid DOB
    active_df = active_df.dropna(subset=['DOB'])
    
    # Extract month and day from DOB
    active_df.loc[:, 'DOB_Month'] = active_df['DOB'].dt.month
    active_df.loc[:, 'DOB_Day'] = active_df['DOB'].dt.day
    
    # Create a DataFrame from date_list for efficient filtering
    date_df = pd.DataFrame(date_list, columns=['DOB_Month', 'DOB_Day'])
    
    # Merge to find matching birthdays
    birthday_df = active_df.merge(date_df, on=['DOB_Month', 'DOB_Day'], how='inner')
    
    # Drop entries with missing email addresses
    birthday_df = birthday_df.dropna(subset=['e-mail address'])
    
    # Convert to list of dictionaries
    birthdays = birthday_df.to_dict(orient='records')
    
    return birthdays

def create_birthday_image(employee, presentation):
    """
    Updates a random slide with the employee's information and exports it as an image.
    
    Parameters:
        employee (dict): Employee information.
        presentation (win32com.client.CDispatch): Open PowerPoint presentation object.
        
    Returns:
        str or None: Path to the exported image, or None if failed.
    """
    try:
        if presentation.Slides.Count == 0:
            logging.error("The presentation does not contain any slides.")
            return None
        
        # Select a random slide
        random_slide_number = random.randint(1, presentation.Slides.Count)
        slide = presentation.Slides(random_slide_number)
        
        # Update shapes with employee information
        slide.Shapes("EmployeeName").TextFrame.TextRange.Text = employee['Last, First M']
        
        # Calculate age
        dob = employee['DOB']
        today = datetime.datetime.today()
        age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
        slide.Shapes("Age").TextFrame.TextRange.Text = str(age)
        
        # Format BirthDate as "Month Day" (e.g., January 01)
        birth_date_formatted = dob.strftime("%B %d")
        slide.Shapes("BirthDate").TextFrame.TextRange.Text = birth_date_formatted
        
        # Export slide as GIF image
        temp_dir = tempfile.gettempdir()
        sanitized_name = employee['Last, First M'].replace(',', '').replace(' ', '_')
        temp_image_path = os.path.join(temp_dir, f"BirthdaySlide_{sanitized_name}.gif")
        slide.Export(temp_image_path, "GIF")
        
        logging.info(f"Birthday image created at: {temp_image_path}")
        return temp_image_path
    except Exception as e:
        logging.error(f"Error creating birthday image for {employee['Last, First M']}: {e}")
        return None

def compose_email(employee, image_path, signature_html, signature_images):
    """
    Composes and sends an email to the employee with the birthday image and signature embedded.

    Parameters:
        employee (dict): Employee information.
        image_path (str): Path to the birthday image to embed in the email.
        signature_html (str): HTML content of the user's signature.
        signature_images (list): List of tuples containing signature image paths and their CIDs.
    """
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = employee['e-mail address']
        mail.CC = "ulyana.stokolosa@absolutecaregivers.com"
        
        # Extract first name for the subject
        name_parts = employee['Last, First M'].split(',')
        first_name = name_parts[1].strip() if len(name_parts) > 1 else employee['Last, First M']
        mail.Subject = f"Happy Birthday {first_name}!!"
        
        # Initialize HTML body with the birthday image
        birthday_image_cid = f"birthday_image_{random.randint(1000,9999)}@example.com"
        html_body = f"""
        <html>
            <body>
                <center><img src="cid:{birthday_image_cid}" height="576" width="768"></center>
                <br>
        """

        # Append the signature HTML
        html_body += signature_html
        html_body += """
            </body>
        </html>
        """
        
        # Add the birthday image as an attachment and set it as inline
        mail.Attachments.Add(image_path)
        attachment = mail.Attachments.Item(mail.Attachments.Count)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", birthday_image_cid)
        
        # Add signature images as attachments and set them as inline
        for img_path, cid in signature_images:
            mail.Attachments.Add(img_path)
            sig_attachment = mail.Attachments.Item(mail.Attachments.Count)
            sig_attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
        
        # Set the HTML body
        mail.HTMLBody = html_body
        
        mail.Display()  # Use .Send() to send automatically
        # Uncomment the line below to send the email without displaying it
        # mail.Send()
        
        logging.info(f"Email composed for {employee['Last, First M']} at {employee['e-mail address']}")
    except Exception as e:
        logging.error(f"Error sending email to {employee['Last, First M']}: {e}")
        raise e  # Re-raise the exception to handle in the calling function

def send_birthday_email(employee, image_path, signature_html, signature_images):
    """
    Attempts to send an email via COM automation with a timeout.
    If it fails or times out, logs an error message.

    Parameters:
        employee (dict): Employee information.
        image_path (str): Path to the birthday image to embed in the email.
        signature_html (str): HTML content of the user's signature.
        signature_images (list): List of tuples containing signature image paths and their CIDs.
    """
    def compose_email_process():
        compose_email(employee, image_path, signature_html, signature_images)

    try:
        process = Process(target=compose_email_process)
        process.start()
        process.join(timeout=5)  # Wait up to 5 seconds

        if process.is_alive():
            logging.error(f"Composing email for {employee['Last, First M']} took too long, terminating process.")
            process.terminate()
            process.join()
            # Since attachments are involved, 'mailto' is not suitable as a fallback.
            # Log the error and proceed to the next employee.
            print(f"Failed to send email to {employee['Last, First M']} due to timeout.")
        else:
            logging.info(f"Email composed successfully for {employee['Last, First M']}")
    except Exception as e:
        logging.error(f"Exception during composing email for {employee['Last, First M']}: {e}")
        print(f"Failed to send email to {employee['Last, First M']} due to an error.")


def main():
    """
    Main function to execute the birthday notification process.
    """
    # ====== Configuration ======
    # Set the password for the Excel file here
    EXCEL_PASSWORD = "abs$1004$N"  
    # ============================

    username = get_current_username()
    if not username:
        return

    base_path = f"C:\\Users\\{username}\\OneDrive - Ability Home Health, LLC\\"
    demographics_path = find_employee_demographics_file(base_path)
    if not demographics_path:
        logging.error("Employee Demographics file not found.")
        print("Employee Demographics file not found.")
        return

    logging.info(f"Employee Demographics file found at: {demographics_path}")
    print(f"Employee Demographics file found at: {demographics_path}")


    # ====== Determine Relevant Dates ======
    today = datetime.datetime.today()
    weekday = today.weekday()  # Monday is 0 and Sunday is 6
    date_list = []

    if weekday == 0:  # If today is Monday
        # Calculate dates for Saturday and Sunday
        saturday = today - datetime.timedelta(days=2)
        sunday = today - datetime.timedelta(days=1)
        monday = today
        date_list = [
            (saturday.month, saturday.day),
            (sunday.month, sunday.day),
            (monday.month, monday.day)
        ]
        logging.info("Today is Monday. Including birthdays from Saturday, Sunday, and Monday.")
        print("Today is Monday. Including birthdays from Saturday, Sunday, and Monday.")
    else:
        # Include only today's date
        date_list = [
            (today.month, today.day)
        ]
        logging.info("Today is not Monday. Including only today's birthdays.")
        print("Today is not Monday. Including only today's birthdays.")
    # ======================================

    birthdays = get_relevant_birthdays(demographics_path, EXCEL_PASSWORD, date_list)
    if not birthdays:
        logging.info("No birthdays found for the specified dates.")
        print("No birthdays found for the specified dates.")
        return

    # ====== Locate the PowerPoint template using hardcoded relative path ======
    # Assuming the script is located in 'AbolsuteAutoSciptApp\daily_tasks\birthday.py'
    # and the PPT is in 'AbolsuteAutoSciptApp\resources\Birthday_PPT.pptx'
    ppt_path = get_resource_path(
        os.path.join("..", "resources", "Birthday_PPT.pptx")
    )  # Relative path up one level

    # For debugging, print the constructed path
    print(f"Looking for PowerPoint template at: {ppt_path}")

    if not os.path.exists(ppt_path):
        logging.error(f"PowerPoint file does not exist at: {ppt_path}")
        print(f"PowerPoint file does not exist at: {ppt_path}")
        return
    else:
        logging.info(f"PowerPoint file found at: {ppt_path}")
        print(f"PowerPoint file found at: {ppt_path}")
    # ======================================================================

    try:
        ppt_app = win32.DispatchEx("PowerPoint.Application")
        # ppt_app.Visible = False  # Removed this line to prevent errors
        presentation = ppt_app.Presentations.Open(ppt_path, WithWindow=False)
        ppt_app.WindowState = 2  # Minimize PowerPoint window
    except Exception as e:
        logging.error(f"Failed to open PowerPoint presentation: {e}")
        print(f"Failed to open PowerPoint presentation: {e}")
        return

    # Retrieve the user's signature
    signature_html = get_default_signature()
    if signature_html:
        # Define the signature directory
        appdata = os.environ.get('APPDATA')
        sig_dir = os.path.join(appdata, 'Microsoft', 'Signatures')
        # Parse and embed signature images
        signature_html, signature_images = embed_images_in_signature(signature_html, sig_dir)
    else:
        signature_images = []
        signature_html = ""

    for employee in birthdays:
        logging.info(f"Processing birthday for {employee['Last, First M']}")
        print(f"\nProcessing birthday for {employee['Last, First M']}")
        image_path = create_birthday_image(employee, presentation)
        if image_path and os.path.exists(image_path):
            send_birthday_email(employee, image_path, signature_html, signature_images)
            try:
                os.remove(image_path)
                logging.info(f"Temporary image {image_path} deleted.")
                print(f"Temporary image {image_path} deleted.")
            except Exception as e:
                logging.error(f"Failed to delete temporary image {image_path}: {e}")
                print(f"Failed to delete temporary image {image_path}: {e}")
        else:
            logging.error(f"Failed to create image for {employee['Last, First M']}")
            print(f"Failed to create image for {employee['Last, First M']}")

    # Close PowerPoint after processing all employees
    try:
        presentation.Close()
        ppt_app.Quit()
        logging.info("PowerPoint closed successfully.")
        print("\nPowerPoint closed successfully.")
    except Exception as e:
        logging.error(f"Failed to close PowerPoint: {e}")
        print(f"Failed to close PowerPoint: {e}")

    logging.info("Birthday notifications completed.")
    print("Birthday notifications completed.")

if __name__ == "__main__":
    main()
