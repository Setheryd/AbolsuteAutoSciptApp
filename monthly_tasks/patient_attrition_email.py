import os
import sys
import matplotlib.pyplot as plt
import win32com.client as win32
from datetime import datetime
from patient_attrition import ChurnAttritionAnalyzer
from patient_data_extractor import PatientDataExtractor
from bs4 import BeautifulSoup  # Needed for signature parsing
import logging
from multiprocessing import Process
import urllib.parse
import webbrowser

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
print("Logging configured.")


def get_resource_path(relative_path):
    """Get the absolute path to the resource, works for PyInstaller executable."""
    print(f"get_resource_path called with relative_path: {relative_path}")
    try:
        # PyInstaller creates a temporary folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
        print(f"Using PyInstaller base path: {base_path}")
    except AttributeError:
        # If not running as an executable, use the current script directory
        base_path = os.path.dirname(os.path.abspath(__file__))
        print(f"Using script directory as base path: {base_path}")

    full_path = os.path.join(base_path, relative_path)
    print(f"Full path resolved: {full_path}")
    return full_path


# Get the parent directory using get_resource_path
parent_dir = get_resource_path(os.path.join(os.pardir))
print(f"Parent directory: {parent_dir}")

# Add the data_extraction directory to the system path
sys.path.append(os.path.join(parent_dir, "data_extraction"))
print(f"Added data_extraction to sys.path: {os.path.join(parent_dir, 'data_extraction')}")

def save_report_and_chart(analyzer):
    print("save_report_and_chart called.")
    # Run the analysis
    print("Loading data...")
    df = analyzer.load_data()
    if df is None:
        print("No data available.")
        return None
    print("Data loaded.")

    # Generate all monthly reports
    print("Generating all monthly reports...")
    report_df = analyzer.generate_all_monthly_reports(df)
    print("All monthly reports generated.")

    # Get last month's report as a dictionary
    print("Generating last month's report...")
    last_month_report = analyzer.generate_monthly_report(df)
    print("Last month's report generated.")

    # Generate the chart and save it as a PNG file
    print("Generating charts...")
    chart_filename = analyzer.generate_charts(report_df)  # Returns absolute path
    print(f"Charts generated and saved to {chart_filename}")

    # Verify that the chart was saved successfully
    if not os.path.exists(chart_filename):
        logging.error(f"Chart was not saved correctly at: {chart_filename}")
        return None

    print("Chart file verified to exist.")
    return {
        "report_month": last_month_report["Report Month"],
        "starting_patient_count": last_month_report["Starting Patient Count"],
        "ending_patient_count": last_month_report["Ending Patient Count"],
        "new_patients": last_month_report["New Patients"],
        "discharged_patients": last_month_report["Discharged Patients"],
        "net_change": last_month_report["Net Change"],
        "churn_rate": last_month_report["Churn Rate (%)"],
        "attrition_rate": last_month_report["Attrition Rate (%)"],
        "chart_filename": chart_filename,
    }

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
            print(f"Found {accounts.Count} Outlook account(s).")
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

def get_signature_by_path(sig_path):
    """ Retrieves the email signature from the specified file path. """
    print(f"get_signature_by_path called with sig_path: {sig_path}")
    try:
        with open(sig_path, 'r', encoding='utf-8') as file:
            signature = file.read()
            print(f"Signature retrieved from {sig_path}")
        return signature
    except Exception as e:
        logging.error(f"Unable to retrieve signature from {sig_path}: {e}")
        return None

def get_default_signature():
    """ Retrieves the user's default email signature. """
    print("get_default_signature called.")
    email = get_default_outlook_email()
    if not email:
        print("No default Outlook email found.")
        return None
    print(f"Default Outlook email: {email}")

    appdata = os.environ.get('APPDATA')
    if not appdata:
        print("APPDATA environment variable not found.")
        return None
    print(f"APPDATA directory: {appdata}")

    sig_dir = os.path.join(appdata, 'Microsoft', 'Signatures')
    print(f"Signature directory: {sig_dir}")
    if not os.path.isdir(sig_dir):
        print("Signature directory does not exist.")
        return None

    print("Looking for signature files...")
    for filename in os.listdir(sig_dir):
        if filename.lower().endswith(('.htm', '.html')):
            base_name = os.path.splitext(filename)[0].lower()
            if email.lower() in base_name:
                sig_path = os.path.join(sig_dir, filename)
                print(f"Signature file found: {sig_path}")
                return get_signature_by_path(sig_path)
    print("No matching signature file found.")
    return None

def embed_images_in_signature(signature_html, sig_dir):
    """ Embeds images as attachments and updates the signature HTML to reference them via CID. """
    print("embed_images_in_signature called.")
    soup = BeautifulSoup(signature_html, 'html.parser')
    images = soup.find_all('img')
    image_attachments = []
    print(f"Found {len(images)} image(s) in signature.")

    for img in images:
        src = img.get('src')
        if not src:
            print("Image with no src attribute found, skipping.")
            continue
        img_path = os.path.join(sig_dir, src)
        if not os.path.isfile(img_path):
            print(f"Image file not found: {img_path}, skipping.")
            continue
        cid = os.path.basename(img_path).replace(' ', '_')
        img['src'] = f"cid:{cid}"
        image_attachments.append((img_path, cid))
        print(f"Embedded image: {img_path} as CID: {cid}")

    print("All images processed.")
    return str(soup), image_attachments

def compose_email_classic(report, signature, chart_filename):
    """
    Composes and displays an email via COM automation for classic Outlook.
    """
    try:
        print("compose_email_classic called.")
        # Create the Outlook application
        print("Creating Outlook application...")
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        print("Outlook email item created.")

        # Email body with last month's results
        print("Preparing email body...")
        body = f"""
        <html>
            <body>
                <p>Hello,</p>
                <p>Please find attached the churn and attrition analysis report for last month ({report['report_month']}):</p>
                <ul>
                    <li>Starting Patient Count: {report['starting_patient_count']}</li>
                    <li>Ending Patient Count: {report['ending_patient_count']}</li>
                    <li>New Patients: {report['new_patients']}</li>
                    <li>Discharged Patients: {report['discharged_patients']}</li>
                    <li>Net Change: {report['net_change']}</li>
                    <li>Churn Rate: {report['churn_rate']}%</li>
                    <li>Attrition Rate: {report['attrition_rate']}%</li>
                </ul>
                <p>See the chart below for a visual representation:</p>
                <img src="cid:chart_image">
                <p>Best regards,</p>
                {signature}
            </body>
        </html>
        """
        print("Email body prepared.")

        # Set the HTML body
        mail.HTMLBody = body
        print("HTML body set.")

        # Embed the chart image
        if chart_filename and os.path.exists(chart_filename):
            print(f"Attaching chart image: {chart_filename}")
            attachment = mail.Attachments.Add(chart_filename)
            # Assign a Content-ID to the attachment
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "chart_image")
            print("Chart image attached and Content-ID set.")
        else:
            logging.error(f"Chart image not found at path: {chart_filename}")

        # Set email parameters
        mail.To = "alexander.nazarov@absolutecaregivers.com"
        mail.CC = "luke.kitchel@absolutecaregivers.com; seth.riley@absolutecaregivers.com"
        mail.Subject = 'Patient Monthly Churn and Attrition Report'
        print("Email recipients and subject set.")
        mail.Display()  # Use .Send() to send it directly

        print("Email prepared successfully.")
    except Exception as e:
        logging.error(f"Failed to compose or display email via COM automation: {e}")
        raise

def send_email(report, signature, chart_filename):
    print("send_email called.")
    # Try composing email via COM automation with a timeout
    try:
        process = Process(target=compose_email_classic, args=(report, signature, chart_filename))
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

    # Prepare email components
    to_addresses = "alexander.nazarov@absolutecaregivers.com"
    cc_addresses = "luke.kitchel@absolutecaregivers.com; seth.riley@absolutecaregivers.com"
    subject = 'Patient Monthly Churn and Attrition Report'

    # Convert HTML content to plain text
    body_text = (
        f"Hello,\n\n"
        f"Please find attached the churn and attrition analysis report for last month ({report['report_month']}):\n\n"
        f"Starting Patient Count: {report['starting_patient_count']}\n"
        f"Ending Patient Count: {report['ending_patient_count']}\n"
        f"New Patients: {report['new_patients']}\n"
        f"Discharged Patients: {report['discharged_patients']}\n"
        f"Net Change: {report['net_change']}\n"
        f"Churn Rate: {report['churn_rate']}%\n"
        f"Attrition Rate: {report['attrition_rate']}%\n\n"
        "Best regards,\n"
    )

    # Add signature if available
    if signature:
        # Remove HTML tags from signature
        soup = BeautifulSoup(signature, 'html.parser')
        signature_text = soup.get_text()
        body_text += signature_text
    else:
        body_text += "[Your Name]\n[Your Position]"

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

def main():
    print("main function called.")
    # Initialize the data extractor and analyzer
    print("Initializing data extractor...")
    extractor = PatientDataExtractor()
    print("Data extractor initialized.")
    print("Initializing analyzer...")
    analyzer = ChurnAttritionAnalyzer(extractor)
    print("Analyzer initialized.")

    # Run the analysis and get the report and chart path
    print("Running analysis and generating report...")
    report = save_report_and_chart(analyzer)

    if report is None:
        print("No data to send in the email.")
        return
    print("Report and chart generated.")

    # Retrieve the default signature or use a fallback
    print("Retrieving default signature...")
    signature = get_default_signature() or "Best regards,<br>Your Name<br>Your Position"
    print("Signature retrieved.")

    # Send the email with the report and embedded chart
    print("Sending email...")
    send_email(report, signature, report["chart_filename"])
    print("Email sent.")

if __name__ == "__main__":
    print("Script started.")
    main()
    print("Script finished.")
