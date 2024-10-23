import os
import sys
import matplotlib.pyplot as plt
import win32com.client as win32
from datetime import datetime
from employee_attrition import ChurnAttritionAnalyzer
from caregiver_data_extractor import CaregiverDataExtractor
from bs4 import BeautifulSoup  # Needed for signature parsing
import logging

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
        "starting_employee_count": last_month_report["Starting Contractor Count"],  # Updated to match the report
        "ending_employee_count": last_month_report["Ending Contractor Count"],      # Updated to match the report
        "new_employees": last_month_report["New Contractors"],                      # Updated to match the report
        "discharged_employees": last_month_report["Terminated Contractors"],        # Updated to match the report
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
        outlook = win32.DispatchEx("Outlook.Application")
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

def send_email(report, signature, chart_filename):
    print("send_email called.")
    # Create the Outlook application
    print("Creating Outlook application...")
    outlook = win32.DispatchEx('outlook.application')
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
                <li>Starting employee Count: {report['starting_employee_count']}</li>
                <li>Ending employee Count: {report['ending_employee_count']}</li>
                <li>New employees: {report['new_employees']}</li>
                <li>Discharged employees: {report['discharged_employees']}</li>
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
    mail.Subject = "Employee Monthly Churn and Attrition Report"
    print("Email recipients and subject set.")
    mail.Display()  # Use .Send() to send it directly

    print("Email prepared successfully.")

def main():
    print("main function called.")
    # Initialize the data extractor and analyzer
    print("Initializing data extractor...")
    extractor = CaregiverDataExtractor()
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
