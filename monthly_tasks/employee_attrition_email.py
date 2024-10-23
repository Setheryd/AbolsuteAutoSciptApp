import os
import sys
import matplotlib.pyplot as plt
import win32com.client as win32
from datetime import datetime
from .employee_attrition import ChurnAttritionAnalyzer
from data_extraction.caregiver_data_extractor import CaregiverDataExtractor
from bs4 import BeautifulSoup  # Needed for signature parsing
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def get_resource_path(relative_path):
    """Get the absolute path to the resource, works for PyInstaller executable."""
    try:
        # PyInstaller creates a temporary folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # If not running as an executable, use the current script directory
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)


# Get the parent directory using get_resource_path
parent_dir = get_resource_path(os.path.join(os.pardir))

# Add the data_extraction directory to the system path
sys.path.append(os.path.join(parent_dir, "data_extraction"))


def save_report_and_chart(analyzer):
    # Run the analysis
    df = analyzer.load_data()
    if df is None:
        print("No data available.")
        return None

    # Generate all monthly reports
    report_df = analyzer.generate_all_monthly_reports(df)

    # Get last month's report as a dictionary
    last_month_report = analyzer.generate_monthly_report(df)

    # Generate the chart and save it as a PNG file
    chart_filename = analyzer.generate_charts(report_df)  # Returns absolute path

    # Verify that the chart was saved successfully
    if not os.path.exists(chart_filename):
        logging.error(f"Chart was not saved correctly at: {chart_filename}")
        return None

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

def get_signature_by_path(sig_path):
    """ Retrieves the email signature from the specified file path. """
    try:
        with open(sig_path, 'r', encoding='utf-8') as file:
            signature = file.read()
        return signature
    except Exception as e:
        logging.error(f"Unable to retrieve signature from {sig_path}: {e}")
        return None

def get_default_signature():
    """ Retrieves the user's default email signature. """
    email = get_default_outlook_email()
    if not email:
        return None

    appdata = os.environ.get('APPDATA')
    if not appdata:
        return None

    sig_dir = os.path.join(appdata, 'Microsoft', 'Signatures')
    if not os.path.isdir(sig_dir):
        return None

    for filename in os.listdir(sig_dir):
        if filename.lower().endswith(('.htm', '.html')):
            base_name = os.path.splitext(filename)[0].lower()
            if email.lower() in base_name:
                sig_path = os.path.join(sig_dir, filename)
                return get_signature_by_path(sig_path)

    return None

def embed_images_in_signature(signature_html, sig_dir):
    """ Embeds images as attachments and updates the signature HTML to reference them via CID. """
    soup = BeautifulSoup(signature_html, 'html.parser')
    images = soup.find_all('img')
    image_attachments = []
    
    for img in images:
        src = img.get('src')
        if not src:
            continue
        img_path = os.path.join(sig_dir, src)
        if not os.path.isfile(img_path):
            continue
        cid = os.path.basename(img_path).replace(' ', '_')
        img['src'] = f"cid:{cid}"
        image_attachments.append((img_path, cid))
    
    return str(soup), image_attachments

def send_email(report, signature, chart_filename):
    # Create the Outlook application
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    # Email body with last month's results
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

    # Set the HTML body
    mail.HTMLBody = body

    # Embed the chart image
    if chart_filename and os.path.exists(chart_filename):
        attachment = mail.Attachments.Add(chart_filename)
        # Assign a Content-ID to the attachment
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "chart_image")
    else:
        logging.error(f"Chart image not found at path: {chart_filename}")

    # Set email parameters
    mail.To = "alexander.nazarov@absolutecaregivers.com"
    mail.CC = "luke.kitchel@absolutecaregivers.com; seth.riley@absolutecaregivers.com"
    mail.Subject = "Employee Monthly Churn and Attrition Report"
    mail.Display()  # Use .Send() to send it directly

    print("Email prepared successfully.")

def main():
    # Initialize the data extractor and analyzer
    extractor = CaregiverDataExtractor()
    analyzer = ChurnAttritionAnalyzer(extractor)
    
    # Run the analysis and get the report and chart path
    report = save_report_and_chart(analyzer)
    
    if report is None:
        print("No data to send in the email.")
        return
    
    # Retrieve the default signature or use a fallback
    signature = get_default_signature() or "Best regards,<br>Your Name<br>Your Position"

    # Send the email with the report and embedded chart
    send_email(report, signature, report["chart_filename"])

if __name__ == "__main__":
    main()
