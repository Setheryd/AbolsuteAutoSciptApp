import os
import sys
import matplotlib.pyplot as plt
import win32com.client as win32
from datetime import datetime
from patient_attrition import ChurnAttritionAnalyzer
from patient_data_extractor import PatientDataExtractor

# Get the parent directory of the current script
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, os.pardir))

# Add the data_extraction directory to the system path
sys.path.append(os.path.join(parent_dir, "data_extraction"))

def save_report_and_chart(analyzer):
    # Run the analysis
    df = analyzer.load_data()
    if df is None:
        print("No data available.")
        return None, None

    # Generate all monthly reports
    report_df = analyzer.generate_all_monthly_reports(df)

    # Get last month's report
    last_month_report = analyzer.generate_monthly_report(df)

    # Generate the chart and save it as a PNG file
    chart_filename = "churn_attrition_chart.png"
    analyzer.generate_charts(report_df)
    plt.savefig(chart_filename)  # Save the chart

    # Format last month's results as a string
    report_str = "\n".join([f"{k}: {v}" for k, v in last_month_report.items()])

    return report_str, chart_filename

def send_email(report_str, chart_filename):
    # Create the Outlook application
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    # Set email attributes
    mail.Subject = 'Monthly Churn and Attrition Analysis Report'
    mail.Body = f"Hello,\n\nPlease find attached the churn and attrition analysis report for last month:\n\n{report_str}\n\nBest regards."
    
    # Attach the chart image
    if chart_filename and os.path.exists(chart_filename):
        mail.Attachments.Add(os.path.join(os.getcwd(), chart_filename))

    # Specify the recipient and send the email
    mail.To = 'recipient@example.com'  # Replace with the actual recipient's email
    mail.Display()

def main():
    # Create the analyzer
    extractor = PatientDataExtractor()
    analyzer = ChurnAttritionAnalyzer(extractor)
    
    # Save the report and chart
    report_str, chart_filename = save_report_and_chart(analyzer)
    
    if report_str and chart_filename:
        # Send the email
        send_email(report_str, chart_filename)
        print("Email sent successfully with the report and chart.")
    else:
        print("Failed to generate the report or chart.")

if __name__ == "__main__":
    main()
